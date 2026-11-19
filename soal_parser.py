"""
soal_parser.py — Membaca naskah soal DOCX (format SMPN 50 Surabaya) menjadi data terstruktur
untuk dibuat menjadi Google Form.

Dikenali:
- Judul bagian berhuruf kapital: SOAL PILIHAN GANDA / PILIHAN GANDA KOMPLEK(S) / BENAR SALAH /
  ISIAN (SINGKAT) / MENJODOHKAN / URAIAN (ESAI). Judul ini sekaligus menentukan jenis soal.
- Kalimat perintah tepat di bawah judul -> petunjuk bagian.
- "Petunjuk Umum" sebelum bagian pertama -> deskripsi Form.
- Nomor soal diketik ("12.") atau penomoran otomatis Word; bila tidak ada nomor sama sekali,
  batas soal ditentukan dari paragraf kosong / soal baru setelah opsi.
- Opsi: "A." diketik, huruf otomatis Word, kotak centang "[ ]"/"☐", dua kolom dengan Tab,
  atau di dalam tabel tata letak.
- Benar/Salah & Menjodohkan dari tabel (setiap baris = satu nomor).
- Persamaan Word (OMML) dan huruf pangkat/indeks diubah menjadi teks biasa (9/16, Iᴬ, H₂O).
- Gambar di dalam soal (gambar di kop/header dan tanda air diabaikan).
- Baris opsional "Kunci: B" dan "Skor: 2" di bawah soal.
"""
from __future__ import annotations

import io
import re
import zipfile
from typing import Any, Dict, List, Optional, Tuple

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]
R = "{%s}" % NS["r"]


def q(tag: str) -> str:
    p, t = tag.split(":")
    return "{%s}%s" % (NS[p], t)


# ------------------------------------------------------------------------------
# Pangkat / indeks -> karakter Unicode
# ------------------------------------------------------------------------------
SUP = dict(zip("0123456789+-=()nibcdefghjklmoprstuvwxyzaABDEGHIJKLMNOPRTUVW",
               "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵇᶜᵈᵉᶠᵍʰʲᵏˡᵐᵒᵖʳˢᵗᵘᵛʷˣʸᶻᵃᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂ"))
SUB = dict(zip("0123456789+-=()aehijklmnoprstuvx",
               "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ"))


def to_sup(s: str) -> str:
    s = s.strip()
    if s and all(c in SUP for c in s):
        return "".join(SUP[c] for c in s)
    return f"^({s})" if len(s) > 1 else f"^{s}"


def to_sub(s: str) -> str:
    s = s.strip()
    if s and all(c in SUB for c in s):
        return "".join(SUB[c] for c in s)
    return f"_({s})" if len(s) > 1 else f"_{s}"


# ------------------------------------------------------------------------------
# Persamaan Word (OMML) -> teks
# ------------------------------------------------------------------------------
def _wrap(s: str) -> str:
    s = s.strip()
    simple = re.fullmatch(r"\d+(?:[.,]\d+)?", s) or re.fullmatch(r"\w[⁰-⁹²³¹ⁿᴬ-ᵂᵃ-ᶻ₀-₉]*", s)
    return s if simple else f"({s})"


def math_to_text(el) -> str:
    tag = etree.QName(el).localname if isinstance(el.tag, str) else ""
    kids = lambda name: el.find(q("m:" + name))
    txt = lambda child: math_to_text(child) if child is not None else ""

    if tag == "t":
        return el.text or ""
    if tag == "f":
        return f"{_wrap(txt(kids('num')))}/{_wrap(txt(kids('den')))}"
    if tag == "sSup":
        return txt(kids("e")) + to_sup(txt(kids("sup")))
    if tag == "sSub":
        return txt(kids("e")) + to_sub(txt(kids("sub")))
    if tag == "sSubSup":
        return txt(kids("e")) + to_sub(txt(kids("sub"))) + to_sup(txt(kids("sup")))
    if tag == "rad":
        deg = txt(kids("deg")).strip()
        return (to_sup(deg) if deg else "") + "√" + _wrap(txt(kids("e")))
    if tag == "d":
        pr = el.find(q("m:dPr"))
        beg, end = "(", ")"
        if pr is not None:
            b = pr.find(q("m:begChr")); e = pr.find(q("m:endChr"))
            if b is not None: beg = b.get(q("m:val"), "(")
            if e is not None: end = e.get(q("m:val"), ")")
        inner = ", ".join(math_to_text(c) for c in el.findall(q("m:e")))
        return f"{beg}{inner}{end}"
    if tag == "nary":
        pr = el.find(q("m:naryPr")); ch = "∫"
        if pr is not None and pr.find(q("m:chr")) is not None:
            ch = pr.find(q("m:chr")).get(q("m:val"), "∫")
        return ch + to_sub(txt(kids("sub"))) + to_sup(txt(kids("sup"))) + txt(kids("e"))
    if tag.endswith("Pr"):
        return ""
    return "".join(math_to_text(c) for c in el)


# ------------------------------------------------------------------------------
# Paket DOCX
# ------------------------------------------------------------------------------
class DocxPackage:
    def __init__(self, data: bytes):
        self.zip = zipfile.ZipFile(io.BytesIO(data))
        self.doc = etree.fromstring(self.zip.read("word/document.xml"))
        self.rels = self._rels("word/_rels/document.xml.rels")
        self.numfmt = self._numbering()
        self.style_num = self._style_numbering()

    def _rels(self, path: str) -> Dict[str, str]:
        try:
            x = etree.fromstring(self.zip.read(path))
        except KeyError:
            return {}
        out = {}
        for r in x:
            tgt = r.get("Target", "")
            out[r.get("Id")] = tgt if tgt.startswith("/") else "word/" + tgt.lstrip("./")
        return out

    def _numbering(self) -> Dict[Tuple[str, str], str]:
        try:
            x = etree.fromstring(self.zip.read("word/numbering.xml"))
        except KeyError:
            return {}
        abstract = {}
        for a in x.findall(q("w:abstractNum")):
            lv = {}
            for l in a.findall(q("w:lvl")):
                f = l.find(q("w:numFmt"))
                lv[l.get(q("w:ilvl"))] = f.get(q("w:val")) if f is not None else "decimal"
            abstract[a.get(q("w:abstractNumId"))] = lv
        out = {}
        for n in x.findall(q("w:num")):
            aid = n.find(q("w:abstractNumId"))
            if aid is None:
                continue
            for ilvl, fmt in abstract.get(aid.get(q("w:val")), {}).items():
                out[(n.get(q("w:numId")), ilvl)] = fmt
        return out

    def _style_numbering(self) -> Dict[str, Tuple[str, str]]:
        try:
            x = etree.fromstring(self.zip.read("word/styles.xml"))
        except KeyError:
            return {}
        out = {}
        for s in x.findall(q("w:style")):
            np_ = s.find(".//" + q("w:numPr"))
            if np_ is not None and np_.find(q("w:numId")) is not None:
                ilvl = np_.find(q("w:ilvl"))
                out[s.get(q("w:styleId"))] = (np_.find(q("w:numId")).get(q("w:val")),
                                              ilvl.get(q("w:val")) if ilvl is not None else "0")
        return out

    def media(self, rid: str) -> Optional[bytes]:
        path = self.rels.get(rid)
        if not path:
            return None
        try:
            return self.zip.read(path)
        except KeyError:
            return None


# ------------------------------------------------------------------------------
# Paragraf
# ------------------------------------------------------------------------------
class Para:
    def __init__(self, pkg: DocxPackage, p):
        self.pkg = pkg
        self.text = self._text(p).replace(" ", " ")
        self.images: List[bytes] = self._images(p)
        self.numfmt: Optional[str] = None
        self.numid: Optional[str] = None
        ppr = p.find(q("w:pPr"))
        if ppr is not None:
            np_ = ppr.find(q("w:numPr"))
            numid = ilvl = None
            if np_ is not None and np_.find(q("w:numId")) is not None:
                numid = np_.find(q("w:numId")).get(q("w:val"))
                il = np_.find(q("w:ilvl"))
                ilvl = il.get(q("w:val")) if il is not None else "0"
            else:
                st = ppr.find(q("w:pStyle"))
                if st is not None and st.get(q("w:val")) in pkg.style_num:
                    numid, ilvl = pkg.style_num[st.get(q("w:val"))]
            if numid and numid != "0":
                self.numid = numid
                self.numfmt = pkg.numfmt.get((numid, ilvl), "decimal")

    # -- teks --
    def _text(self, el) -> str:
        out = []
        for c in el:
            if not isinstance(c.tag, str):
                continue
            tag = etree.QName(c).localname
            ns = etree.QName(c).namespace
            if ns == NS["m"] and tag in ("oMath", "oMathPara"):
                out.append(math_to_text(c))
            elif tag == "r":
                out.append(self._run(c))
            elif tag in ("hyperlink", "ins", "smartTag", "fldSimple", "customXml"):
                out.append(self._text(c))
            elif tag == "sdt":
                if c.find(".//" + q("w14:checkbox")) is not None:
                    out.append("☐ ")
                content = c.find(q("w:sdtContent"))
                if content is not None:
                    out.append(self._text(content))
            elif tag == "AlternateContent":
                fb = c.find(q("mc:Fallback"))
                ch = c.find(q("mc:Choice"))
                for part in (ch, fb):
                    if part is not None:
                        out.append(self._text(part))
                        break
        return "".join(out)

    def _run(self, r) -> str:
        va = r.find("./" + q("w:rPr") + "/" + q("w:vertAlign"))
        mode = va.get(q("w:val")) if va is not None else None
        buf = []
        for c in r:
            if not isinstance(c.tag, str):
                continue
            tag = etree.QName(c).localname
            if tag == "t":
                buf.append(c.text or "")
            elif tag == "tab":
                buf.append("\t")
            elif tag in ("br", "cr"):
                buf.append("\n")
            elif tag == "sym":
                ch = c.get(q("w:char"), "")
                buf.append({"F0FC": "✓", "F0FB": "✗", "F0A8": "☐", "F06F": "☐", "F0FE": "☒"}.get(ch.upper(), ""))
            elif tag == "noBreakHyphen":
                buf.append("-")
        s = "".join(buf)
        if mode == "superscript" and s.strip():
            return to_sup(s)
        if mode == "subscript" and s.strip():
            return to_sub(s)
        return s

    # -- gambar --
    def _images(self, p) -> List[bytes]:
        imgs = []
        for drawing in p.iter(q("w:drawing")):
            anchor = drawing.find(q("wp:anchor"))
            if anchor is not None and anchor.get("behindDoc") == "1":
                continue  # tanda air / gambar latar
            for blip in drawing.iter(q("a:blip")):
                rid = blip.get(q("r:embed")) or blip.get(q("r:link"))
                data = self.pkg.media(rid) if rid else None
                if data:
                    imgs.append(data)
        for im in p.iter(q("v:imagedata")):
            rid = im.get(q("r:id"))
            data = self.pkg.media(rid) if rid else None
            if data:
                imgs.append(data)
        return imgs

    @property
    def clean(self) -> str:
        return re.sub(r"[ \t]+", " ", self.text.replace("\n", " ")).strip()

    @property
    def is_blank(self) -> bool:
        return not self.clean and not self.images


class Table:
    def __init__(self, pkg: DocxPackage, tbl):
        self.rows: List[List[List[Para]]] = []
        for tr in tbl.findall(q("w:tr")):
            row = []
            for tc in tr.findall(q("w:tc")):
                row.append([Para(pkg, p) for p in tc.iter(q("w:p"))])
            self.rows.append(row)

    def cell_text(self, cell: List[Para]) -> str:
        return " ".join(p.clean for p in cell if p.clean).strip()

    def paragraphs(self) -> List[Para]:
        return [p for row in self.rows for cell in row for p in cell]


def iter_blocks(pkg: DocxPackage):
    body = pkg.doc.find(q("w:body"))

    def walk(parent):
        for c in parent:
            if not isinstance(c.tag, str):
                continue
            if c.tag == q("w:p"):
                yield Para(pkg, c)
            elif c.tag == q("w:tbl"):
                yield Table(pkg, c)
            elif c.tag == q("w:sdt"):
                content = c.find(q("w:sdtContent"))
                if content is not None:
                    yield from walk(content)
    yield from walk(body)


# ------------------------------------------------------------------------------
# Pola teks
# ------------------------------------------------------------------------------
HEADING_RE = re.compile(
    r"^(?:SOAL\s+|BAGIAN\s+[A-Z0-9]+\s*[:.\-–]?\s*)?("
    r"PILIHAN\s+GANDA\s+KOMPLEKS?|PILIHAN\s+GANDA|BENAR[\s\-/]*(?:ATAU\s+)?SALAH|"
    r"ISIAN(?:\s+SINGKAT)?|MENJODOHKAN|URAIAN|ESAI|ESSAY)\b", re.I)
INSTR_RE = re.compile(
    r"^(pilihlah|pilih|berilah|berikan|isilah|jodohkan|jawablah|kerjakan|tentukan|bacalah|"
    r"lengkapilah|tuliskan|silanglah|centanglah|lingkarilah)\b", re.I)
NUM_RE = re.compile(r"^\s*(\d{1,3})\s*[.)]\s+")
OPT_RE = re.compile(r"^\s*\(?([A-Ea-e])\s*[.)]\s*(?=\S|$)")
OPT_SPLIT_RE = re.compile(r"(?:\t+|\s{3,})(?=\(?[A-Ea-e]\s*[.)]\s)")
CHECK_RE = re.compile(r"^\s*(?:\[\s*\]|☐|□|▢|❏|◻|⬜)\s*")
KUNCI_RE = re.compile(r"^\s*kunci(?:\s+jawaban)?\s*[:：=]\s*(.+)$", re.I)
SKOR_RE = re.compile(r"^\s*(?:skor|poin|bobot|nilai)\s*[:：=]\s*(\d+(?:[.,]\d+)?)\s*$", re.I)
ANSWER_LINE_RE = re.compile(r"^\s*(?:jawab(?:an)?\s*:?)?\s*[.…_\-\s]{5,}$", re.I)

TIPE_DARI_JUDUL = [
    (re.compile(r"KOMPLEK", re.I), "pg_kompleks"),
    (re.compile(r"PILIHAN\s+GANDA", re.I), "pilihan_ganda"),
    (re.compile(r"BENAR", re.I), "benar_salah"),
    (re.compile(r"ISIAN", re.I), "isian"),
    (re.compile(r"MENJODOHKAN", re.I), "menjodohkan"),
    (re.compile(r"URAIAN|ESAI|ESSAY", re.I), "uraian"),
]
LETTER_FMTS = ("upperLetter", "lowerLetter")


def is_heading(p: Para) -> Optional[str]:
    t = p.clean
    if not t or len(t) > 70:
        return None
    letters = [c for c in t if c.isalpha()]
    if not letters or sum(c.isupper() for c in letters) / len(letters) < 0.8:
        return None
    if not HEADING_RE.match(t):
        return None
    for rx, tipe in TIPE_DARI_JUDUL:
        if rx.search(t):
            return tipe
    return None


def is_number_marker(p: Para) -> bool:
    return bool(NUM_RE.match(p.text)) or (p.numfmt is not None and p.numfmt not in LETTER_FMTS
                                           and p.numfmt not in ("bullet", "none"))


def strip_number(t: str) -> Tuple[Optional[int], str]:
    m = NUM_RE.match(t)
    if m:
        return int(m.group(1)), t[m.end():].strip()
    return None, t.strip()


# ------------------------------------------------------------------------------
# Struktur hasil
# ------------------------------------------------------------------------------
def new_question(tipe: str) -> Dict[str, Any]:
    return {"no": None, "tipe": tipe, "stem": [], "opsi": [], "gambar": [],
            "kunci_raw": None, "poin": None, "pilihan": [], "peringatan": []}


def _apply_meta(q_: Dict[str, Any], p: Para) -> bool:
    """Baris 'Kunci: ...' / 'Skor: ...'. True jika paragraf ini metadata."""
    m = KUNCI_RE.match(p.clean)
    if m:
        q_["kunci_raw"] = m.group(1).strip()
        return True
    m = SKOR_RE.match(p.clean)
    if m:
        q_["poin"] = float(m.group(1).replace(",", "."))
        return True
    return False


def split_option_paragraph(p: Para) -> List[Tuple[Optional[str], str]]:
    """Mengembalikan daftar (huruf, teks) jika paragraf berisi opsi; [] jika bukan."""
    raw = p.text.replace("\n", "\t")
    if CHECK_RE.match(raw):
        return [(None, CHECK_RE.sub("", raw).strip())]
    if p.numfmt in LETTER_FMTS:
        return [("#auto", re.sub(r"\s+", " ", raw).strip())]
    if OPT_RE.match(raw):
        parts = OPT_SPLIT_RE.split(raw)
        out = []
        for part in parts:
            m = OPT_RE.match(part)
            if m:
                out.append((m.group(1).upper(), re.sub(r"\s+", " ", part[m.end():]).strip()))
            elif out:
                out[-1] = (out[-1][0], (out[-1][1] + " " + part.strip()).strip())
        return out
    return []


# ------------------------------------------------------------------------------
# Pembaca per jenis bagian
# ------------------------------------------------------------------------------
def _flatten(blocks) -> List[Para]:
    out = []
    for b in blocks:
        if isinstance(b, Table):
            out.extend(b.paragraphs())
            out.append(Para.__new__(Para))  # pemisah kosong
            out[-1].text, out[-1].images, out[-1].numfmt, out[-1].numid = "", [], None, None
        else:
            out.append(b)
    return out


def parse_option_section(blocks, tipe: str) -> List[Dict[str, Any]]:
    paras = _flatten(blocks)
    numbered = any(is_number_marker(p) for p in paras if p.clean)
    soal: List[Dict[str, Any]] = []
    cur: Optional[Dict[str, Any]] = None
    pending_imgs: List[bytes] = []
    auto_letter: Dict[str, int] = {}

    def start(p: Optional[Para]):
        nonlocal cur, auto_letter
        cur = new_question(tipe)
        auto_letter = {}
        cur["gambar"].extend(pending_imgs)
        pending_imgs.clear()
        soal.append(cur)
        if p is not None:
            no, t = strip_number(p.text.replace("\n", " "))
            cur["no"] = no
            if t.strip():
                cur["stem"].append(re.sub(r"[ \t]+", " ", t).strip())

    for p in paras:
        if cur is not None and _apply_meta(cur, p):
            continue
        if p.is_blank:
            continue
        opts = split_option_paragraph(p) if p.clean or p.numfmt in LETTER_FMTS else []
        if opts:
            if cur is None:
                start(None)
                cur["peringatan"].append("Opsi ditemukan tanpa teks soal di atasnya.")
            for letter, t in opts:
                if letter == "#auto":
                    i = auto_letter.get(p.numid, 0)
                    auto_letter[p.numid] = i + 1
                    letter = "ABCDEFGH"[min(i, 7)]
                cur["opsi"].append({"huruf": letter, "teks": t})
            cur["gambar"].extend(p.images)
            continue

        if not p.clean:  # paragraf berisi gambar saja
            (cur["gambar"] if cur is not None and not cur["opsi"] else pending_imgs).extend(p.images)
            continue

        if numbered:
            new = cur is None or is_number_marker(p) or bool(cur["opsi"])
        else:
            new = cur is None or bool(cur["opsi"])
        if new:
            start(p)
        else:
            cur["stem"].append(strip_number(p.clean)[1])
        cur["gambar"].extend(p.images)

    for s in soal:
        letters = [o["huruf"] for o in s["opsi"]]
        if all(l is not None for l in letters) and len(set(letters)) == len(letters):
            s["opsi"].sort(key=lambda o: o["huruf"])
        for i, o in enumerate(s["opsi"]):
            o["huruf"] = "ABCDEFGH"[min(i, 7)]
    soal.extend([])
    if pending_imgs and soal:
        soal[-1]["gambar"].extend(pending_imgs)
    return soal


def parse_text_section(blocks, tipe: str) -> List[Dict[str, Any]]:
    paras = _flatten(blocks)
    numbered = any(is_number_marker(p) for p in paras if p.clean)
    soal: List[Dict[str, Any]] = []
    cur: Optional[Dict[str, Any]] = None
    closed = True

    for p in paras:
        if cur is not None and _apply_meta(cur, p):
            continue
        if p.is_blank:
            if not numbered and cur is not None and (cur["stem"] or cur["gambar"]):
                closed = True
            continue
        if p.clean and ANSWER_LINE_RE.match(p.clean):
            if not numbered:
                closed = True
            cur and cur["gambar"].extend(p.images)
            continue
        new = (cur is None) or (is_number_marker(p) if numbered else closed)
        if new and p.clean:
            cur = new_question(tipe)
            soal.append(cur)
            no, t = strip_number(p.clean)
            cur["no"] = no
            if t:
                cur["stem"].append(t)
            closed = False
        elif p.clean:
            cur["stem"].append(strip_number(p.clean)[1] if numbered else p.clean)
        if p.images:
            if cur is None:
                cur = new_question(tipe); soal.append(cur); closed = False
            cur["gambar"].extend(p.images)
    for s in soal:
        s["stem"] = [re.sub(r"\s*(?:Jawab(?:an)?\s*:)?\s*[.…_]{5,}\s*$", "", t).strip() for t in s["stem"]]
        s["stem"] = [t for t in s["stem"] if t]
    return soal


def _table_rows(tbl: Table) -> Tuple[List[str], List[List[List[Para]]]]:
    rows = tbl.rows
    if not rows:
        return [], []
    head = [tbl.cell_text(c) for c in rows[0]]
    is_header = sum(len(h) for h in head) < 45 and not any(NUM_RE.match(h + " ") and len(h) > 4 for h in head)
    return (head if is_header else []), (rows[1:] if is_header else rows)


def parse_benar_salah(blocks) -> List[Dict[str, Any]]:
    soal = []
    for b in blocks:
        if isinstance(b, Table):
            head, rows = _table_rows(b)
            hl = [h.lower() for h in head]
            ib = next((i for i, h in enumerate(hl) if h.startswith("benar")), None)
            is_ = next((i for i, h in enumerate(hl) if h.startswith("salah")), None)
            for row in rows:
                texts = [b.cell_text(c) for c in row]
                if not any(texts):
                    continue
                skip = {ib, is_}
                cand = [(i, t) for i, t in enumerate(texts) if i not in skip]
                i_st, statement = max(cand, key=lambda it: len(it[1])) if cand else (0, "")
                if not statement:
                    continue
                q_ = new_question("benar_salah")
                no_cell = next((t for i, t in cand if i != i_st and re.fullmatch(r"\d{1,3}\s*[.)]?", t)), None)
                q_["no"] = int(re.match(r"\d+", no_cell).group()) if no_cell else None
                q_["stem"] = [strip_number(statement)[1]]
                q_["opsi"] = [{"huruf": "A", "teks": "Benar"}, {"huruf": "B", "teks": "Salah"}]
                marks = lambda i: i is not None and i < len(texts) and texts[i].strip() not in ("", "□", "☐")
                if marks(ib) and not marks(is_):
                    q_["kunci_raw"] = "Benar"
                elif marks(is_) and not marks(ib):
                    q_["kunci_raw"] = "Salah"
                for c in row:
                    for p in c:
                        q_["gambar"].extend(p.images)
                soal.append(q_)
        elif isinstance(b, Para) and b.clean and not INSTR_RE.match(b.clean):
            if soal and _apply_meta(soal[-1], b):
                continue
            q_ = new_question("benar_salah")
            q_["no"], t = strip_number(b.clean)
            q_["stem"] = [t]
            q_["opsi"] = [{"huruf": "A", "teks": "Benar"}, {"huruf": "B", "teks": "Salah"}]
            soal.append(q_)
    return soal


def parse_menjodohkan(blocks) -> List[Dict[str, Any]]:
    soal = []
    for b in blocks:
        if not isinstance(b, Table):
            if isinstance(b, Para) and soal and b.clean:
                _apply_meta(soal[-1], b)
            continue
        head, rows = _table_rows(b)
        if not rows:
            continue
        ncol = max(len(r) for r in rows)
        cols = [[(b.cell_text(r[i]) if i < len(r) else "") for r in rows] for i in range(ncol)]
        content_cols = [i for i, col in enumerate(cols)
                        if any(t and not re.fullmatch(r"\d{1,3}\s*[.)]?", t) for t in col)]
        if len(content_cols) < 2:
            continue
        left, right = content_cols[0], content_cols[-1]
        num_col = next((i for i in range(ncol) if i not in content_cols), None)

        pilihan = []
        auto_i = 0
        for r in rows:
            if right >= len(r):
                continue
            cell = r[right]
            t = b.cell_text(cell)
            if not t:
                continue
            m = OPT_RE.match(t)
            if m:
                pilihan.append({"huruf": m.group(1).upper(), "teks": t[m.end():].strip()})
            elif cell and cell[0].numfmt in LETTER_FMTS:
                pilihan.append({"huruf": "ABCDEFGHIJ"[auto_i], "teks": t})
                auto_i += 1
            else:
                pilihan.append({"huruf": None, "teks": t})
        if all(p_["huruf"] is None for p_ in pilihan):
            for i, p_ in enumerate(pilihan):
                p_["huruf"] = "ABCDEFGHIJKLMNOP"[i]

        for ri, r in enumerate(rows):
            t = b.cell_text(r[left]) if left < len(r) else ""
            if not t:
                continue
            q_ = new_question("menjodohkan")
            no, t2 = strip_number(t)
            if no is None and num_col is not None and num_col < len(r):
                mm = re.match(r"\d+", b.cell_text(r[num_col]))
                no = int(mm.group()) if mm else None
            q_["no"] = no
            q_["stem"] = [t2]
            q_["pilihan"] = pilihan
            for c in r[left: left + 1]:
                for p in c:
                    q_["gambar"].extend(p.images)
            soal.append(q_)
    return soal


# ------------------------------------------------------------------------------
# Fungsi utama
# ------------------------------------------------------------------------------
def parse_docx(data: bytes) -> Dict[str, Any]:
    pkg = DocxPackage(data)
    blocks = list(iter_blocks(pkg))

    pre: List[Any] = []
    sections: List[Dict[str, Any]] = []
    for b in blocks:
        tipe = is_heading(b) if isinstance(b, Para) else None
        if tipe:
            sections.append({"judul": b.clean, "tipe": tipe, "petunjuk": "", "_blocks": []})
        elif sections:
            sections[-1]["_blocks"].append(b)
        else:
            pre.append(b)

    # ---- bagian pembuka: judul, mapel, kelas, petunjuk umum ----
    pre_paras = [p for b in pre for p in (b.paragraphs() if isinstance(b, Table) else [b])]
    judul_baris, petunjuk_umum, mapel, kelas = [], [], "", ""
    in_pet = False
    for p in pre_paras:
        t = p.clean
        if not t:
            continue
        mm = re.search(r"MAPEL\s*:\s*(.+?)(?:\s{2,}|$)", p.text.replace("\t", "  "), re.I)
        if mm and mm.group(1).strip(" .…:"):
            mapel = mm.group(1).strip()
            if re.fullmatch(r"(?:[A-Za-z] )+[A-Za-z]", mapel):
                mapel = mapel.replace(" ", "")
        if re.match(r"^petunjuk(\s+umum)?\b", t, re.I):
            in_pet = True
            continue
        if in_pet:
            petunjuk_umum.append(strip_number(t)[1])
        elif not re.match(r"^(NAMA|KELAS\s*:|NO\.?\s*ABSEN|MAPEL)", t, re.I):
            judul_baris.append(t)
            mk = re.fullmatch(r"KELAS\s+([0-9IVX]+)", t, re.I)
            if mk:
                kelas = mk.group(1)

    nomor = 0
    for sec in sections:
        blks = sec.pop("_blocks")
        # petunjuk bagian = paragraf perintah pertama
        for i, b in enumerate(blks):
            if isinstance(b, Para) and b.clean:
                if INSTR_RE.match(b.clean) or (b.clean.endswith("!") and len(b.clean) < 300):
                    sec["petunjuk"] = b.clean
                    blks = blks[:i] + blks[i + 1:]
                    # kalimat perintah lanjutan (mis. "Perhatikan soal berapa banyak...!")
                break
        t = sec["tipe"]
        if t in ("pilihan_ganda", "pg_kompleks"):
            soal = parse_option_section(blks, t)
        elif t == "benar_salah":
            soal = parse_benar_salah(blks)
        elif t == "menjodohkan":
            soal = parse_menjodohkan(blks)
        else:
            soal = parse_text_section(blks, t)

        for s in soal:
            nomor = s["no"] if (s["no"] and s["no"] > nomor) else nomor + 1
            s["no"] = nomor
            s["teks"] = "\n".join(x for x in s.pop("stem") if x)
            _validate(s)
        sec["soal"] = soal
        if not soal:
            sec["peringatan"] = "Bagian ini tidak berisi soal yang terbaca."

    return {
        "judul_baris": judul_baris,
        "mapel": mapel,
        "kelas": kelas,
        "petunjuk_umum": petunjuk_umum,
        "bagian": sections,
    }


def _validate(s: Dict[str, Any]) -> None:
    w = s["peringatan"]
    if not s["teks"] and not s["gambar"]:
        w.append("Teks soal kosong.")
    if s["tipe"] in ("pilihan_ganda", "pg_kompleks"):
        if len(s["opsi"]) < 2:
            w.append(f"Hanya {len(s['opsi'])} opsi terbaca.")
        if any(not o["teks"] for o in s["opsi"]):
            w.append("Ada opsi yang kosong.")
        teks = [o["teks"].strip().lower() for o in s["opsi"] if o["teks"]]
        if len(set(teks)) != len(teks):
            w.append("Ada opsi yang sama persis (Google Form tidak mengizinkan).")
    if s["tipe"] == "menjodohkan" and len(s["pilihan"]) < 2:
        w.append("Pilihan jawaban menjodohkan tidak terbaca.")
    # kunci dari baris "Kunci:"
    k = s.get("kunci_raw")
    if k and s["tipe"] in ("pilihan_ganda", "pg_kompleks"):
        letters = re.findall(r"\b([A-Ha-h])\b", k)
        s["kunci_raw"] = [l.upper() for l in letters] if letters else None


# ------------------------------------------------------------------------------
# Paket soal untuk skrip GAS "Buat Form"
# ------------------------------------------------------------------------------
PAKET_FORMAT = "formprint-paket-v1"


def image_to_png_b64(data: bytes, max_w: int = 1600) -> Optional[str]:
    """Mengubah gambar ke PNG (base64). None jika format tidak didukung (mis. EMF/WMF)."""
    import base64
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        im.load()
        if im.width > max_w:
            im = im.resize((max_w, int(im.height * max_w / im.width)))
        if im.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            im = im.convert("RGBA")
        buf = io.BytesIO()
        im.save(buf, format="PNG", optimize=True)
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return None


def default_key_index(s: Dict[str, Any]) -> List[int]:
    """Kunci dari dokumen (baris 'Kunci:' / tanda ✓) dalam bentuk indeks opsi."""
    k = s.get("kunci_raw")
    if not k:
        return []
    if s["tipe"] in ("pilihan_ganda", "pg_kompleks") and isinstance(k, list):
        return [ "ABCDEFGH".index(l) for l in k if l in "ABCDEFGH" and "ABCDEFGH".index(l) < len(s["opsi"])]
    if s["tipe"] == "benar_salah":
        return [0] if str(k).lower().startswith("b") else [1]
    return []


def split_title(teks: str) -> Tuple[str, str]:
    """Paragraf terakhir menjadi judul butir; paragraf sebelumnya (wacana) menjadi deskripsi."""
    parts = [t for t in teks.split("\n") if t.strip()]
    if not parts:
        return "", ""
    return parts[-1], "\n\n".join(parts[:-1])


def build_paket(parsed: Dict[str, Any], edits: Dict[str, Dict[str, Any]], settings: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """
    edits[str(no)] = {"teks": str, "opsi": [str], "kunci": [int], "poin": int}
    settings: judul, deskripsi, identitas (list), wajib, acak_soal, tampilkan_nomor,
              batasi_organisasi, poin_default
    """
    catatan: List[str] = []
    bagian = []
    for sec in parsed["bagian"]:
        out_soal = []
        for s in sec["soal"]:
            e = edits.get(str(s["no"]), {})
            teks = e.get("teks", s["teks"])
            judul, desk = split_title(teks)
            if s["tipe"] == "menjodohkan":
                opsi = [f'{p["huruf"]}. {p["teks"]}' if p.get("huruf") else p["teks"] for p in s["pilihan"]]
            else:
                opsi = e.get("opsi", [o["teks"] for o in s["opsi"]])
            gambar = []
            for i, g in enumerate(s["gambar"]):
                b64 = image_to_png_b64(g)
                if b64:
                    gambar.append({"data": b64, "mime": "image/png"})
                else:
                    catatan.append(f"No. {s['no']}: gambar ke-{i + 1} formatnya tidak didukung (mis. EMF), tambahkan manual di Form.")
            poin = int(round(float(e.get("poin", s.get("poin") or 0) or 0))) or int(settings.get("poin_default") or 0)
            out_soal.append({
                "no": s["no"], "tipe": s["tipe"], "judul": judul, "deskripsi": desk,
                "opsi": opsi if s["tipe"] in ("pilihan_ganda", "pg_kompleks", "benar_salah", "menjodohkan") else [],
                "kunci": [int(i) for i in e.get("kunci", default_key_index(s))],
                "poin": poin, "gambar": gambar,
            })
        if out_soal:
            bagian.append({"judul": sec["judul"], "petunjuk": sec.get("petunjuk", ""), "soal": out_soal})
    paket = {
        "format": PAKET_FORMAT,
        "judul": settings.get("judul", "Soal"),
        "deskripsi": settings.get("deskripsi", ""),
        "identitas": settings.get("identitas", []),
        "pengaturan": {k: settings.get(k) for k in ("wajib", "acak_soal", "tampilkan_nomor", "batasi_organisasi")},
        "template": settings.get("template") or {"pakai": False},
        "kelas_pilihan": settings.get("kelas_pilihan") or [],
        "bagian": bagian,
    }
    return paket, catatan
