import copy
import html
import io
import json
import os
import re
import shutil
import string
import subprocess
import sys
import tempfile
import threading
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import requests
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from streamlit_pdf_viewer import pdf_viewer
from PIL import Image

# Pustaka Google API untuk Opsi C (Service Account & Folder Drive Sekolah)
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False

ALPHABETS = list(string.ascii_uppercase)
pdf_lock = threading.Lock()


# ==============================================================================
# 1. FILTER BAGIAN PENUTUP ("TAHAP AKHIR / KLIK SELESAI") YANG AMAN
# ==============================================================================

def is_closing_question(q: Dict[str, Any]) -> bool:
    """
    Mendeteksi apakah suatu butir pertanyaan HANYALAH tombol konfirmasi 'KLIK SELESAI'.
    Soal ujian asli (yang punya >= 2 opsi, atau kisi/uraian) tidak akan pernah terhapus.
    """
    q_text = (q.get("pertanyaan") or "").strip().lower()
    opsi = q.get("opsi", [])

    # Soal pilihan ganda konfirmasi selesai umumnya hanya punya 1 opsi ("SELESAI")
    if len(opsi) == 1:
        opt_text = (opsi[0].get("teks") or "").strip().lower()
        if opt_text in {"selesai", "a. selesai", "ya", "sudah", "kirim", "submit"} or "selesai" in opt_text:
            return True
        if any(k in q_text for k in ["klik selesai", "yakin dengan jawabanmu", "tahap akhir"]):
            return True

    # Jika tanpa opsi tetapi teks pertanyaannya murni instruksi klik selesai
    if len(opsi) == 0 and any(k in q_text for k in ["jika kamu yakin dengan jawabanmu", "klik selesai!"]):
        return True

    return False


IDENTITY_PATTERN = re.compile(
    r"^\s*(nama(\s+(lengkap|siswa|peserta))?|kelas|rombel|mapel|mata\s+pelajaran|"
    r"no\.?\s*absen|nomor\s+absen|absen|nis|nisn|no\.?\s*peserta|nomor\s+peserta|"
    r"token|e-?mail|alamat\s+email|sekolah|asal\s+sekolah|tanggal(\s+ujian)?)\s*[:.]?\s*\*?\s*$",
    re.IGNORECASE,
)


def is_identity_question(q: Dict[str, Any]) -> bool:
    """True jika butir hanyalah isian identitas peserta (bukan soal ujian)."""
    return bool(IDENTITY_PATTERN.match((q.get("pertanyaan") or "").strip()))


def filter_valid_exam_sections(raw_sections: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Membuang bagian penutup ('SELAMAT... TAHAP AKHIR / KLIK SELESAI') 
    TANPA pernah membuang bagian yang berisi soal-soal ujian.
    """
    filtered_sections = []
    total_sec = len(raw_sections)

    for idx, sec in enumerate(raw_sections):
        soal_asli = sec.get("soal", [])

        # Bagian identitas (Nama, Kelas, Mapel, No. Absen, Token, dst.) selalu dibuang,
        # apa pun cara Form dibaca (multi-bagian, salinan satu halaman, atau API).
        if soal_asli and all(is_identity_question(q) for q in soal_asli):
            continue
        judul = (sec.get("nama_bagian") or "").lower()
        petunjuk = (sec.get("petunjuk") or "").lower()

        # Jika bagian ini memiliki > 1 soal, ini PASTI bagian soal ujian.
        # Kita hanya membuang butir konfirmasi 'KLIK SELESAI' jika kebetulan diselipkan di akhir.
        valid_soal = [q for q in soal_asli if not is_closing_question(q)]

        # Jika ini bagian paling akhir dan sejak awal hanya berisi <= 1 soal,
        # cek apakah judulnya 'tahap akhir' atau soalnya adalah konfirmasi selesai
        if idx == total_sec - 1 and len(soal_asli) <= 1:
            if "tahap akhir" in judul or "tahap akhir" in petunjuk or "akhiri dengan doa" in petunjuk:
                continue
            if len(valid_soal) == 1 and len(valid_soal[0].get("opsi", [])) <= 1 and valid_soal[0].get("tipe") == "pilihan_ganda":
                continue

        # Simpan bagian selama masih memiliki soal valid
        if valid_soal:
            sec_copy = copy.deepcopy(sec)
            sec_copy["soal"] = valid_soal
            filtered_sections.append(sec_copy)

    # Pengaman Fallback: Jika karena alasan tertentu hasil filter kosong padahal raw_sections ada isinya,
    # kembalikan seluruh bagian yang memiliki soal agar dokumen tidak pernah kosong.
    if not filtered_sections and raw_sections:
        filtered_sections = [copy.deepcopy(s) for s in raw_sections if s.get("soal")]

    # Urutkan ulang nomor soal (untuk kebutuhan Lembar Jawaban & Kunci Jawaban)
    global_no = 1
    for sec in filtered_sections:
        for local_no, q in enumerate(sec.get("soal", []), start=1):
            q["no"] = local_no
            q["no_global"] = global_no
            global_no += 1

    return filtered_sections


# ==============================================================================
# 2. UTILITAS SISTEM, HEADER HALAMAN 1, & PENGUNCI HALAMAN
# ==============================================================================

def get_libreoffice_command() -> List[str]:
    if shutil.which("libreoffice"):
        return ["libreoffice"]
    if shutil.which("soffice"):
        return ["soffice"]

    if sys.platform.startswith("win"):
        for path in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]:
            if os.path.exists(path):
                return [path]
    elif sys.platform == "darwin":
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_path):
            return [mac_path]
    return []


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> Optional[str]:
    cmd_base = get_libreoffice_command()
    if not cmd_base:
        return None

    cmd = cmd_base + ["--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir]
    try:
        env = os.environ.copy()
        if not sys.platform.startswith("win"):
            env["HOME"] = "/tmp"

        with pdf_lock:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)

        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception:
        return None


def strip_legacy_numbering_tags(docx_path: str) -> None:
    """Menghapus tag penomoran otomatis pada template Naskah Soal."""
    doc_obj = Document(docx_path)
    patterns = [
        r"\{\{\s*item\.no_global\s*\}\}\.\s*",
        r"\{\{\s*item\.no\s*\}\}\.\s*",
        r"\{\{\s*b\.sub_no\s*\}\}\.\s*",
        r"\{\{\s*b\.no_angka\s*\}\}\.\s*",
        r"\{\{\s*k\.no_angka\s*\}\}\.\s*",
        r"\(\{\{\s*b\.sub_no\s*\}\}\)\s*",
        r"\(\{\{\s*k\.no_angka\s*\}\}\)\s*",
    ]
    combined = re.compile("|".join(patterns))

    for p in doc_obj.paragraphs:
        if "item.pertanyaan" in p.text:
            p.paragraph_format.left_indent = Mm(0)
            p.paragraph_format.first_line_indent = Mm(0)
        for r in p.runs:
            if r.text and any(k in r.text for k in ["no_global", "item.no", "sub_no", "no_angka"]):
                r.text = combined.sub("", r.text)

    for tbl in doc_obj.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.text and any(k in r.text for k in ["no_global", "item.no", "sub_no", "no_angka"]):
                            r.text = combined.sub("", r.text)

    doc_obj.save(docx_path)


def clear_header_completely(header_obj) -> None:
    """Menghapus seluruh teks, gambar logo, dan baris kosong pada objek Header Word."""
    hdr_el = header_obj._element
    for child in list(hdr_el):
        hdr_el.remove(child)
    hdr_el.append(OxmlElement("w:p"))


def configure_header_first_page_only(docx_obj) -> None:
    """Memastikan Kop Surat di Header hanya tampil di Halaman 1 dan bersih di Halaman 2 dst."""
    if not docx_obj.sections:
        return

    sec0 = docx_obj.sections[0]
    sectPr0 = sec0._sectPr

    default_hdr_ref = None
    first_hdr_ref = None
    for ref in sectPr0.findall(qn("w:headerReference")):
        if ref.get(qn("w:type")) == "default":
            default_hdr_ref = ref
        elif ref.get(qn("w:type")) == "first":
            first_hdr_ref = ref

    has_default_content = False
    if default_hdr_ref is not None:
        hdr_el = sec0.header._element
        has_text = bool("".join(hdr_el.itertext()).strip())
        has_tbl = bool(hdr_el.findall(".//" + qn("w:tbl")))
        has_img = bool(hdr_el.findall(".//" + qn("w:drawing")) or hdr_el.findall(".//" + qn("w:pict")))
        has_default_content = has_text or has_tbl or has_img

    if has_default_content and default_hdr_ref is not None:
        first_has_content = False
        if first_hdr_ref is not None:
            first_el = sec0.first_page_header._element
            first_has_content = (
                bool("".join(first_el.itertext()).strip())
                or bool(first_el.findall(".//" + qn("w:tbl")))
                or bool(first_el.findall(".//" + qn("w:drawing")))
                or bool(first_el.findall(".//" + qn("w:pict")))
            )
        if not first_has_content:
            if first_hdr_ref is not None:
                sectPr0.remove(first_hdr_ref)
            default_hdr_ref.set(qn("w:type"), "first")
        else:
            clear_header_completely(sec0.header)

    if sectPr0.find(qn("w:titlePg")) is None:
        sectPr0.append(OxmlElement("w:titlePg"))

    for sec in docx_obj.sections[1:]:
        sPr = sec._sectPr
        titlePg = sPr.find(qn("w:titlePg"))
        if titlePg is not None:
            sPr.remove(titlePg)
        sec.header.is_linked_to_previous = False
        clear_header_completely(sec.header)


EMPTY_HDR_RID = "rIdFormPrintEmptyHdr"
EMPTY_HDR_NAME = "header_formprint_empty.xml"

def fix_header_first_page_only_file(docx_path: str) -> None:
    """
    Memastikan Kop (header) hanya tampil di halaman 1 dengan menyunting XML DOCX
    secara langsung SETELAH disimpan (perubahan header lewat python-docx tidak
    selalu ikut tersimpan setelah docxtpl merender header).
    - Section pertama: header "first" = kop, header "default" = header kosong, titlePg aktif.
    - Section berikutnya: tidak punya header sendiri (mewarisi section pertama).
    """
    with zipfile.ZipFile(docx_path) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    doc = files["word/document.xml"].decode("utf-8")
    sect_prs = list(re.finditer(r"<w:sectPr\b[^>]*?(?:/>|>.*?</w:sectPr>)", doc, re.S))
    if not sect_prs:
        return

    def fix_first(sp: str) -> str:
        if sp.endswith("/>"):
            sp = sp[:-2] + "></w:sectPr>"
        refs = re.findall(r"<w:headerReference\b[^>]*/>", sp)
        types = {re.search(r'w:type="(\w+)"', r).group(1): r for r in refs if re.search(r'w:type="(\w+)"', r)}
        if not types:
            return sp
        if "first" not in types and "default" in types:
            sp = sp.replace(types["default"], types["default"].replace('w:type="default"', 'w:type="first"'), 1)
        elif "default" in types:
            sp = sp.replace(types["default"], "", 1)
        open_tag = re.match(r"<w:sectPr\b[^>]*>", sp).group(0)
        sp = sp.replace(open_tag, open_tag + f'<w:headerReference w:type="default" r:id="{EMPTY_HDR_RID}"/>', 1)
        if "<w:titlePg" not in sp:
            m = re.search(r"<w:(textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b", sp)
            pos = m.start() if m else sp.rindex("</w:sectPr>")
            sp = sp[:pos] + "<w:titlePg/>" + sp[pos:]
        return sp

    def fix_later(sp: str) -> str:
        sp = re.sub(r"<w:headerReference\b[^>]*/>", "", sp)
        sp = re.sub(r"<w:titlePg\b[^>]*/>", "", sp)
        return sp

    pieces, last = [], 0
    for i, m in enumerate(sect_prs):
        pieces.append(doc[last:m.start()])
        pieces.append(fix_first(m.group(0)) if i == 0 else fix_later(m.group(0)))
        last = m.end()
    pieces.append(doc[last:])
    doc = "".join(pieces)

    if EMPTY_HDR_RID in doc:
        if 'xmlns:r=' not in doc[:2000]:
            doc = doc.replace("<w:document ", '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        files[f"word/{EMPTY_HDR_NAME}"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>'
        ).encode("utf-8")
        rels = files["word/_rels/document.xml.rels"].decode("utf-8")
        if EMPTY_HDR_RID not in rels:
            rels = rels.replace("</Relationships>",
                f'<Relationship Id="{EMPTY_HDR_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="{EMPTY_HDR_NAME}"/></Relationships>')
            files["word/_rels/document.xml.rels"] = rels.encode("utf-8")
        ct = files["[Content_Types].xml"].decode("utf-8")
        if EMPTY_HDR_NAME not in ct:
            ct = ct.replace("</Types>",
                f'<Override PartName="/word/{EMPTY_HDR_NAME}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
            files["[Content_Types].xml"] = ct.encode("utf-8")

    files["word/document.xml"] = doc.encode("utf-8")
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in files.items():
            zout.writestr(n, b)
    shutil.move(tmp, docx_path)


def lock_layout_prevent_page_split(docx_obj) -> None:
    """Mengunci paragraf soal, gambar, dan baris tabel agar tidak terpotong lintas halaman.
    (Pengaturan Kop halaman 1 dilakukan oleh fix_header_first_page_only_file setelah disimpan.)"""
    for p in docx_obj.paragraphs:
        text = p.text.strip()
        p.paragraph_format.keep_together = True
        if text.startswith("Petunjuk:"):
            # JANGAN dirantai ke soal pertama: judul + petunjuk panjang + soal bergambar
            # menjadi satu blok yang terlalu tinggi, sehingga seluruhnya terlempar ke halaman 2.
            p.paragraph_format.keep_with_next = False
        elif "w:drawing" in p._p.xml:
            p.paragraph_format.keep_with_next = True

    for table in docx_obj.tables:
        for row_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            tblHeader = trPr.find(qn("w:tblHeader"))
            if tblHeader is not None:
                trPr.remove(tblHeader)

            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))

            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_together = True
                    if len(table.rows) <= 3 and row_idx < len(table.rows) - 1:
                        p.paragraph_format.keep_with_next = True


def clean_option_prefix(text: str, expected_letter: str) -> str:
    t = text.strip()
    pattern = rf"^(\({expected_letter}\)|{expected_letter}[\.\)])\s+"
    return re.sub(pattern, "", t, flags=re.IGNORECASE).strip()


IMAGE_WARNINGS: List[str] = []
IMAGE_DEBUG: Dict[str, Any] = {}


def download_image_to_stream(img_url: str) -> Optional[io.BytesIO]:
    """
    Mengunduh gambar dan memastikan hasilnya benar-benar gambar PNG/JPEG.
    python-docx tidak mendukung WebP, dan server Google kadang mengirim WebP
    atau halaman HTML (login/error) dengan status 200.
    """
    if not img_url:
        return None
    # URL gambar Google Form berakhiran "=w150" dsb. (ukuran tampilan di layar).
    # Minta versi beresolusi lebih tinggi agar tajam saat dicetak; jika gagal,
    # kembali ke URL asli.
    hi_res = re.sub(r"=w\d+(-h\d+)?$", "=w1200", img_url)
    if hi_res != img_url:
        result = _download_image(hi_res, record_warning=False)
        if result is not None:
            return result
    return _download_image(img_url, record_warning=True)


def _download_image(img_url: str, record_warning: bool = True) -> Optional[io.BytesIO]:
    try:
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept": "image/png,image/jpeg,image/gif;q=0.9,*/*;q=0.5",
        }
        resp = requests.get(img_url, headers=headers, timeout=15)
        if resp.status_code != 200 or not resp.content:
            if record_warning:
                IMAGE_WARNINGS.append(f"Gagal mengunduh gambar (HTTP {resp.status_code}).")
            return None

        IMAGE_DEBUG["gambar_terunduh"] = IMAGE_DEBUG.get("gambar_terunduh", 0) + 1
        img = Image.open(io.BytesIO(resp.content))
        img.load()
        if (img.format or "").upper() in {"PNG", "JPEG"}:
            return io.BytesIO(resp.content)

        # WebP / GIF / format lain -> konversi ke PNG agar bisa disisipkan ke DOCX
        if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            img = img.convert("RGBA")
        out = io.BytesIO()
        img.save(out, format="PNG")
        out.seek(0)
        return out
    except Exception as e:
        if record_warning:
            IMAGE_WARNINGS.append(f"Gambar tidak dapat dibaca: {type(e).__name__}.")
        return None


def find_image_url(data: Any) -> Optional[str]:
    if isinstance(data, str):
        if "googleusercontent.com" in data or "/forms-images" in data or (
            data.startswith("http") and any(ext in data.lower() for ext in [".png", ".jpg", ".jpeg", ".webp"])
        ):
            return data
    elif isinstance(data, list):
        for item in data:
            found = find_image_url(item)
            if found:
                return found
    elif isinstance(data, dict):
        for value in data.values():
            found = find_image_url(value)
            if found:
                return found
    return None

def extract_html_images_by_item(html_text: str, raw_items: list) -> Dict[int, List[str]]:
    """Mengekstrak URL gambar dari tag <img> pada satu halaman HTML bagian Google Form."""
    # Buang seluruh <script>/<style> (termasuk FB_PUBLIC_LOAD_DATA_) TANPA memotong
    # isi halaman. Versi lama memotong di "FB_PUBLIC_LOAD_DATA_": bila skrip itu
    # muncul sebelum isi <form>, seluruh tag <img> ikut terbuang.
    html_body = re.sub(r"<script\b[^>]*>.*?</script>", "", html_text, flags=re.DOTALL | re.IGNORECASE)
    html_body = re.sub(r"<style\b[^>]*>.*?</style>", "", html_body, flags=re.DOTALL | re.IGNORECASE)
    item_positions = []

    for it in raw_items:
        if not it or not isinstance(it, list) or not it[0]:
            continue
        item_id = it[0]
        for marker in [f'data-item-id="{item_id}"', f"%.@.[{item_id},", f"[{item_id},&quot;"]:
            pos = html_body.find(marker)
            if pos != -1:
                item_positions.append((pos, item_id))
                break

    item_positions.sort(key=lambda x: x[0])
    img_map: Dict[int, List[str]] = {}
    img_pattern = re.compile(
        r'<img[^>]+?(?:data-src|src)=["\'](https://[^"\']*(?:googleusercontent\.com|ggpht\.com|drive\.google\.com|docs\.google\.com/forms-images)[^"\']+)["\']',
        re.IGNORECASE
    )

    for idx, (start_pos, item_id) in enumerate(item_positions):
        end_pos = item_positions[idx + 1][0] if idx + 1 < len(item_positions) else len(html_body)
        chunk = html_body[start_pos:end_pos]
        urls = []
        for u in img_pattern.findall(chunk):
            clean_u = html.unescape(u)
            if clean_u not in urls:
                urls.append(clean_u)
        if urls:
            img_map[item_id] = urls

    return img_map


def build_dummy_answer_for_subq(sub_q: list, item_type: int) -> str:
    """Membuat isian dummy otomatis (termasuk membaca opsi Kelas/Token) agar bisa melewati Bagian 1."""
    # 1. Jika soal memiliki opsi pilihan (Dropdown Kelas, Pilihan Ganda, Kotak Centang)
    if len(sub_q) > 1 and sub_q[1] and isinstance(sub_q[1], list):
        for opt in sub_q[1]:
            if opt and opt[0] is not None and str(opt[0]).strip() != "":
                return str(opt[0]).strip()

    # 2. Jika soal isian memiliki validasi khusus (misal Token atau Angka)
    if len(sub_q) > 4 and sub_q[4] and isinstance(sub_q[4], list):
        try:
            rule = sub_q[4][0]
            v_type, v_op = rule[0], rule[1]
            v_params = rule[2] if len(rule) > 2 and rule[2] else []
            if v_params and v_params[0] is not None:
                p0 = str(v_params[0]).strip()
                if v_type == 1:  # Validasi angka
                    num = float(p0)
                    if v_op == 1:
                        return str(int(num + 1))
                    elif v_op == 3:
                        return str(int(num - 1))
                    return str(int(num)) if num.is_integer() else str(num)
                return p0.strip("^$")
            if v_type == 1:
                return "1"
        except Exception:
            pass

    if item_type == 9:
        return "2026-01-01"
    if item_type == 10:
        return "08:00"
    return "1"


def _get_hidden_input(html_text: str, name: str) -> Optional[str]:
    """Membaca nilai <input type="hidden" name="..."> dari HTML Google Form."""
    for pat in (
        rf'<input[^>]*name="{re.escape(name)}"[^>]*value="([^"]*)"',
        rf'<input[^>]*value="([^"]*)"[^>]*name="{re.escape(name)}"',
    ):
        m = re.search(pat, html_text)
        if m:
            return html.unescape(m.group(1))
    return None


def build_page_entries(page_items: list) -> Dict[str, str]:
    """
    Membuat isian dummy dalam format field asli formulir (entry.XXXX) untuk
    SATU halaman — persis seperti yang dikirim browser saat klik 'Berikutnya'.
    """
    data: Dict[str, str] = {}
    for it in page_items:
        if not it or len(it) < 5 or not it[4] or not isinstance(it[4], list):
            continue
        item_type = it[3]
        if item_type == 13:  # unggah file: tidak bisa diisi otomatis
            continue
        for sub_q in it[4]:
            if not sub_q or not isinstance(sub_q, list) or not sub_q[0]:
                continue
            eid = sub_q[0]
            if item_type == 9:  # tanggal dikirim terpisah: tahun/bulan/hari
                data[f"entry.{eid}_year"] = "2026"
                data[f"entry.{eid}_month"] = "1"
                data[f"entry.{eid}_day"] = "1"
            elif item_type == 10:  # waktu dikirim terpisah: jam/menit
                data[f"entry.{eid}_hour"] = "08"
                data[f"entry.{eid}_minute"] = "00"
            else:
                data[f"entry.{eid}"] = build_dummy_answer_for_subq(sub_q, item_type)
    return data


def fetch_all_sections_images(viewform_url: str, first_html: str, raw_items: list) -> Dict[int, List[str]]:
    """
    Mensimulasikan klik tombol 'Berikutnya' (POST /formResponse) untuk membuka
    Bagian 2, 3, dst. sehingga seluruh gambar soal di balik halaman identitas dapat diambil.

    Google Form hanya menampilkan <img> untuk halaman yang SEDANG dibuka. Karena
    Bagian 1 (identitas) dilewati, semua gambar soal berasal dari halaman 2 dst.
    Jadi bila navigasi ini gagal, TIDAK ADA satu pun gambar yang terambil.
    """
    pages_items: List[list] = [[]]
    for it in raw_items:
        if it and len(it) >= 4 and it[3] == 8:
            pages_items.append([])
        else:
            pages_items[-1].append(it)

    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0", "Referer": viewform_url}

    # Ambil ulang halaman 1 memakai session yang sama agar cookie & fbzx konsisten
    try:
        r0 = session.get(viewform_url, headers=headers, timeout=10)
        if r0.status_code == 200 and "FB_PUBLIC_LOAD_DATA_" in r0.text:
            first_html = r0.text
    except Exception:
        pass

    img_map = extract_html_images_by_item(first_html, raw_items)

    # ---- Data diagnosa gambar (ditampilkan di panel "Diagnosa gambar") ----
    all_imgs = re.findall(r"<img\b[^>]*>", first_html, flags=re.IGNORECASE)
    IMAGE_DEBUG.clear()
    IMAGE_DEBUG.update({
        "jumlah_tag_img_di_html": len(all_imgs),
        "contoh_tag_img": [t[:300] for t in all_imgs[:8]],
        "jumlah_item_di_form": len(raw_items),
        "item_yang_ditemukan_di_html": sum(
            1 for it in raw_items if it and isinstance(it, list) and it[0] and (
                f'data-item-id="{it[0]}"' in first_html or f"%.@.[{it[0]}," in first_html
                or f"[{it[0]},&quot;" in first_html)
        ),
        "item_tipe_gambar_(11)": sum(1 for it in raw_items if it and len(it) >= 4 and it[3] == 11),
        "item_dengan_gambar_terpetakan": {str(k): len(v) for k, v in img_map.items()},
        "jumlah_halaman_form": len(pages_items),
        "ukuran_html": len(first_html),
    })
    IMAGE_DEBUG["_html"] = first_html

    if len(pages_items) <= 1:
        return img_map

    action_match = re.search(r'<form[^>]+action=["\'](https://[^"\']+formResponse)["\']', first_html)
    form_response_url = action_match.group(1) if action_match else re.sub(r"/viewform.*$", "/formResponse", viewform_url)

    fbzx = _get_hidden_input(first_html, "fbzx") or "0"
    page_history = _get_hidden_input(first_html, "pageHistory") or "0"
    partial = _get_hidden_input(first_html, "partialResponse")
    needs_email = 'name="emailAddress"' in first_html

    for _ in range(len(pages_items)):
        current_page = int(page_history.split(",")[-1])
        if current_page >= len(pages_items) - 1:
            break

        # Jawaban halaman AKTIF dikirim sebagai entry.XXX; jawaban halaman
        # sebelumnya sudah dibawa oleh partialResponse dari server.
        post_data: Dict[str, str] = {
            "fvv": "1",
            "fbzx": fbzx,
            "pageHistory": page_history,
            "continue": "1",
        }
        if partial:
            post_data["partialResponse"] = partial
        if needs_email:
            post_data["emailAddress"] = "contoh@example.com"
        post_data.update(build_page_entries(pages_items[current_page]))

        try:
            resp_page = session.post(form_response_url, data=post_data, headers=headers, timeout=15)
        except Exception as e:
            IMAGE_WARNINGS.append(f"Gagal membuka Bagian {current_page + 2}: {type(e).__name__}.")
            break
        if resp_page.status_code != 200:
            if resp_page.status_code in (401, 403):
                IMAGE_WARNINGS.append(
                    f"Google menolak membuka Bagian {current_page + 2} tanpa login (HTTP {resp_page.status_code}). "
                    "Solusi: buat salinan Form versi satu halaman (ganti pembatas bagian dengan "
                    "'Judul & deskripsi'), lalu gunakan link salinan tersebut."
                )
            else:
                IMAGE_WARNINGS.append(f"Gagal membuka Bagian {current_page + 2} (HTTP {resp_page.status_code}).")
            break

        new_history = _get_hidden_input(resp_page.text, "pageHistory")
        if not new_history or new_history == page_history:
            IMAGE_WARNINGS.append(
                f"Tidak bisa melewati Bagian {current_page + 1} secara otomatis "
                "(isian wajib/validasi tidak terpenuhi). Gambar pada bagian berikutnya tidak terambil."
            )
            break

        img_map.update(extract_html_images_by_item(resp_page.text, raw_items))
        page_history = new_history
        partial = _get_hidden_input(resp_page.text, "partialResponse") or partial
        fbzx = _get_hidden_input(resp_page.text, "fbzx") or fbzx

    return img_map

# ==============================================================================
# 3. KONEKSI GOOGLE SERVICE ACCOUNT (OPSI C - FOLDER DRIVE & KUNCI OTOMATIS)
# ==============================================================================

def get_google_credentials():
    if not GOOGLE_API_AVAILABLE:
        return None

    raw_json = None
    if "GOOGLE_CREDENTIALS_JSON" in st.secrets:
        raw_json = st.secrets["GOOGLE_CREDENTIALS_JSON"]
    elif os.environ.get("GOOGLE_CREDENTIALS_JSON"):
        raw_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")

    if not raw_json:
        return None

    try:
        info = json.loads(raw_json) if isinstance(raw_json, str) else dict(raw_json)
        scopes = [
            "https://www.googleapis.com/auth/forms.body.readonly",
            "https://www.googleapis.com/auth/drive.readonly",
        ]
        return service_account.Credentials.from_service_account_info(info, scopes=scopes)
    except Exception:
        return None


def list_forms_in_drive_folder(creds, folder_id: str) -> List[Dict[str, str]]:
    try:
        drive_service = build("drive", "v3", credentials=creds)
        query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.form' and trashed=false"
        results = drive_service.files().list(
            q=query,
            pageSize=100,
            fields="files(id, name)",
            orderBy="name"
        ).execute()
        return results.get("files", [])
    except Exception:
        return []


def extract_form_id_from_url(url_or_id: str) -> Optional[str]:
    clean = url_or_id.strip()
    if "/forms/d/e/" in clean:
        return None
    m = re.search(r"/forms/d/([a-zA-Z0-9_-]+)", clean)
    if m:
        return m.group(1)
    if re.match(r"^[a-zA-Z0-9_-]{20,}$", clean):
        return clean
    return None


# ==============================================================================
# 4. EKSTRAKSI VIA GOOGLE FORMS API & LINK PUBLIK
# ==============================================================================

def extract_via_forms_api(form_id: str, creds) -> Dict[str, Any]:
    forms_service = build("forms", "v1", credentials=creds)
    form_data = forms_service.forms().get(formId=form_id).execute()

    info = form_data.get("info", {})
    form_title = info.get("title") or "Naskah Soal Ujian"
    form_desc = info.get("description") or ""
    items = form_data.get("items", [])

    has_multiple_sections = any("pageBreakItem" in it for it in items)
    in_first_section = has_multiple_sections

    raw_sections = []
    current_section = {"nama_bagian": "", "petunjuk": "", "soal": []}
    pending_api_img = None

    for item in items:
        if "pageBreakItem" in item:
            if in_first_section:
                in_first_section = False
            else:
                if current_section["soal"] or current_section["nama_bagian"]:
                    raw_sections.append(current_section)

            current_section = {
                "nama_bagian": item.get("title") or "",
                "petunjuk": item.get("description") or "",
                "soal": []
            }
            continue

        if in_first_section:
            continue
        if "imageItem" in item:
            img_uri = item["imageItem"].get("image", {}).get("contentUri")
            if img_uri:
                pending_api_img = download_image_to_stream(img_uri)
            continue

        if "questionGroupItem" in item:
            group = item["questionGroupItem"]
            q_title = item.get("title") or ""
            q_img_url = group.get("image", {}).get("contentUri")
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else pending_api_img
            pending_api_img = None

            columns = [
                opt.get("value", "").strip()
                for opt in group.get("grid", {}).get("columns", {}).get("options", [])
            ]
            questions = group.get("questions", [])

            rows = []
            group_poin = 0
            for idx, rq in enumerate(questions):
                row_text = rq.get("rowQuestion", {}).get("title", "").strip()
                grading = rq.get("grading", {})
                r_poin = grading.get("pointValue", 0)
                group_poin += r_poin
                correct_vals = [
                    ans.get("value", "").strip()
                    for ans in grading.get("correctAnswers", {}).get("answers", [])
                ]

                kunci_idx = -1
                kunci_huruf = ""
                kunci_teks = ", ".join(correct_vals)
                for c_i, col_val in enumerate(columns):
                    if col_val in correct_vals:
                        kunci_idx = c_i
                        kunci_huruf = ALPHABETS[c_i] if c_i < len(ALPHABETS) else str(c_i + 1)
                        break

                rows.append({
                    "idx": idx,
                    "sub_no": chr(97 + idx) if idx < 26 else str(idx + 1),
                    "no_angka": str(idx + 1),
                    "teks": row_text,
                    "is_benar": (kunci_idx == 0),
                    "is_salah": (kunci_idx == 1),
                    "kunci_idx": kunci_idx,
                    "kunci_huruf": kunci_huruf,
                    "kunci_teks": kunci_teks,
                    "poin": r_poin
                })

            cols_lower = [c.lower() for c in columns]
            opsi_matriks_standar = {"a", "b", "c", "d", "1", "2", "3", "4"}

            if len(columns) == 1:
                pgk_opts = []
                for i, r in enumerate(rows):
                    letter = ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})"
                    pgk_opts.append({
                        "huruf": letter,
                        "teks": clean_option_prefix(r["teks"], letter),
                        "img_bytes": None,
                        "ada_gambar": False,
                        "is_kunci": bool(r["kunci_teks"])
                    })
                current_section["soal"].append({
                    "item_id": str(item.get("itemId") or ""),
                    "pertanyaan": q_title, "tipe": "pg_kompleks",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": True, "opsi_bergambar": False, "opsi": pgk_opts,
                    "baris": [], "kolom_header": [], "kiri": [], "kanan": [],
                    "poin": group_poin, "kunci_teks": ""
                })
            elif len(columns) == 2 and any(k in cols_lower for k in ["benar", "salah", "b", "s", "true", "false", "ya", "tidak"]):
                current_section["soal"].append({
                    "item_id": str(item.get("itemId") or ""),
                    "pertanyaan": q_title, "tipe": "benar_salah",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [],
                    "poin": group_poin, "kunci_teks": ""
                })
            elif len(columns) <= 4 and all(c.lower() in opsi_matriks_standar for c in columns):
                current_section["soal"].append({
                    "item_id": str(item.get("itemId") or ""),
                    "pertanyaan": q_title, "tipe": "kisi_pilihan_ganda",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [],
                    "poin": group_poin, "kunci_teks": ""
                })
            else:
                pilihan_kanan = [
                    {
                        "huruf": ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})",
                        "teks": clean_option_prefix(col_text, ALPHABETS[i] if i < len(ALPHABETS) else "")
                    }
                    for i, col_text in enumerate(columns)
                ]
                current_section["soal"].append({
                    "item_id": str(item.get("itemId") or ""),
                    "pertanyaan": q_title, "tipe": "menjodohkan",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": [], "kolom_header": columns, "kiri": rows, "kanan": pilihan_kanan,
                    "poin": group_poin, "kunci_teks": ""
                })
            continue

        if "questionItem" in item:
            q_item = item["questionItem"]
            q_obj = q_item.get("question", {})
            q_text = item.get("title") or ""

            q_img_url = q_item.get("image", {}).get("contentUri")
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else pending_api_img
            pending_api_img = None

            grading = q_obj.get("grading", {})
            poin = grading.get("pointValue", 0)
            correct_vals = [
                ans.get("value", "").strip()
                for ans in grading.get("correctAnswers", {}).get("answers", [])
            ]

            options = []
            tipe_soal = "isian"

            if "choiceQuestion" in q_obj:
                cq = q_obj["choiceQuestion"]
                c_type = cq.get("type", "RADIO")
                raw_opts = cq.get("options", [])

                for idx, opt in enumerate(raw_opts):
                    raw_val = str(opt.get("value", "")).strip()
                    letter = ALPHABETS[idx] if idx < len(ALPHABETS) else f"({idx+1})"
                    opt_img_url = opt.get("image", {}).get("contentUri")
                    opt_img_bytes = download_image_to_stream(opt_img_url) if opt_img_url else None

                    options.append({
                        "huruf": letter,
                        "teks": clean_option_prefix(raw_val, letter),
                        "img_bytes": opt_img_bytes,
                        "ada_gambar": opt_img_bytes is not None,
                        "is_kunci": (raw_val in correct_vals)
                    })

                tipe_soal = "pg_kompleks" if c_type == "CHECKBOX" else "pilihan_ganda"

            current_section["soal"].append({
                    "item_id": str(item.get("itemId") or ""),
                "pertanyaan": q_text, "tipe": tipe_soal,
                "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                "ada_opsi": len(options) > 0,
                "opsi_bergambar": any(o["ada_gambar"] for o in options),
                "opsi": options, "baris": [], "kolom_header": [], "kiri": [], "kanan": [],
                "poin": poin, "kunci_teks": ", ".join(correct_vals)
            })

    if current_section["soal"] or current_section["nama_bagian"]:
        raw_sections.append(current_section)

    daftar_bagian = filter_valid_exam_sections(raw_sections)
    total_poin = sum(q.get("poin", 0) for sec in daftar_bagian for q in sec["soal"])

    return {
        "judul_form": form_title,
        "deskripsi_form": form_desc,
        "total_poin": total_poin if total_poin > 0 else 100,
        "daftar_bagian": daftar_bagian
    }


def extract_via_public_html(form_url: str) -> Dict[str, Any]:
    clean_url = re.sub(r"/(edit|viewform|formResponse)/?$", "", form_url.strip().split("?")[0])
    clean_url = clean_url.rstrip("/") + "/viewform"

    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(clean_url, headers=headers, timeout=10)
    resp.raise_for_status()

    pattern = r"FB_PUBLIC_LOAD_DATA_\s*=\s*(.*?);</script>"
    match = re.search(pattern, resp.text, re.DOTALL)
    if not match:
        raise ValueError("Formulir tidak ditemukan atau belum disetel ke mode Publik.")

    raw_data = json.loads(match.group(1))
    form_title = raw_data[1][8] or "Naskah Soal Ujian"
    form_desc = raw_data[1][0] or ""
    raw_items = raw_data[1][1] or []

    has_multiple_sections = any(it and len(it) >= 4 and it[3] == 8 for it in raw_items)

    # MODE SATU HALAMAN: Form tanpa pembatas bagian, tetapi memakai item
    # "Judul & deskripsi" (tipe 6) sebagai judul bagian. Semua soal tampil di
    # halaman pertama sehingga gambar bisa diambil tanpa meniru klik
    # "Berikutnya" (yang bisa ditolak Google dengan HTTP 401).
    # Item sebelum judul pertama (misalnya isian identitas) dilewati.
    single_page_mode = (not has_multiple_sections) and any(
        it and len(it) >= 4 and it[3] == 6 for it in raw_items
    )
    section_marker_type = 6 if single_page_mode else 8
    in_first_section = has_multiple_sections or single_page_mode

    raw_sections = []
    current_section = {"nama_bagian": "", "petunjuk": "", "soal": []}

    html_img_map = fetch_all_sections_images(resp.url, resp.text, raw_items)
    pending_standalone_img = None

    for item in raw_items:
        if not item or len(item) < 4:
            continue
        item_type = item[3]

        if item_type == section_marker_type:
            if in_first_section:
                in_first_section = False
            elif current_section["soal"] or current_section["nama_bagian"]:
                raw_sections.append(current_section)

            current_section = {
                "nama_bagian": item[1] or "",
                "petunjuk": item[2] or "",
                "soal": []
            }
            continue

        if in_first_section:
            continue

        # Item "Tambahkan gambar" di Google Form bertipe 11 (tipe 6 = Judul & Deskripsi).
        # Versi lama hanya mengecek tipe 6, sehingga gambar mandiri selalu terlewat.
        if item_type in (6, 11):
            imgs = html_img_map.get(item[0], [])
            img_url = imgs[0] if imgs else find_image_url(item)
            if img_url:
                pending_standalone_img = download_image_to_stream(img_url)
            continue

        if item_type == 7 and len(item) >= 5 and item[4]:
            q_title = item[1] or ""
            rows_data = item[4] or []
            item_imgs = html_img_map.get(item[0], [])
            q_img_url = item_imgs[0] if item_imgs else find_image_url(item)
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else pending_standalone_img
            pending_standalone_img = None

            columns = []
            if rows_data and len(rows_data[0]) > 1 and rows_data[0][1]:
                for col_item in rows_data[0][1]:
                    if col_item and col_item[0]:
                        columns.append(str(col_item[0]).strip())

            rows = []
            for idx, r in enumerate(rows_data):
                row_text = (r[3][0] if isinstance(r[3], list) else str(r[3])) if (len(r) > 3 and r[3]) else ""
                rows.append({
                    "idx": idx,
                    "sub_no": chr(97 + idx) if idx < 26 else str(idx + 1),
                    "no_angka": str(idx + 1),
                    "teks": row_text.strip(),
                    "is_benar": False, "is_salah": False,
                    "kunci_idx": -1, "kunci_huruf": "", "kunci_teks": ""
                })

            cols_lower = [c.lower() for c in columns]
            opsi_matriks_standar = {"a", "b", "c", "d", "1", "2", "3", "4"}

            if len(columns) == 1:
                pgk_opts = []
                for i, b in enumerate(rows):
                    letter = ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})"
                    pgk_opts.append({
                        "huruf": letter,
                        "teks": clean_option_prefix(b["teks"], letter),
                        "img_bytes": None, "ada_gambar": False, "is_kunci": False
                    })
                current_section["soal"].append({
                    "item_id": str(item[0]),
                    "pertanyaan": q_title, "tipe": "pg_kompleks",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": True, "opsi_bergambar": False, "opsi": pgk_opts,
                    "baris": [], "kolom_header": [], "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
                })
            elif len(columns) == 2 and any(k in cols_lower for k in ["benar", "salah", "b", "s", "true", "false", "ya", "tidak"]):
                current_section["soal"].append({
                    "item_id": str(item[0]),
                    "pertanyaan": q_title, "tipe": "benar_salah",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
                })
            elif len(columns) <= 4 and all(c.lower() in opsi_matriks_standar for c in columns):
                current_section["soal"].append({
                    "item_id": str(item[0]),
                    "pertanyaan": q_title, "tipe": "kisi_pilihan_ganda",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
                })
            else:
                pilihan_kanan = [
                    {
                        "huruf": ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})",
                        "teks": clean_option_prefix(col_text, ALPHABETS[i] if i < len(ALPHABETS) else "")
                    }
                    for i, col_text in enumerate(columns)
                ]
                current_section["soal"].append({
                    "item_id": str(item[0]),
                    "pertanyaan": q_title, "tipe": "menjodohkan",
                    "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": [], "kolom_header": columns, "kiri": rows, "kanan": pilihan_kanan, "poin": 0, "kunci_teks": ""
                })
            continue

        if len(item) >= 5 and item[4]:
            q_text = item[1] or ""
            item_imgs = list(html_img_map.get(item[0], []))

            choice_data = item[4][0]
            raw_opt_list = choice_data[1] if (len(choice_data) > 1 and choice_data[1]) else []
            n_opt_imgs = sum(1 for o in raw_opt_list if o and len(o) > 5 and o[5])

            # Ambil URL gambar pertanyaan utama dari blok HTML soal
            q_img_url = None
            # Di FB_PUBLIC_LOAD_DATA_, gambar pertanyaan tercatat pada indeks 9.
            has_q_img = len(item) > 9 and bool(item[9])
            if item_imgs and (has_q_img or len(item_imgs) > n_opt_imgs):
                q_img_url = item_imgs.pop(0)
            else:
                q_img_url = find_image_url(item)

            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else pending_standalone_img
            pending_standalone_img = None

            options = []
            for idx, opt_item in enumerate(raw_opt_list):
                if not opt_item:
                    continue
                raw_opt_text = str(opt_item[0] or "").strip()
                letter = ALPHABETS[idx] if idx < len(ALPHABETS) else f"({idx+1})"

                # Ambil URL gambar opsi jawaban jika opsi tersebut memiliki gambar
                if len(opt_item) > 5 and opt_item[5] and item_imgs:
                    opt_img_url = item_imgs.pop(0)
                else:
                    opt_img_url = find_image_url(opt_item)

                opt_img_bytes = download_image_to_stream(opt_img_url) if opt_img_url else None
                options.append({
                    "huruf": letter,
                    "teks": clean_option_prefix(raw_opt_text, letter),
                    "img_bytes": opt_img_bytes,
                    "ada_gambar": opt_img_bytes is not None,
                    "is_kunci": False
                })
            
            ada_opsi = len(options) > 0
            tipe_soal = "pg_kompleks" if (item_type == 4 and ada_opsi) else ("pilihan_ganda" if ada_opsi else "isian")

            current_section["soal"].append({
                    "item_id": str(item[0]),
                "pertanyaan": q_text, "tipe": tipe_soal,
                "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                "ada_opsi": ada_opsi, "opsi_bergambar": any(o["ada_gambar"] for o in options),
                "opsi": options, "baris": [], "kolom_header": [], "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
            })

    if current_section["soal"] or current_section["nama_bagian"]:
        raw_sections.append(current_section)

    daftar_bagian = filter_valid_exam_sections(raw_sections)

    return {
        "judul_form": form_title,
        "deskripsi_form": form_desc,
        "total_poin": 100,
        "daftar_bagian": daftar_bagian
    }


def build_context_for_doc(raw_context: Dict[str, Any], doc: DocxTemplate, mode_kunci: bool = False, judul_dokumen: str = "") -> Dict[str, Any]:
    ctx = copy.deepcopy(raw_context)
    ctx["mode_kunci"] = mode_kunci
    if judul_dokumen:
        ctx["judul_dokumen"] = judul_dokumen

    for bagian in ctx["daftar_bagian"]:
        for item in bagian["soal"]:
            if item.get("img_bytes"):
                item["gambar"] = InlineImage(doc, io.BytesIO(item["img_bytes"].getvalue()), width=Mm(65))
            else:
                item["gambar"] = None

            for opt in item.get("opsi", []):
                if opt.get("img_bytes"):
                    opt["gambar"] = InlineImage(doc, io.BytesIO(opt["img_bytes"].getvalue()), width=Mm(32))
                else:
                    opt["gambar"] = None

    ctx["daftar_pertanyaan"] = [q for sec in ctx["daftar_bagian"] for q in sec["soal"]]
    return ctx


# ==============================================================================
# 4b. KUNCI JAWABAN & SKOR (diisi guru / dimuat dari file JSON)
# ==============================================================================

KUNCI_FORMAT = "formprint-kunci-v1"
TIPE_LABEL = {
    "pilihan_ganda": "Pilihan Ganda",
    "pg_kompleks": "PG Kompleks",
    "benar_salah": "Benar / Salah",
    "kisi_pilihan_ganda": "Kisi Pilihan Ganda",
    "menjodohkan": "Menjodohkan",
    "isian": "Isian / Uraian",
}
TIPE_URUT = ["pilihan_ganda", "pg_kompleks", "benar_salah", "kisi_pilihan_ganda", "menjodohkan", "isian"]
TIPE_BERBARIS = ("benar_salah", "kisi_pilihan_ganda", "menjodohkan")


def _rows_of(q: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Baris yang diberi kunci & skor sendiri: premis (menjodohkan) atau baris kisi."""
    return (q.get("kiri") or []) if q.get("tipe") == "menjodohkan" else (q.get("baris") or [])


def _norm(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def question_key(q: Dict[str, Any]) -> str:
    """Kunci identitas soal yang stabil: ID butir Form, atau nomor jika tidak ada."""
    return str(q.get("item_id") or f"no{q.get('no_global')}")


def tipe_kunci(q: Dict[str, Any]) -> str:
    t = q.get("tipe", "isian")
    return t if t in TIPE_LABEL else "isian"


def empty_key_entry(q: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "no": q.get("no_global"),
        "tipe": tipe_kunci(q),
        "judul": (q.get("pertanyaan") or "")[:120],
        "poin": 0.0,
        "pg": [],
        "baris": [-1] * len(q.get("baris") or []),
        "jodoh": [""] * len(q.get("kiri") or []),
        "poin_baris": [0.0] * len(_rows_of(q)) if tipe_kunci(q) in TIPE_BERBARIS else [],
        "teks": "",
    }


def key_entry_from_question(q: Dict[str, Any]) -> Dict[str, Any]:
    """Kunci yang sudah ada di data Form (mode Forms API)."""
    e = empty_key_entry(q)
    e["poin"] = float(q.get("poin") or 0)
    t = e["tipe"]
    if t in TIPE_BERBARIS:
        e["poin_baris"] = [float(r.get("poin") or 0) for r in _rows_of(q)]
    if t in ("pilihan_ganda", "pg_kompleks"):
        e["pg"] = [o["huruf"] for o in q.get("opsi", []) if o.get("is_kunci")]
    elif t in ("benar_salah", "kisi_pilihan_ganda"):
        e["baris"] = [int(b.get("kunci_idx", -1)) for b in q.get("baris", [])]
    elif t == "menjodohkan":
        e["jodoh"] = [k.get("kunci_huruf", "") or "" for k in q.get("kiri", [])]
    else:
        e["teks"] = q.get("kunci_teks", "") or ""
    return e


def entry_has_key(e: Dict[str, Any]) -> bool:
    t = e.get("tipe")
    if t in ("pilihan_ganda", "pg_kompleks"):
        return bool(e.get("pg"))
    if t in ("benar_salah", "kisi_pilihan_ganda"):
        return bool(e.get("baris")) and all(i >= 0 for i in e["baris"])
    if t == "menjodohkan":
        return bool(e.get("jodoh")) and all(e["jodoh"])
    return bool((e.get("teks") or "").strip())


def init_key_state(raw_ctx: Dict[str, Any], previous: Optional[Dict[str, Any]] = None) -> Dict[str, Dict[str, Any]]:
    """
    Menyusun daftar kunci untuk semua soal. Urutan prioritas:
    kunci lama di sesi ini (jika soalnya sama) > kunci dari data Form > kosong.
    """
    previous = previous or {}
    state: Dict[str, Dict[str, Any]] = {}
    for sec in raw_ctx.get("daftar_bagian", []):
        for q in sec.get("soal", []):
            k = question_key(q)
            base = key_entry_from_question(q)
            old = previous.get(k)
            if old and old.get("tipe") == base["tipe"] and old.get("by_text") is not None:
                new = dict(base)
                new["poin"] = float(old.get("poin") or 0)
                new["teks"] = old.get("teks") or ""
                base = restore_text_keys(new, old["by_text"], q)
            elif old and old.get("tipe") == base["tipe"] \
                    and len(old.get("baris", [])) == len(base["baris"]) \
                    and len(old.get("jodoh", [])) == len(base["jodoh"]):
                base = dict(old)
                if len(base.get("poin_baris") or []) != len(_rows_of(q)) and base["tipe"] in TIPE_BERBARIS:
                    base["poin_baris"] = [0.0] * len(_rows_of(q))
            base["no"] = q.get("no_global")
            state[k] = attach_text_keys(base, q)
    return state


def _match_option_letters(q: Dict[str, Any], values: List[str]) -> List[str]:
    letters = []
    for v in values:
        for o in q.get("opsi", []):
            if _norm(v) == _norm(o["teks"]) or _norm(clean_option_prefix(str(v), o["huruf"])) == _norm(o["teks"]):
                if o["huruf"] not in letters:
                    letters.append(o["huruf"])
                break
    return letters


def attach_text_keys(e: Dict[str, Any], q: Dict[str, Any]) -> Dict[str, Any]:
    """
    Menyimpan kunci dalam bentuk TEKS (bukan nomor urut/huruf), agar tetap benar
    walaupun Form memakai 'acak urutan' sehingga urutan opsi/baris berubah setiap dibuka.
    """
    t = e.get("tipe")
    bt: Dict[str, Any] = {}
    if t in ("pilihan_ganda", "pg_kompleks"):
        m = {o["huruf"]: o["teks"] for o in q.get("opsi", [])}
        bt["pg"] = [m[h] for h in e.get("pg") or [] if h in m]
    elif t in TIPE_BERBARIS:
        rows = _rows_of(q)
        if t == "menjodohkan":
            m = {r["huruf"]: r["teks"] for r in q.get("kanan", [])}
            bt["jodoh"] = {r["teks"]: m.get(h, "") for r, h in zip(rows, e.get("jodoh") or []) if h}
        else:
            headers = q.get("kolom_header") or []
            bt["baris"] = {r["teks"]: headers[i] for r, i in zip(rows, e.get("baris") or []) if 0 <= i < len(headers)}
        bt["poin_baris"] = {r["teks"]: float(v or 0) for r, v in zip(rows, e.get("poin_baris") or [])}
    e["by_text"] = bt
    return e


def restore_text_keys(e: Dict[str, Any], bt: Dict[str, Any], q: Dict[str, Any]) -> Dict[str, Any]:
    """Kebalikan attach_text_keys: menyusun ulang kunci untuk urutan opsi/baris saat ini."""
    t = e.get("tipe")
    if t in ("pilihan_ganda", "pg_kompleks") and "pg" in bt:
        e["pg"] = _match_option_letters(q, bt["pg"])
    elif t in TIPE_BERBARIS:
        rows = _rows_of(q)
        rmap = {_norm(k): v for k, v in (bt.get("jodoh") or bt.get("baris") or {}).items()}
        pmap = {_norm(k): v for k, v in (bt.get("poin_baris") or {}).items()}
        if t == "menjodohkan":
            kmap = {_norm(r["teks"]): r["huruf"] for r in q.get("kanan", [])}
            e["jodoh"] = [kmap.get(_norm(rmap.get(_norm(r["teks"]), "")), "") for r in rows]
        else:
            headers = [_norm(h) for h in (q.get("kolom_header") or [])]
            e["baris"] = [headers.index(_norm(rmap[_norm(r["teks"])])) if _norm(rmap.get(_norm(r["teks"]), "")) in headers else -1
                          for r in rows]
        if pmap:
            e["poin_baris"] = [float(pmap.get(_norm(r["teks"]), 0) or 0) for r in rows]
            e["poin"] = sum(e["poin_baris"])
    e["by_text"] = bt
    return e


def merge_key_file(state: Dict[str, Dict[str, Any]], raw_ctx: Dict[str, Any], data: Dict[str, Any]) -> Tuple[int, int]:
    """
    Memuat file kunci ke state. Mendukung:
      - file dari skrip GAS   (format 'formprint-gas-kunci-v1')
      - file simpanan FormPrint (format 'formprint-kunci-v1')
    Mengembalikan (jumlah soal terisi, jumlah butir file yang tidak cocok).
    """
    questions = [q for sec in raw_ctx.get("daftar_bagian", []) for q in sec.get("soal", [])]
    by_id = {str(q.get("item_id")): q for q in questions if q.get("item_id")}
    by_title: Dict[str, Dict[str, Any]] = {}
    for q in questions:
        by_title.setdefault(_norm(q.get("pertanyaan")), q)

    terisi, tidak_cocok = 0, 0
    fmt = data.get("format", "")

    if fmt == KUNCI_FORMAT:
        for k, e in (data.get("kunci") or {}).items():
            q = by_id.get(k) or by_title.get(_norm(e.get("judul")))
            if not q:
                tidak_cocok += 1
                continue
            cur = state[question_key(q)]
            if e.get("tipe") != cur["tipe"]:
                tidak_cocok += 1
                continue
            new = dict(cur)
            new["poin"] = float(e.get("poin") or 0)
            if e.get("by_text") is not None:
                new["teks"] = str(e.get("teks") or "")
                state[question_key(q)] = restore_text_keys(new, e["by_text"], q)
                terisi += 1
                continue
            new["pg"] = [h for h in e.get("pg", []) if any(o["huruf"] == h for o in q.get("opsi", []))]
            if len(e.get("baris", [])) == len(cur["baris"]):
                new["baris"] = [int(i) for i in e["baris"]]
            if len(e.get("jodoh", [])) == len(cur["jodoh"]):
                new["jodoh"] = [str(h) for h in e["jodoh"]]
            if len(e.get("poin_baris") or []) == len(cur.get("poin_baris") or []):
                new["poin_baris"] = [float(v or 0) for v in e["poin_baris"]]
            new["teks"] = str(e.get("teks") or "")
            state[question_key(q)] = attach_text_keys(new, q)
            terisi += 1
        return terisi, tidak_cocok

    if fmt == "formprint-gas-kunci-v1":
        for it in data.get("items", []):
            q = by_id.get(str(it.get("id"))) or by_title.get(_norm(it.get("judul")))
            if not q:
                tidak_cocok += 1
                continue
            cur = state[question_key(q)]
            if it.get("poin"):
                cur["poin"] = float(it["poin"])
            if it.get("benar") and cur["tipe"] in ("pilihan_ganda", "pg_kompleks"):
                cur["pg"] = _match_option_letters(q, it["benar"])
            attach_text_keys(cur, q)
            terisi += 1
        return terisi, tidak_cocok

    raise ValueError("Format file kunci tidak dikenali.")


def export_key_file(state: Dict[str, Dict[str, Any]], judul_form: str,
                    raw_ctx: Optional[Dict[str, Any]] = None) -> bytes:
    if raw_ctx:
        state = copy.deepcopy(state)
        for sec in raw_ctx.get("daftar_bagian", []):
            for q in sec.get("soal", []):
                e = state.get(question_key(q))
                if e:
                    attach_text_keys(e, q)
    return json.dumps(
        {"format": KUNCI_FORMAT, "judul_form": judul_form, "kunci": state},
        ensure_ascii=False, indent=2,
    ).encode("utf-8")


def _fmt_poin(p: float) -> Any:
    return int(p) if float(p).is_integer() else round(p, 2)


def apply_keys_to_context(raw_ctx: Dict[str, Any], state: Dict[str, Dict[str, Any]], default_poin: float = 0) -> Dict[str, Any]:
    """Menyalin data Form dan mengisi kunci + skor dari state guru."""
    ctx = copy.deepcopy(raw_ctx)
    total = 0.0
    for sec in ctx.get("daftar_bagian", []):
        for q in sec.get("soal", []):
            e = state.get(question_key(q)) or empty_key_entry(q)
            t = e["tipe"]
            dflt = float(default_poin or 0)
            if t in TIPE_BERBARIS and _rows_of(q):
                # Skor per baris (premis / pernyataan); skor soal = jumlahnya
                pb = list(e.get("poin_baris") or [])
                pb += [0.0] * (len(_rows_of(q)) - len(pb))
                row_poin = [float(v or 0) or dflt for v in pb[: len(_rows_of(q))]]
                for r, v in zip(_rows_of(q), row_poin):
                    r["poin"] = _fmt_poin(v)
                poin = sum(row_poin) or float(e.get("poin") or 0)
            else:
                poin = float(e.get("poin") or 0) or dflt
            q["poin"] = _fmt_poin(poin)
            total += poin

            if t in ("pilihan_ganda", "pg_kompleks"):
                benar = set(e.get("pg") or [])
                for o in q.get("opsi", []):
                    o["is_kunci"] = o["huruf"] in benar
                pilihan = [o for o in q.get("opsi", []) if o["is_kunci"]]
                q["kunci_huruf"] = ", ".join(o["huruf"] for o in pilihan)
                q["kunci_teks"] = ", ".join(f'{o["huruf"]}. {o["teks"]}' for o in pilihan)

            elif t in ("benar_salah", "kisi_pilihan_ganda"):
                headers = q.get("kolom_header") or []
                for b, idx in zip(q.get("baris", []), e.get("baris") or []):
                    idx = int(idx)
                    b["kunci_idx"] = idx
                    b["is_benar"] = idx == 0
                    b["is_salah"] = idx == 1
                    b["kunci_huruf"] = headers[idx] if 0 <= idx < len(headers) else ""
                    b["kunci_teks"] = b["kunci_huruf"]

            elif t == "menjodohkan":
                kanan = {r["huruf"]: r["teks"] for r in q.get("kanan", [])}
                for k, h in zip(q.get("kiri", []), e.get("jodoh") or []):
                    k["kunci_huruf"] = h or ""
                    k["kunci_teks"] = kanan.get(h, "") if h else ""

            else:
                q["kunci_teks"] = (e.get("teks") or "").strip()

    ctx["total_poin"] = _fmt_poin(total)
    return ctx


def render_doc_bytes(tpl_bytes: bytes, ctx_raw: Dict[str, Any], mode_kunci: bool, judul_dokumen: str,
                     out_name: str) -> Dict[str, Optional[bytes]]:
    """Merender satu dokumen dari template (bytes) -> {'docx': ..., 'pdf': ...}."""
    with tempfile.TemporaryDirectory() as tmpdir:
        p_tpl = os.path.join(tmpdir, "tpl.docx")
        p_out = os.path.join(tmpdir, out_name)
        with open(p_tpl, "wb") as f:
            f.write(tpl_bytes)
        doc = DocxTemplate(p_tpl)
        ctx = build_context_for_doc(ctx_raw, doc, mode_kunci=mode_kunci, judul_dokumen=judul_dokumen)
        doc.render(ctx)
        lock_layout_prevent_page_split(doc.docx)
        doc.save(p_out)
        fix_header_first_page_only_file(p_out)
        with open(p_out, "rb") as f:
            docx_b = f.read()
        pdf_p = convert_docx_to_pdf(p_out, tmpdir)
        pdf_b = open(pdf_p, "rb").read() if (pdf_p and os.path.exists(pdf_p)) else None
    return {"docx": docx_b, "pdf": pdf_b}



# ==============================================================================
# 5. ANTARMUKA STREAMLIT (FULL CUSTOM CSS ALA TAILWIND / ILOVEPDF)
# ==============================================================================

st.set_page_config(
    page_title="FormPrint Studio - Konversi Google Form ke Naskah Cetak",
    page_icon="🖨️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    .stApp {
        background-color: #f8fafc;
    }

    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display: none;}

    .block-container {
        max-width: 820px !important;
        padding-top: 2.5rem !important;
        padding-bottom: 4rem !important;
    }

    [data-testid="stForm"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 2rem 2.2rem !important;
        box-shadow: 0 4px 20px -2px rgba(15, 23, 42, 0.05);
    }

    .stTextInput input {
        border-radius: 10px !important;
        border: 1px solid #cbd5e1 !important;
        padding: 0.65rem 0.9rem !important;
        font-size: 0.92rem !important;
        background-color: #f8fafc !important;
    }
    .stTextInput input:focus {
        border-color: #4f46e5 !important;
        box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.15) !important;
        background-color: #ffffff !important;
    }

    [data-testid="stFileUploader"] section {
        border-radius: 12px !important;
        border: 1.5px dashed #cbd5e1 !important;
        background-color: #f8fafc !important;
        padding: 1rem !important;
        transition: all 0.2s ease;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #6366f1 !important;
        background-color: #eef2ff !important;
    }

    [data-testid="stFormSubmitButton"] button {
        background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%) !important;
        color: white !important;
        font-weight: 700 !important;
        font-size: 0.98rem !important;
        padding: 0.75rem 1.5rem !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.25) !important;
        transition: all 0.2s ease !important;
        margin-top: 0.5rem !important;
    }
    [data-testid="stFormSubmitButton"] button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(79, 70, 229, 0.35) !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e2e8f0;
        padding: 6px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 0.9rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #4f46e5 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    .stTextInput label, .stSelectbox label, [data-testid="stFileUploader"] label {
        font-weight: 600 !important;
        color: #1e293b !important;
        font-size: 0.9rem !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="text-align: center; margin-bottom: 2rem;">
    <span style="background-color: #e0e7ff; color: #4338ca; padding: 6px 14px; border-radius: 9999px; font-size: 0.75rem; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase;">
        Otomatisasi Dokumen Ujian
    </span>
    <h1 style="color: #0f172a; font-size: 2.25rem; font-weight: 800; margin-top: 0.85rem; margin-bottom: 0.5rem; letter-spacing: -0.025em;">
        FormPrint Studio
    </h1>
    <p style="color: #64748b; font-size: 1rem; max-width: 580px; margin: 0 auto; line-height: 1.5;">
        Konversi <b>Google Form</b> menjadi <b>Naskah Soal</b>, <b>Lembar Jawaban Siswa</b>, dan <b>Kunci Jawaban</b> siap cetak (DOCX & PDF) dalam satu klik.
    </p>
</div>
""", unsafe_allow_html=True)

creds = get_google_credentials()
drive_folder_id = st.secrets.get("DRIVE_FOLDER_ID", "") if "DRIVE_FOLDER_ID" in st.secrets else os.environ.get("DRIVE_FOLDER_ID", "")

selected_drive_form_id = None
if creds and drive_folder_id:
    forms_list = list_forms_in_drive_folder(creds, drive_folder_id)
    if forms_list:
        st.markdown(f"""
        <div style="background-color: #ecfdf5; border: 1px solid #a7f3d0; color: #065f46; padding: 12px 16px; border-radius: 12px; font-size: 0.88rem; font-weight: 600; margin-bottom: 1rem;">
            📁 Terhubung ke Folder Google Drive Sekolah ({len(forms_list)} Google Form tersedia)
        </div>
        """, unsafe_allow_html=True)
        pilihan_nama = ["-- Gunakan Link Manual di Bawah --"] + [f"{f['name']}" for f in forms_list]
        pilihan_user = st.selectbox("Pilih Soal dari Folder Google Drive Sekolah:", pilihan_nama)
        if pilihan_user != "-- Gunakan Link Manual di Bawah --":
            idx_pilih = pilihan_nama.index(pilihan_user) - 1
            selected_drive_form_id = forms_list[idx_pilih]["id"]


# ------------------------------------------------------------------------------
# Panel Isi Kunci Jawaban & Skor (Tahap 2)
# ------------------------------------------------------------------------------
KOSONG = "—"


def _w(k: str, suffix: str) -> str:
    return f"kw_{k}_{suffix}"


def seed_key_widgets(state: Dict[str, Dict[str, Any]], raw_ctx: Dict[str, Any]) -> None:
    """Mengisi nilai awal semua isian di panel kunci dari state."""
    for sec in raw_ctx.get("daftar_bagian", []):
        for q in sec.get("soal", []):
            k = question_key(q)
            e = state.get(k) or empty_key_entry(q)
            t = e["tipe"]
            st.session_state[_w(k, "poin")] = float(e.get("poin") or 0)
            for i, v in enumerate(e.get("poin_baris") or []):
                st.session_state[_w(k, f"pr{i}")] = float(v or 0)
            if t == "pilihan_ganda":
                st.session_state[_w(k, "pg")] = (e.get("pg") or [KOSONG])[0]
            elif t == "pg_kompleks":
                st.session_state[_w(k, "pgk")] = list(e.get("pg") or [])
            elif t in ("benar_salah", "kisi_pilihan_ganda"):
                headers = q.get("kolom_header") or []
                for i, idx in enumerate(e.get("baris") or []):
                    st.session_state[_w(k, f"r{i}")] = headers[idx] if 0 <= idx < len(headers) else KOSONG
            elif t == "menjodohkan":
                for i, h in enumerate(e.get("jodoh") or []):
                    st.session_state[_w(k, f"j{i}")] = h or KOSONG
            else:
                st.session_state[_w(k, "teks")] = e.get("teks") or ""


def read_key_widgets(state: Dict[str, Dict[str, Any]], raw_ctx: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Membaca isian guru di panel kunci kembali ke state."""
    new_state = copy.deepcopy(state)
    for sec in raw_ctx.get("daftar_bagian", []):
        for q in sec.get("soal", []):
            k = question_key(q)
            e = new_state.setdefault(k, empty_key_entry(q))
            t = e["tipe"]
            e["poin"] = float(st.session_state.get(_w(k, "poin"), 0) or 0)
            if t in TIPE_BERBARIS:
                e["poin_baris"] = [float(st.session_state.get(_w(k, f"pr{i}"), 0) or 0) for i in range(len(_rows_of(q)))]
                e["poin"] = sum(e["poin_baris"])
            if t == "pilihan_ganda":
                v = st.session_state.get(_w(k, "pg"), KOSONG)
                e["pg"] = [] if v == KOSONG else [v]
            elif t == "pg_kompleks":
                e["pg"] = list(st.session_state.get(_w(k, "pgk"), []))
            elif t in ("benar_salah", "kisi_pilihan_ganda"):
                headers = q.get("kolom_header") or []
                e["baris"] = []
                for i in range(len(q.get("baris") or [])):
                    v = st.session_state.get(_w(k, f"r{i}"), KOSONG)
                    e["baris"].append(headers.index(v) if v in headers else -1)
            elif t == "menjodohkan":
                e["jodoh"] = []
                for i in range(len(q.get("kiri") or [])):
                    v = st.session_state.get(_w(k, f"j{i}"), KOSONG)
                    e["jodoh"].append("" if v == KOSONG else v)
            else:
                e["teks"] = st.session_state.get(_w(k, "teks"), "") or ""
            attach_text_keys(e, q)
    return new_state


def _ringkas(teks: str, n: int = 90) -> str:
    teks = re.sub(r"\s+", " ", teks or "").strip()
    return teks if len(teks) <= n else teks[: n - 1] + "…"


def render_key_panel() -> None:
    raw_ctx = st.session_state["raw_ctx"]
    state = st.session_state.get("kunci_state") or {}

    st.markdown("---")
    st.markdown("### 🔑 Isi Kunci Jawaban & Skor")

    if not st.session_state.get("tpl_ljk_bytes"):
        st.info("Unggah **Template Lembar Jawaban** lalu Generate ulang untuk membuat Kunci Jawaban Guru.")
        return

    st.caption(
        "Tahap 2: isi kunci per nomor, lalu tekan **Generate Kunci Jawaban**. "
        "Kunci PG/kotak centang/dropdown bisa dimuat otomatis dari file kunci skrip GAS."
    )

    # --- Muat file kunci (dari GAS atau simpanan sebelumnya) ---
    up = st.file_uploader("Muat file kunci (.json dari GAS atau simpanan FormPrint)", type=["json"], key="up_kunci")
    if up is not None:
        sig = f"{up.name}-{up.size}"
        if st.session_state.get("kunci_loaded_sig") != sig:
            try:
                data = json.loads(up.getvalue().decode("utf-8"))
                terisi, tidak_cocok = merge_key_file(state, raw_ctx, data)
                st.session_state["kunci_state"] = state
                seed_key_widgets(state, raw_ctx)
                st.session_state["kunci_loaded_sig"] = sig
                msg = f"✅ Kunci dimuat untuk {terisi} soal."
                if tidak_cocok:
                    msg += f" {tidak_cocok} butir di file tidak cocok dengan soal Form ini (dilewati)."
                st.session_state["kunci_load_msg"] = msg
            except Exception as e:
                st.session_state["kunci_load_msg"] = f"❌ File kunci tidak bisa dibaca: {e}"
                st.session_state["kunci_loaded_sig"] = sig
    if st.session_state.get("kunci_load_msg"):
        st.caption(st.session_state["kunci_load_msg"])

    # --- Kelompokkan soal per jenis ---
    groups: Dict[str, List[Dict[str, Any]]] = {t: [] for t in TIPE_URUT}
    for sec in raw_ctx.get("daftar_bagian", []):
        for q in sec.get("soal", []):
            groups[tipe_kunci(q)].append(q)

    with st.form("form_kunci"):
        st.number_input(
            "Skor default untuk soal/baris yang skornya masih 0",
            min_value=0.0, step=0.5, key="kw_default_poin",
            help="Dipakai untuk semua soal yang skornya belum diisi. Untuk menjodohkan, Benar/Salah, "
                 "dan kisi, skor default berlaku per baris. Biarkan 0 jika tidak dipakai.",
        )

        for t in TIPE_URUT:
            qs = groups[t]
            if not qs:
                continue
            n_kosong = sum(1 for q in qs if not entry_has_key(state.get(question_key(q)) or empty_key_entry(q)))
            label = f"{TIPE_LABEL[t]} — {len(qs)} soal" + (f" ({n_kosong} belum ada kunci)" if n_kosong else " ✓")
            with st.expander(label, expanded=n_kosong > 0):
                for q in qs:
                    k = question_key(q)
                    st.markdown(f"**No. {q.get('no_global')}** — {_ringkas(q.get('pertanyaan', ''))}")
                    if t in TIPE_BERBARIS:
                        st.caption("Skor diisi per baris; skor soal = jumlah skor baris.")
                        headers = q.get("kolom_header") or []
                        kanan = q.get("kanan") or []
                        m_kanan = {r["huruf"]: _ringkas(r["teks"], 40) for r in kanan}
                        for i, r in enumerate(_rows_of(q)):
                            c_kunci, c_poin = st.columns([4, 1])
                            with c_kunci:
                                if t == "menjodohkan":
                                    st.selectbox(
                                        f"{r.get('no_angka')}. {_ringkas(r.get('teks', ''), 70)}",
                                        [KOSONG] + [x["huruf"] for x in kanan], key=_w(k, f"j{i}"),
                                        format_func=lambda h, m=m_kanan: h if h == KOSONG else f"{h}. {m.get(h, '')}",
                                    )
                                else:
                                    lbl = f"({r.get('sub_no')})" if t == "benar_salah" else f"{r.get('no_angka')}."
                                    st.selectbox(f"{lbl} {_ringkas(r.get('teks', ''), 70)}", [KOSONG] + headers, key=_w(k, f"r{i}"))
                            with c_poin:
                                st.number_input("Skor", min_value=0.0, step=0.5, key=_w(k, f"pr{i}"))
                        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
                        continue
                    c_kunci, c_poin = st.columns([4, 1])
                    with c_poin:
                        st.number_input("Skor", min_value=0.0, step=0.5, key=_w(k, "poin"))
                    with c_kunci:
                        opsi = q.get("opsi", [])
                        teks_opsi = {o["huruf"]: _ringkas(o["teks"], 50) for o in opsi}
                        if t == "pilihan_ganda":
                            st.selectbox(
                                "Kunci", [KOSONG] + [o["huruf"] for o in opsi], key=_w(k, "pg"),
                                format_func=lambda h, m=teks_opsi: h if h == KOSONG else f"{h}. {m.get(h, '')}",
                            )
                        elif t == "pg_kompleks":
                            st.multiselect(
                                "Kunci (boleh lebih dari satu)", [o["huruf"] for o in opsi], key=_w(k, "pgk"),
                                format_func=lambda h, m=teks_opsi: f"{h}. {m.get(h, '')}",
                            )
                        else:
                            st.text_input("Kunci / pedoman jawaban", key=_w(k, "teks"))
                    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        submit_kunci = st.form_submit_button("🔑 Generate Kunci Jawaban", use_container_width=True, type="primary")

    if submit_kunci:
        new_state = read_key_widgets(state, raw_ctx)
        st.session_state["kunci_state"] = new_state
        ctx = apply_keys_to_context(raw_ctx, new_state, st.session_state.get("kw_default_poin", 0))
        kosong = [str(v["no"]) for v in sorted(new_state.values(), key=lambda v: v["no"] or 0) if not entry_has_key(v)]
        with st.spinner("Menyusun Kunci Jawaban Guru..."):
            try:
                st.session_state["hasil_files"]["kunci_guru"] = render_doc_bytes(
                    st.session_state["tpl_ljk_bytes"], ctx, True, "KUNCI JAWABAN & PEDOMAN SKOR", "Kunci_Jawaban_Guru.docx"
                )
                st.session_state["kunci_info"] = {"kosong": kosong, "total": ctx["total_poin"]}
            except Exception as e:
                st.error(f"Gagal membuat Kunci Jawaban: {e}")
                return
        st.rerun()

    info = st.session_state.get("kunci_info")
    if info:
        if info["kosong"]:
            st.warning("Nomor tanpa kunci (tercetak \"Sesuai pedoman penskoran guru\"): " + ", ".join(info["kosong"]))
        st.success(f"Kunci Jawaban Guru sudah dibuat (tab 🔑 di atas). Total skor: {info['total']}")

    st.download_button(
        "💾 Simpan kunci (.json) untuk dipakai lagi",
        data=export_key_file(st.session_state.get("kunci_state") or {}, raw_ctx.get("judul_form", ""), raw_ctx),
        file_name=f"kunci_{st.session_state.get('judul_naskah', 'form')}.json",
        mime="application/json",
        use_container_width=True,
        key="dl_kunci_json",
    )


with st.form("form_utama"):
    st.markdown("""
    <div style="margin-bottom: 0.5rem;">
        <span style="font-size: 1.05rem; font-weight: 700; color: #0f172a;">Konfigurasi Dokumen</span>
        <p style="font-size: 0.8rem; color: #64748b; margin-top: 2px;">Bagian 1 (Identitas) dan Bagian Akhir (Konfirmasi Selesai) otomatis dilewati.</p>
    </div>
    """, unsafe_allow_html=True)

    form_url = st.text_input(
        "1. Link Google Form",
        placeholder="https://docs.google.com/forms/d/.../edit atau /viewform"
    )

    col_tpl1, col_tpl2 = st.columns(2)
    with col_tpl1:
        tpl_naskah = st.file_uploader("2. Template Naskah Soal (.docx)", type=["docx"])
    with col_tpl2:
        tpl_ljk = st.file_uploader("3. Template Lembar Jawaban (.docx)", type=["docx"])

    submit_btn = st.form_submit_button("✨ Generate Dokumen Siap Cetak", use_container_width=True, type="primary")

with st.expander("📖 Lihat Cara Pakai & Informasi Filter Otomatis"):
    st.markdown("""
    * **Filter Bagian Otomatis:** Bagian 1 (Identitas) dan Bagian Terakhir (*"SELAMAT... KAMU SUDAH DI TAHAP AKHIR / KLIK SELESAI"*) otomatis dibuang tanpa memengaruhi bagian soal utama.
    * **Tanpa Penomoran Otomatis:** Teks pertanyaan dicetak sesuai isi Form tanpa tambahan nomor urut otomatis.
    * **Header Halaman 1 Saja:** Kop Surat di Header Word otomatis dikunci agar hanya tampil di halaman pertama.
    * **Kunci Jawaban 2 Tahap:** Setelah Generate, isi kunci & skor di panel **🔑 Isi Kunci Jawaban & Skor**, lalu tekan *Generate Kunci Jawaban*. Kunci PG/kotak centang/dropdown bisa dimuat dari file kunci skrip GAS. Simpan kunci (.json) agar tidak perlu mengetik ulang.
    """)

if submit_btn:
    target_input = selected_drive_form_id or form_url.strip()
    if not target_input:
        st.error("Mohon pilih Google Form dari daftar atau masukkan link Google Form terlebih dahulu.")
    elif not tpl_naskah and not tpl_ljk:
        st.error("Mohon unggah minimal satu file Template Word (Template Naskah Soal atau Template Lembar Jawaban).")
    else:
        with st.spinner("Sedang mengekstrak bagian soal, mengunduh gambar, & menyusun dokumen..."):
            try:
                IMAGE_WARNINGS.clear()
                IMAGE_DEBUG.clear()
                form_id = extract_form_id_from_url(target_input)
                if creds and form_id:
                    raw_ctx = extract_via_forms_api(form_id, creds)
                    st.session_state["mode_api"] = True
                else:
                    raw_ctx = extract_via_public_html(target_input)
                    st.session_state["mode_api"] = False

                judul_bersih = re.sub(r"[^a-zA-Z0-9_-]", "_", raw_ctx["judul_form"])
                st.session_state["judul_naskah"] = judul_bersih
                st.session_state["hasil_files"] = {}

                with tempfile.TemporaryDirectory() as tmpdir:
                    if tpl_naskah:
                        p_tpl_soal = os.path.join(tmpdir, "tpl_soal.docx")
                        p_out_soal = os.path.join(tmpdir, "Naskah_Soal.docx")
                        with open(p_tpl_soal, "wb") as f:
                            f.write(tpl_naskah.getbuffer())

                        strip_legacy_numbering_tags(p_tpl_soal)

                        doc_soal = DocxTemplate(p_tpl_soal)
                        ctx_soal = build_context_for_doc(raw_ctx, doc_soal, mode_kunci=False)
                        doc_soal.render(ctx_soal)
                        lock_layout_prevent_page_split(doc_soal.docx)
                        doc_soal.save(p_out_soal)
                        fix_header_first_page_only_file(p_out_soal)

                        with open(p_out_soal, "rb") as f:
                            docx_b = f.read()
                        pdf_p = convert_docx_to_pdf(p_out_soal, tmpdir)
                        pdf_b = open(pdf_p, "rb").read() if (pdf_p and os.path.exists(pdf_p)) else None

                        st.session_state["hasil_files"]["naskah"] = {"docx": docx_b, "pdf": pdf_b}

                    if tpl_ljk:
                        p_tpl_ljk = os.path.join(tmpdir, "tpl_ljk.docx")
                        with open(p_tpl_ljk, "wb") as f:
                            f.write(tpl_ljk.getbuffer())

                        # 1. Lembar Jawaban Siswa (Kosong)
                        p_out_ljk = os.path.join(tmpdir, "Lembar_Jawaban_Siswa.docx")
                        doc_ljk = DocxTemplate(p_tpl_ljk)
                        ctx_ljk = build_context_for_doc(raw_ctx, doc_ljk, mode_kunci=False, judul_dokumen="LEMBAR JAWABAN SISWA")
                        doc_ljk.render(ctx_ljk)
                        lock_layout_prevent_page_split(doc_ljk.docx)
                        doc_ljk.save(p_out_ljk)
                        fix_header_first_page_only_file(p_out_ljk)

                        with open(p_out_ljk, "rb") as f:
                            ljk_docx = f.read()
                        ljk_pdf_p = convert_docx_to_pdf(p_out_ljk, tmpdir)
                        ljk_pdf = open(ljk_pdf_p, "rb").read() if (ljk_pdf_p and os.path.exists(ljk_pdf_p)) else None
                        st.session_state["hasil_files"]["ljk_siswa"] = {"docx": ljk_docx, "pdf": ljk_pdf}

                        # Kunci Jawaban Guru dibuat di Tahap 2 (panel "Isi Kunci Jawaban & Skor")
                        st.session_state["tpl_ljk_bytes"] = bytes(tpl_ljk.getbuffer())
                    else:
                        st.session_state["tpl_ljk_bytes"] = None

                st.session_state["raw_ctx"] = raw_ctx
                st.session_state["kunci_state"] = init_key_state(raw_ctx, st.session_state.get("kunci_state"))
                st.session_state["kunci_info"] = None
                st.session_state["kunci_loaded_sig"] = None
                st.session_state["kunci_load_msg"] = None
                seed_key_widgets(st.session_state["kunci_state"], raw_ctx)
                st.session_state["generated"] = True
                st.session_state["image_warnings"] = list(dict.fromkeys(IMAGE_WARNINGS))
                st.session_state["image_debug"] = dict(IMAGE_DEBUG)
                st.session_state["image_debug"]["soal_dengan_gambar"] = sum(
                    1 for sec in raw_ctx["daftar_bagian"] for q in sec["soal"]
                    if q.get("img_bytes") or any(o.get("img_bytes") for o in q.get("opsi", []))
                )

            except Exception as e:
                st.session_state["generated"] = False
                st.error(f"Gagal memproses dokumen: {str(e)}")

# Panel Hasil & Pratinjau
if st.session_state.get("generated"):
    hasil = st.session_state.get("hasil_files", {})
    judul = st.session_state.get("judul_naskah", "Dokumen")

    st.markdown("""
    <div style="background-color: #ecfdf5; border: 1px solid #6ee7b7; color: #065f46; padding: 14px 18px; border-radius: 12px; font-weight: 600; font-size: 0.92rem; margin-top: 1.5rem; margin-bottom: 1rem;">
        ✅ Dokumen siap cetak berhasil dibuat! Pilih tab di bawah untuk mengunduh atau melihat pratinjau.
    </div>
    """, unsafe_allow_html=True)

    img_warn = st.session_state.get("image_warnings") or []
    if img_warn:
        st.warning("⚠️ Sebagian gambar mungkin tidak ikut:\n\n" + "\n".join(f"- {w}" for w in img_warn))

    dbg = st.session_state.get("image_debug") or {}
    if dbg:
        with st.expander("🔍 Diagnosa gambar"):
            st.json({k: v for k, v in dbg.items() if not k.startswith("_")})
            if dbg.get("_html"):
                st.download_button(
                    "⬇️ Unduh HTML Form (untuk analisis)",
                    data=dbg["_html"].encode("utf-8"),
                    file_name="form_page.html",
                    mime="text/html",
                    key="dl_debug_html",
                )

    tab_labels = []
    if "naskah" in hasil:
        tab_labels.append("📄 Naskah Soal")
    if "ljk_siswa" in hasil:
        tab_labels.append("📝 Lembar Jawaban Siswa")
    if "kunci_guru" in hasil:
        tab_labels.append("🔑 Kunci Jawaban Guru")

    tabs = st.tabs(tab_labels)
    tab_idx = 0

    for key_id, nama_file in [("naskah", "Naskah_Soal"), ("ljk_siswa", "Lembar_Jawaban"), ("kunci_guru", "Kunci_Jawaban")]:
        if key_id in hasil:
            with tabs[tab_idx]:
                item_data = hasil[key_id]
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        label=f"⬇️ Unduh DOCX ({nama_file})",
                        data=item_data["docx"],
                        file_name=f"{nama_file}_{judul}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key=f"dl_docx_{key_id}"
                    )
                with c2:
                    if item_data["pdf"]:
                        st.download_button(
                            label=f"🖨️ Unduh PDF ({nama_file})",
                            data=item_data["pdf"],
                            file_name=f"{nama_file}_{judul}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            type="primary",
                            key=f"dl_pdf_{key_id}"
                        )
                if item_data["pdf"]:
                    pdf_viewer(input=item_data["pdf"], height=680, key=f"pv_{key_id}")
            tab_idx += 1

    if st.session_state.get("raw_ctx"):
        render_key_panel()

st.markdown("""
<div style="text-align: center; margin-top: 3.5rem; color: #94a3b8; font-size: 0.78rem;">
    FormPrint Studio • Sistem Otomatisasi Naskah Ujian & Lembar Jawaban
</div>
""", unsafe_allow_html=True)
