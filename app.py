import copy
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
from typing import Any, Dict, List, Optional

import requests
import streamlit as st
from docx import Document
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate, InlineImage
from streamlit_pdf_viewer import pdf_viewer

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
# 1. UTILITAS SISTEM, HEADER HALAMAN 1, & PENGUNCI HALAMAN
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
    """
    Menghapus tag penomoran otomatis ({{ item.no_global }}., {{ b.sub_no }}., dll.)
    secara otomatis jika pengguna mengunggah file template versi lama.
    """
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


def configure_header_first_page_only(docx_obj) -> None:
    """
    Memastikan jika template memiliki Header (Kop Surat di area Header Word),
    header tersebut hanya muncul di Halaman 1 saja dan tidak berulang di halaman 2 dst.
    """
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

    # Cek apakah header default berisi teks, tabel, atau gambar/logo
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
            )
        if not first_has_content:
            if first_hdr_ref is not None:
                sectPr0.remove(first_hdr_ref)
            default_hdr_ref.set(qn("w:type"), "first")
        else:
            # Jika first_page_header sudah ada isinya, kosongkan default header agar halaman 2+ bersih
            for p in list(sec0.header.paragraphs):
                p.text = ""
            for tbl in list(sec0.header.tables):
                tbl._element.getparent().remove(tbl._element)

    # Aktifkan Different First Page (<w:titlePg/>) pada section pertama
    if sectPr0.find(qn("w:titlePg")) is None:
        sectPr0.append(OxmlElement("w:titlePg"))

    # Untuk section lanjutan (misal section 2 kolom setelah Continuous Break),
    # matikan warisan header agar di halaman 2 dst benar-benar bersih tanpa header
    for sec in docx_obj.sections[1:]:
        sPr = sec._sectPr
        titlePg = sPr.find(qn("w:titlePg"))
        if titlePg is not None:
            sPr.remove(titlePg)
        sec.header.is_linked_to_previous = False
        for p in list(sec.header.paragraphs):
            p.text = ""
        for tbl in list(sec.header.tables):
            tbl._element.getparent().remove(tbl._element)


def lock_layout_prevent_page_split(docx_obj) -> None:
    """Mengunci paragraf soal, gambar, dan baris tabel agar tidak terpotong lintas halaman."""
    configure_header_first_page_only(docx_obj)

    for p in docx_obj.paragraphs:
        text = p.text.strip()
        p.paragraph_format.keep_together = True
        if "w:drawing" in p._p.xml or text.startswith("Petunjuk:"):
            p.paragraph_format.keep_with_next = True

    for table in docx_obj.tables:
        for row_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            # Matikan tblHeader (Repeat Header Rows) agar header tabel tidak berulang di halaman baru
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
    """Mencegah duplikasi huruf opsi (misal 'A. A. Jawaban') jika guru sudah mengetik huruf di Form."""
    t = text.strip()
    pattern = rf"^(\({expected_letter}\)|{expected_letter}[\.\)])\s+"
    return re.sub(pattern, "", t, flags=re.IGNORECASE).strip()


def download_image_to_stream(img_url: str) -> Optional[io.BytesIO]:
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(img_url, headers=headers, timeout=8)
        if resp.status_code == 200:
            return io.BytesIO(resp.content)
    except Exception:
        pass
    return None


def find_image_url(data: Any) -> Optional[str]:
    if isinstance(data, str):
        if "googleusercontent.com" in data or (
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


# ==============================================================================
# 2. KONEKSI GOOGLE SERVICE ACCOUNT (OPSI C - FOLDER DRIVE & KUNCI OTOMATIS)
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
# 3. EKSTRAKSI VIA GOOGLE FORMS API (LOMPATI BAGIAN 1 IDENTITAS)
# ==============================================================================

def extract_via_forms_api(form_id: str, creds) -> Dict[str, Any]:
    forms_service = build("forms", "v1", credentials=creds)
    form_data = forms_service.forms().get(formId=form_id).execute()

    info = form_data.get("info", {})
    form_title = info.get("title") or "Naskah Soal Ujian"
    form_desc = info.get("description") or ""
    items = form_data.get("items", [])

    # Deteksi apakah form memiliki lebih dari 1 bagian (ada pageBreakItem)
    has_multiple_sections = any("pageBreakItem" in it for it in items)
    # Jika punya banyak bagian, lompati Bagian 1 (Identitas) dan mulai di Bagian 2
    in_first_section = has_multiple_sections

    daftar_bagian = []
    current_section = {
        "nama_bagian": "",
        "petunjuk": "",
        "soal": []
    }
    global_counter = 1
    total_poin = 0

    for item in items:
        # A. Deteksi Pembatas Bagian (Page Break -> Masuk ke Bagian 2, 3, dst.)
        if "pageBreakItem" in item:
            if in_first_section:
                # Baru saja melewati Bagian 1 (Identitas), sekarang masuk Bagian 2
                in_first_section = False
            else:
                if current_section["soal"]:
                    daftar_bagian.append(current_section)

            current_section = {
                "nama_bagian": item.get("title") or "",
                "petunjuk": item.get("description") or "",
                "soal": []
            }
            continue

        # Lompati seluruh item jika masih berada di Bagian 1 (Identitas)
        if in_first_section:
            continue

        # B. Deteksi Soal Kisi / Grid (questionGroupItem)
        if "questionGroupItem" in item:
            group = item["questionGroupItem"]
            q_title = item.get("title") or ""
            q_img_url = group.get("image", {}).get("contentUri")
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else None

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

            total_poin += group_poin
            nomor_di_bagian = len(current_section["soal"]) + 1

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
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "pg_kompleks", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": True, "opsi_bergambar": False, "opsi": pgk_opts,
                    "baris": [], "kolom_header": [], "kiri": [], "kanan": [],
                    "poin": group_poin, "kunci_teks": ""
                })
            elif len(columns) == 2 and any(k in cols_lower for k in ["benar", "salah", "b", "s", "true", "false", "ya", "tidak"]):
                current_section["soal"].append({
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "benar_salah", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [],
                    "poin": group_poin, "kunci_teks": ""
                })
            elif len(columns) <= 4 and all(c.lower() in opsi_matriks_standar for c in columns):
                current_section["soal"].append({
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "kisi_pilihan_ganda", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
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
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "menjodohkan", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": [], "kolom_header": columns, "kiri": rows, "kanan": pilihan_kanan,
                    "poin": group_poin, "kunci_teks": ""
                })
            global_counter += 1
            continue

        # C. Deteksi Soal Tunggal (questionItem)
        if "questionItem" in item:
            q_item = item["questionItem"]
            q_obj = q_item.get("question", {})
            q_text = item.get("title") or ""

            q_img_url = q_item.get("image", {}).get("contentUri")
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else None

            grading = q_obj.get("grading", {})
            poin = grading.get("pointValue", 0)
            total_poin += poin
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

            nomor_di_bagian = len(current_section["soal"]) + 1
            current_section["soal"].append({
                "no": nomor_di_bagian,
                "no_global": global_counter,
                "pertanyaan": q_text,
                "tipe": tipe_soal,
                "img_bytes": q_img_bytes,
                "ada_gambar": q_img_bytes is not None,
                "ada_opsi": len(options) > 0,
                "opsi_bergambar": any(o["ada_gambar"] for o in options),
                "opsi": options,
                "baris": [],
                "kolom_header": [],
                "kiri": [],
                "kanan": [],
                "poin": poin,
                "kunci_teks": ", ".join(correct_vals)
            })
            global_counter += 1

    if current_section["soal"]:
        daftar_bagian.append(current_section)

    return {
        "judul_form": form_title,
        "deskripsi_form": form_desc,
        "total_poin": total_poin if total_poin > 0 else 100,
        "daftar_bagian": daftar_bagian
    }


# ==============================================================================
# 4. EKSTRAKSI VIA LINK PUBLIK (LOMPATI BAGIAN 1 IDENTITAS)
# ==============================================================================

def extract_via_public_html(form_url: str) -> Dict[str, Any]:
    clean_url = form_url.strip().split("?")[0]
    if not clean_url.endswith("/viewform"):
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

    # Cek apakah form memiliki lebih dari 1 bagian (Tipe 8 = Section Break)
    has_multiple_sections = any(it and len(it) >= 4 and it[3] == 8 for it in raw_items)
    # Jika punya banyak bagian, lompati Bagian 1 (Identitas) dan mulai di Bagian 2
    in_first_section = has_multiple_sections

    daftar_bagian = []
    current_section = {"nama_bagian": "", "petunjuk": "", "soal": []}
    global_counter = 1

    for item in raw_items:
        if not item or len(item) < 4:
            continue
        item_type = item[3]

        # A. Pembatas Bagian (Tipe 8 -> Masuk Bagian 2, 3, dst.)
        if item_type == 8:
            if in_first_section:
                in_first_section = False
            else:
                if current_section["soal"]:
                    daftar_bagian.append(current_section)

            current_section = {
                "nama_bagian": item[1] or "",
                "petunjuk": item[2] or "",
                "soal": []
            }
            continue

        # Lompati semua item pada Bagian 1 (Identitas)
        if in_first_section:
            continue

        # B. Soal Kisi / Grid (Tipe 7)
        if item_type == 7 and len(item) >= 5 and item[4]:
            q_title = item[1] or ""
            rows_data = item[4] or []
            q_img_url = find_image_url(item[1:4])
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else None

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
                    "is_benar": False,
                    "is_salah": False,
                    "kunci_idx": -1,
                    "kunci_huruf": "",
                    "kunci_teks": ""
                })

            cols_lower = [c.lower() for c in columns]
            opsi_matriks_standar = {"a", "b", "c", "d", "1", "2", "3", "4"}
            nomor_di_bagian = len(current_section["soal"]) + 1

            if len(columns) == 1:
                pgk_opts = []
                for i, b in enumerate(rows):
                    letter = ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})"
                    pgk_opts.append({
                        "huruf": letter,
                        "teks": clean_option_prefix(b["teks"], letter),
                        "img_bytes": None,
                        "ada_gambar": False,
                        "is_kunci": False
                    })
                current_section["soal"].append({
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "pg_kompleks", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": True, "opsi_bergambar": False, "opsi": pgk_opts,
                    "baris": [], "kolom_header": [], "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
                })
            elif len(columns) == 2 and any(k in cols_lower for k in ["benar", "salah", "b", "s", "true", "false", "ya", "tidak"]):
                current_section["soal"].append({
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "benar_salah", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": rows, "kolom_header": columns, "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
                })
            elif len(columns) <= 4 and all(c.lower() in opsi_matriks_standar for c in columns):
                current_section["soal"].append({
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "kisi_pilihan_ganda", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
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
                    "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_title,
                    "tipe": "menjodohkan", "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                    "ada_opsi": False, "opsi_bergambar": False, "opsi": [],
                    "baris": [], "kolom_header": columns, "kiri": rows, "kanan": pilihan_kanan, "poin": 0, "kunci_teks": ""
                })
            global_counter += 1
            continue

        # C. Soal Standar (PG, PG Kompleks, Isian)
        if len(item) >= 5 and item[4]:
            q_text = item[1] or ""
            q_img_url = find_image_url(item[1:4])
            q_img_bytes = download_image_to_stream(q_img_url) if q_img_url else None

            options = []
            choice_data = item[4][0]
            if len(choice_data) > 1 and choice_data[1]:
                for idx, opt_item in enumerate(choice_data[1]):
                    if not opt_item:
                        continue
                    raw_opt_text = str(opt_item[0] or "").strip()
                    letter = ALPHABETS[idx] if idx < len(ALPHABETS) else f"({idx+1})"
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
            nomor_di_bagian = len(current_section["soal"]) + 1

            current_section["soal"].append({
                "no": nomor_di_bagian, "no_global": global_counter, "pertanyaan": q_text,
                "tipe": tipe_soal, "img_bytes": q_img_bytes, "ada_gambar": q_img_bytes is not None,
                "ada_opsi": ada_opsi, "opsi_bergambar": any(o["ada_gambar"] for o in options),
                "opsi": options, "baris": [], "kolom_header": [], "kiri": [], "kanan": [], "poin": 0, "kunci_teks": ""
            })
            global_counter += 1

    if current_section["soal"]:
        daftar_bagian.append(current_section)

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

with st.form("form_utama"):
    st.markdown("""
    <div style="margin-bottom: 0.5rem;">
        <span style="font-size: 1.05rem; font-weight: 700; color: #0f172a;">Konfigurasi Dokumen</span>
        <p style="font-size: 0.8rem; color: #64748b; margin-top: 2px;">Bagian 1 (Identitas) di Google Form otomatis dilompati. Ekstraksi dimulai dari Bagian 2.</p>
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

with st.expander("📖 Lihat Cara Pakai & Informasi Template"):
    st.markdown("""
    * **Otomatis Lompati Identitas:** Bagian 1 pada Google Form diabaikan dan aplikasi langsung memproses mulai dari Bagian 2.
    * **Tanpa Penomoran Otomatis:** Teks pertanyaan dicetak sesuai isi Form tanpa tambahan nomor urut otomatis.
    * **Header Halaman 1 Saja:** Jika template memiliki Header Word (Kop Surat), sistem otomatis menguncinya agar hanya tampil di halaman pertama.
    """)

if submit_btn:
    target_input = selected_drive_form_id or form_url.strip()
    if not target_input:
        st.error("Mohon pilih Google Form dari daftar atau masukkan link Google Form terlebih dahulu.")
    elif not tpl_naskah and not tpl_ljk:
        st.error("Mohon unggah minimal satu file Template Word (Template Naskah Soal atau Template Lembar Jawaban).")
    else:
        with st.spinner("Sedang mengekstrak soal mulai Bagian 2, mengunduh gambar, & menyusun dokumen..."):
            try:
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

                        # Bersihkan tag penomoran otomatis jika pengguna memakai file template versi lama
                        strip_legacy_numbering_tags(p_tpl_soal)

                        doc_soal = DocxTemplate(p_tpl_soal)
                        ctx_soal = build_context_for_doc(raw_ctx, doc_soal, mode_kunci=False)
                        doc_soal.render(ctx_soal)
                        lock_layout_prevent_page_split(doc_soal.docx)
                        doc_soal.save(p_out_soal)

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

                        with open(p_out_ljk, "rb") as f:
                            ljk_docx = f.read()
                        ljk_pdf_p = convert_docx_to_pdf(p_out_ljk, tmpdir)
                        ljk_pdf = open(ljk_pdf_p, "rb").read() if (ljk_pdf_p and os.path.exists(ljk_pdf_p)) else None
                        st.session_state["hasil_files"]["ljk_siswa"] = {"docx": ljk_docx, "pdf": ljk_pdf}

                        # 2. Lembar Kunci Jawaban Guru (Terisi)
                        p_out_kunci = os.path.join(tmpdir, "Kunci_Jawaban_Guru.docx")
                        doc_kunci = DocxTemplate(p_tpl_ljk)
                        ctx_kunci = build_context_for_doc(raw_ctx, doc_kunci, mode_kunci=True, judul_dokumen="KUNCI JAWABAN & PEDOMAN SKOR")
                        doc_kunci.render(ctx_kunci)
                        lock_layout_prevent_page_split(doc_kunci.docx)
                        doc_kunci.save(p_out_kunci)

                        with open(p_out_kunci, "rb") as f:
                            kunci_docx = f.read()
                        kunci_pdf_p = convert_docx_to_pdf(p_out_kunci, tmpdir)
                        kunci_pdf = open(kunci_pdf_p, "rb").read() if (kunci_pdf_p and os.path.exists(kunci_pdf_p)) else None
                        st.session_state["hasil_files"]["kunci_guru"] = {"docx": kunci_docx, "pdf": kunci_pdf}

                st.session_state["generated"] = True

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

st.markdown("""
<div style="text-align: center; margin-top: 3.5rem; color: #94a3b8; font-size: 0.78rem;">
    FormPrint Studio • Sistem Otomatisasi Naskah Ujian & Lembar Jawaban
</div>
""", unsafe_allow_html=True)
