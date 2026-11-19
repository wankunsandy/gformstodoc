"""
FormPrint Studio – halaman "Soal → Google Form"
Membaca naskah soal DOCX, menampilkan pratinjau yang bisa disunting, mengisi kunci & skor,
lalu menghasilkan paket soal (.json) untuk skrip GAS "FormPrint – Buat Form".
"""
import json
import re

import streamlit as st

from soal_parser import build_paket, default_key_index, parse_docx

st.set_page_config(page_title="FormPrint Studio - Soal ke Google Form", page_icon="📝", layout="centered")

TIPE_LABEL = {
    "pilihan_ganda": "Pilihan Ganda",
    "pg_kompleks": "PG Kompleks",
    "benar_salah": "Benar / Salah",
    "isian": "Isian Singkat",
    "menjodohkan": "Menjodohkan",
    "uraian": "Uraian",
}
KOSONG = "—"

# ------------------------------------------------------------------------------
# PENGATURAN SEKOLAH — ubah setiap tahun ajaran
# ------------------------------------------------------------------------------
# Rombel per jenjang (huruf kelas yang tersedia tahun ajaran ini)
KELAS_PER_JENJANG = {
    "7": "ABCDEFGH",   # 7A–7H
    "8": "ABCDEFGHI",  # 8A–8I
    "9": "ABCDEFG",    # 9A–9G
}


def daftar_kelas(jenjang: str) -> list:
    return [f"{jenjang}{h}" for h in KELAS_PER_JENJANG.get(str(jenjang), "")]


def k(no, name):
    return f"sf_{no}_{name}"


def seed(parsed):
    """Nilai awal semua isian dari hasil pembacaan dokumen."""
    for sec in parsed["bagian"]:
        st.session_state[f"sf_sec_{sec['judul']}_pet"] = sec.get("petunjuk", "")
        for s in sec["soal"]:
            no = s["no"]
            st.session_state[k(no, "teks")] = s["teks"]
            for i, o in enumerate(s["opsi"]):
                st.session_state[k(no, f"o{i}")] = o["teks"]
            st.session_state[k(no, "poin")] = int(s.get("poin") or 0)
            kunci = default_key_index(s)
            if s["tipe"] in ("pilihan_ganda", "benar_salah"):
                st.session_state[k(no, "kunci")] = s["opsi"][kunci[0]]["huruf"] if kunci else KOSONG
            elif s["tipe"] == "pg_kompleks":
                st.session_state[k(no, "kunci_m")] = [s["opsi"][i]["huruf"] for i in kunci]
            elif s["tipe"] == "menjodohkan":
                st.session_state[k(no, "kunci")] = KOSONG

    judul = " – ".join(x for x in [parsed["judul_baris"][0] if parsed["judul_baris"] else "Soal",
                                     parsed.get("mapel", ""),
                                     f"Kelas {parsed['kelas']}" if parsed.get("kelas") else ""] if x)
    st.session_state["sf_judul"] = judul
    st.session_state["sf_desk"] = "\n".join(f"{i + 1}. {t}" for i, t in enumerate(parsed["petunjuk_umum"]))
    st.session_state["sf_identitas"] = "Nama Lengkap, Kelas, No. Absen"
    st.session_state["sf_mapel"] = parsed.get("mapel", "")
    jenjang = str(parsed.get("kelas") or "")
    st.session_state["sf_jenjang"] = jenjang if jenjang in KELAS_PER_JENJANG else list(KELAS_PER_JENJANG)[0]
    st.session_state["sf_kelas_pilihan"] = ""
    st.session_state["sf_hari"] = ""
    st.session_state["sf_waktu"] = ""


def collect_edits(parsed):
    edits = {}
    for sec in parsed["bagian"]:
        for s in sec["soal"]:
            no = s["no"]
            e = {"teks": st.session_state.get(k(no, "teks"), s["teks"]),
                 "poin": st.session_state.get(k(no, "poin"), 0)}
            letters = [o["huruf"] for o in s["opsi"]]
            if s["tipe"] in ("pilihan_ganda", "pg_kompleks"):
                e["opsi"] = [st.session_state.get(k(no, f"o{i}"), o["teks"]) for i, o in enumerate(s["opsi"])]
            if s["tipe"] in ("pilihan_ganda", "benar_salah"):
                v = st.session_state.get(k(no, "kunci"), KOSONG)
                e["kunci"] = [letters.index(v)] if v in letters else []
            elif s["tipe"] == "pg_kompleks":
                e["kunci"] = [letters.index(v) for v in st.session_state.get(k(no, "kunci_m"), []) if v in letters]
            elif s["tipe"] == "menjodohkan":
                v = st.session_state.get(k(no, "kunci"), KOSONG)
                hs = [p["huruf"] for p in s["pilihan"]]
                e["kunci"] = [hs.index(v)] if v in hs else []
            edits[str(no)] = e
    return edits


def ringkas(t, n=60):
    t = re.sub(r"\s+", " ", t or "").strip()
    return t if len(t) <= n else t[: n - 1] + "…"


# ==============================================================================
st.markdown("## 📝 Soal → Google Form")
st.caption("Ubah naskah soal DOCX menjadi Google Form kuis, lengkap dengan kunci dan skor.")

with st.expander("ℹ️ Cara pakai"):
    st.markdown("""
1. Unggah **naskah soal (.docx)** dengan judul bagian seperti *SOAL PILIHAN GANDA*, *SOAL PILIHAN GANDA KOMPLEK*,
   *SOAL BENAR SALAH*, *SOAL ISIAN SINGKAT*, *SOAL MENJODOHKAN*, *URAIAN*.
2. Periksa hasil pembacaan. Perbaiki teks/opsi bila perlu, lalu isi **kunci** dan **skor**.
3. Tekan **Buat Paket Soal**, unduh file **.json**.
4. Buka aplikasi web GAS **"FormPrint – Buat Form"**, unggah file tersebut, klik **Buat Form**.

*Opsional di dokumen:* tulis `Kunci: B` dan `Skor: 2` di bawah soal agar terisi otomatis.
*Catatan Google Form:* gambar pada **opsi** jawaban dan kunci **isian singkat** tidak bisa dibuat otomatis.
""")

up = st.file_uploader("Naskah soal (.docx)", type=["docx"], key="sf_upload")
if up is None:
    st.stop()

sig = f"{up.name}-{up.size}"
if st.session_state.get("sf_sig") != sig:
    try:
        parsed = parse_docx(up.getvalue())
    except Exception as e:
        st.error(f"Dokumen tidak bisa dibaca: {e}")
        st.stop()
    st.session_state["sf_parsed"] = parsed
    st.session_state["sf_sig"] = sig
    st.session_state.pop("sf_paket", None)
    seed(parsed)

parsed = st.session_state["sf_parsed"]
semua = [s for sec in parsed["bagian"] for s in sec["soal"]]

if not parsed["bagian"]:
    st.error("Tidak ada judul bagian soal yang dikenali (mis. 'SOAL PILIHAN GANDA'). Periksa format dokumen.")
    st.stop()

# ---- Ringkasan ----
st.success(f"Terbaca **{len(semua)} soal** dalam **{len(parsed['bagian'])} bagian**.")
cols = st.columns(min(len(parsed["bagian"]), 6))
for c, sec in zip(cols, parsed["bagian"]):
    c.metric(TIPE_LABEL.get(sec["tipe"], sec["tipe"]), len(sec["soal"]))

masalah = [(s["no"], w) for s in semua for w in s["peringatan"]]
masalah += [(None, f"{sec['judul']}: {sec['peringatan']}") for sec in parsed["bagian"] if sec.get("peringatan")]
if masalah:
    st.warning("Perlu diperiksa:\n\n" + "\n".join(f"- {'No. ' + str(n) + ': ' if n else ''}{w}" for n, w in masalah))

# ---- Form pemeriksaan ----
with st.form("sf_form"):
    st.markdown("#### ⚙️ Pengaturan Form")
    st.text_input("Judul Form", key="sf_judul")
    st.checkbox("Gunakan Form template sekolah (tema, header, email & identitas ikut template)",
                value=True, key="sf_template")
    t1, t2 = st.columns(2)
    with t1:
        st.text_input("Mata Pelajaran", key="sf_mapel")
        st.text_input("Hari/Tanggal", key="sf_hari", placeholder="mis. Senin, 1 Desember 2026")
    with t2:
        st.selectbox("Jenjang kelas", list(KELAS_PER_JENJANG), key="sf_jenjang")
        st.text_input("Waktu (menit)", key="sf_waktu", placeholder="mis. 90")
    st.text_input("Pilihan Kelas (kosongkan = otomatis dari jenjang)", key="sf_kelas_pilihan",
                  placeholder="otomatis, mis. 9A, 9B, 9C, …",
                  help="Daftar rombel per jenjang diatur di bagian atas file halaman ini (KELAS_PER_JENJANG).")
    with st.expander("Jika TIDAK memakai template"):
        st.text_area("Deskripsi Form (Petunjuk Umum)", key="sf_desk", height=110)
        st.text_input("Isian identitas di Bagian 1 (pisahkan dengan koma)", key="sf_identitas")
    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("Semua soal wajib dijawab", value=True, key="sf_wajib")
        st.checkbox("Tampilkan nomor soal", value=True, key="sf_nomor")
    with c2:
        st.checkbox("Acak urutan soal", value=False, key="sf_acak")
        st.checkbox("Hanya pengguna organisasi", value=False, key="sf_org",
                    help="Responden harus login dengan akun sekolah.")
    st.number_input("Skor default untuk soal yang skornya masih 0", min_value=0, step=1, value=1, key="sf_poin_default")

    st.markdown("#### 🔍 Periksa Soal, Kunci & Skor")
    for sec in parsed["bagian"]:
        n_kunci = sum(1 for s in sec["soal"] if default_key_index(s))
        with st.expander(f"{sec['judul']} — {len(sec['soal'])} soal", expanded=False):
            st.text_input("Petunjuk bagian", key=f"sf_sec_{sec['judul']}_pet")
            if sec["tipe"] == "menjodohkan" and sec["soal"]:
                st.caption("Pilihan jawaban: " + " · ".join(
                    f"{p['huruf']}. {ringkas(p['teks'], 40)}" for p in sec["soal"][0]["pilihan"]))
            for s in sec["soal"]:
                no = s["no"]
                st.markdown(f"**No. {no}**" + (" ⚠️" if s["peringatan"] else ""))
                for g in s["gambar"]:
                    try:
                        st.image(g, width=240)
                    except Exception:
                        st.caption("(gambar tidak bisa ditampilkan)")
                st.text_area("Teks soal", key=k(no, "teks"), height=90 if len(s["teks"]) > 150 else 68,
                             label_visibility="collapsed")
                letters = [o["huruf"] for o in s["opsi"]]
                if s["tipe"] in ("pilihan_ganda", "pg_kompleks"):
                    for i, o in enumerate(s["opsi"]):
                        st.text_input(f"Opsi {o['huruf']}", key=k(no, f"o{i}"))
                ck, cp = st.columns([3, 1])
                with ck:
                    if s["tipe"] in ("pilihan_ganda", "benar_salah"):
                        st.selectbox("Kunci", [KOSONG] + letters, key=k(no, "kunci"),
                                     format_func=lambda h, m={o["huruf"]: o["teks"] for o in s["opsi"]}:
                                     h if h == KOSONG else f"{h}. {ringkas(m.get(h, ''), 40)}")
                    elif s["tipe"] == "pg_kompleks":
                        st.multiselect("Kunci (boleh lebih dari satu)", letters, key=k(no, "kunci_m"))
                    elif s["tipe"] == "menjodohkan":
                        st.selectbox("Kunci", [KOSONG] + [p["huruf"] for p in s["pilihan"]], key=k(no, "kunci"),
                                     format_func=lambda h, m={p["huruf"]: p["teks"] for p in s["pilihan"]}:
                                     h if h == KOSONG else f"{h}. {ringkas(m.get(h, ''), 40)}")
                    elif s["tipe"] == "isian":
                        st.caption("Kunci isian singkat diisi manual di Google Form (tidak didukung Apps Script).")
                    else:
                        st.caption("Uraian dinilai manual.")
                with cp:
                    st.number_input("Skor", min_value=0, step=1, key=k(no, "poin"))
                st.divider()

    buat = st.form_submit_button("📦 Buat Paket Soal", type="primary", use_container_width=True)

if buat:
    for sec in parsed["bagian"]:
        sec["petunjuk"] = st.session_state.get(f"sf_sec_{sec['judul']}_pet", sec.get("petunjuk", ""))
    settings = {
        "judul": st.session_state["sf_judul"].strip() or "Soal",
        "deskripsi": st.session_state["sf_desk"],
        "identitas": [x.strip() for x in st.session_state["sf_identitas"].split(",") if x.strip()],
        "wajib": st.session_state["sf_wajib"],
        "acak_soal": st.session_state["sf_acak"],
        "tampilkan_nomor": st.session_state["sf_nomor"],
        "batasi_organisasi": st.session_state["sf_org"],
        "poin_default": st.session_state["sf_poin_default"],
        "template": {
            "pakai": st.session_state["sf_template"],
            "mapel": st.session_state["sf_mapel"].strip(),
            "hari_tanggal": st.session_state["sf_hari"].strip(),
            "waktu": st.session_state["sf_waktu"].strip(),
            "jenjang": st.session_state["sf_jenjang"],
        },
        "kelas_pilihan": [x.strip() for x in st.session_state["sf_kelas_pilihan"].split(",") if x.strip()]
                         or daftar_kelas(st.session_state["sf_jenjang"]),
    }
    with st.spinner("Menyusun paket soal..."):
        paket, catatan = build_paket(parsed, collect_edits(parsed), settings)
    st.session_state["sf_paket"] = (paket, catatan)

if st.session_state.get("sf_paket"):
    paket, catatan = st.session_state["sf_paket"]
    tanpa_kunci = [s["no"] for b in paket["bagian"] for s in b["soal"]
                   if s["tipe"] in ("pilihan_ganda", "pg_kompleks", "benar_salah", "menjodohkan") and not s["kunci"]]
    total = sum(s["poin"] for b in paket["bagian"] for s in b["soal"])
    st.success(f"Paket siap: {sum(len(b['soal']) for b in paket['bagian'])} soal, total skor {total}.")
    if tanpa_kunci:
        st.warning("Belum ada kunci: No. " + ", ".join(map(str, tanpa_kunci)))
    for c in catatan:
        st.info(c)
    nama = re.sub(r"[^A-Za-z0-9_-]+", "_", paket["judul"]).strip("_")[:60] or "paket_soal"
    st.download_button("⬇️ Unduh Paket Soal (.json)", data=json.dumps(paket, ensure_ascii=False).encode("utf-8"),
                       file_name=f"{nama}.json", mime="application/json", use_container_width=True)
    st.caption("Unggah file ini ke aplikasi web GAS **FormPrint – Buat Form**. "
               "Paket ini berisi kunci jawaban: jangan dibagikan ke siswa.")
