"""
ui_theme.py — Tampilan bersama FormPrint Studio (dipakai semua halaman).

- Warna lewat variabel CSS yang otomatis berganti mengikuti mode terang/gelap perangkat.
- Navigasi antarhalaman.
- Kartu "Alur Penggunaan" berisi link ke aplikasi web GAS.

Link aplikasi GAS diatur di Streamlit Secrets (Settings → Secrets):
    GAS_SALINAN_URL   = "https://script.google.com/macros/s/.../exec"   # Salinan Form Satu Halaman
    GAS_BUAT_FORM_URL = "https://script.google.com/macros/s/.../exec"   # Buat Google Form dari Soal
"""
import html
import os

import streamlit as st

THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {
  --fp-bg: #f8fafc;          --fp-card: #ffffff;        --fp-border: #e2e8f0;
  --fp-input-bg: #f8fafc;    --fp-input-border: #cbd5e1;
  --fp-text: #0f172a;        --fp-label: #1e293b;       --fp-muted: #64748b;   --fp-faint: #94a3b8;
  --fp-accent: #4f46e5;      --fp-accent-strong: #4338ca;
  --fp-accent-soft: #e0e7ff; --fp-accent-soft-text: #4338ca; --fp-hover-bg: #eef2ff;
  --fp-tabs-bg: #e2e8f0;     --fp-tab-active: #ffffff;
  --fp-ok-bg: #ecfdf5;       --fp-ok-border: #a7f3d0;   --fp-ok-text: #065f46;
  --fp-shadow: rgba(15, 23, 42, 0.05);
}
@media (prefers-color-scheme: dark) {
  :root {
    --fp-bg: #0e1117;          --fp-card: #161b26;        --fp-border: #2a3242;
    --fp-input-bg: #1b2130;    --fp-input-border: #3a4456;
    --fp-text: #f1f5f9;        --fp-label: #e2e8f0;       --fp-muted: #94a3b8;   --fp-faint: #64748b;
    --fp-accent: #818cf8;      --fp-accent-strong: #6366f1;
    --fp-accent-soft: #312e81; --fp-accent-soft-text: #c7d2fe; --fp-hover-bg: #1e1b4b;
    --fp-tabs-bg: #1f2633;     --fp-tab-active: #2a3242;
    --fp-ok-bg: #052e22;       --fp-ok-border: #065f46;   --fp-ok-text: #a7f3d0;
    --fp-shadow: rgba(0, 0, 0, 0.35);
  }
}

html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif !important; }
.stApp { background-color: var(--fp-bg); }

#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.stDeployButton {display: none;}

.block-container { max-width: 820px !important; padding-top: 2.5rem !important; padding-bottom: 4rem !important; }

[data-testid="stForm"] {
  background-color: var(--fp-card);
  border: 1px solid var(--fp-border);
  border-radius: 16px;
  padding: 2rem 2.2rem !important;
  box-shadow: 0 4px 20px -2px var(--fp-shadow);
}

.stTextInput input {
  border-radius: 10px !important;
  border: 1px solid var(--fp-input-border) !important;
  padding: 0.65rem 0.9rem !important;
  font-size: 0.92rem !important;
  background-color: var(--fp-input-bg) !important;
}
.stTextInput input:focus {
  border-color: var(--fp-accent) !important;
  box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2) !important;
  background-color: var(--fp-card) !important;
}

[data-testid="stFileUploader"] section {
  border-radius: 12px !important;
  border: 1.5px dashed var(--fp-input-border) !important;
  background-color: var(--fp-input-bg) !important;
  padding: 1rem !important;
  transition: all 0.2s ease;
}
[data-testid="stFileUploader"] section:hover {
  border-color: var(--fp-accent) !important;
  background-color: var(--fp-hover-bg) !important;
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

.stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: var(--fp-tabs-bg); padding: 6px; border-radius: 12px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 16px; font-weight: 600; font-size: 0.9rem; }
.stTabs [aria-selected="true"] {
  background-color: var(--fp-tab-active) !important;
  color: var(--fp-accent) !important;
  box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

.stTextInput label, .stSelectbox label, [data-testid="stFileUploader"] label {
  font-weight: 600 !important;
  color: var(--fp-label) !important;
  font-size: 0.9rem !important;
}

/* ---- Kartu alur penggunaan ---- */
.fp-steps { display: grid; gap: 10px; margin: 0.4rem 0 1.6rem 0; }
.fp-step {
  display: flex; gap: 14px; align-items: flex-start;
  background: var(--fp-card); border: 1px solid var(--fp-border); border-radius: 14px;
  padding: 12px 16px; box-shadow: 0 2px 10px -4px var(--fp-shadow);
}
.fp-step-no {
  flex: 0 0 30px; height: 30px; border-radius: 9999px;
  background: var(--fp-accent-soft); color: var(--fp-accent-soft-text);
  font-weight: 800; font-size: 0.9rem; display: flex; align-items: center; justify-content: center;
}
.fp-step-title { color: var(--fp-text); font-weight: 700; font-size: 0.93rem; margin: 3px 0 2px 0; }
.fp-step-desc { color: var(--fp-muted); font-size: 0.84rem; line-height: 1.5; }
.fp-step-desc b { color: var(--fp-label); }
.fp-link {
  display: inline-block; margin-top: 6px; padding: 5px 12px; border-radius: 9999px;
  background: var(--fp-accent-soft); color: var(--fp-accent-soft-text) !important;
  font-weight: 700; font-size: 0.8rem; text-decoration: none !important;
}
.fp-link:hover { filter: brightness(1.08); }
.fp-link-missing { display: inline-block; margin-top: 6px; color: var(--fp-faint); font-size: 0.78rem; font-style: italic; }
.fp-flow-title { color: var(--fp-text); font-weight: 800; font-size: 1.02rem; margin: 0.2rem 0 0.5rem 0; }
</style>
"""


def apply_theme() -> None:
    st.markdown(THEME_CSS, unsafe_allow_html=True)


def nav_links() -> None:
    try:
        c1, c2 = st.columns(2)
        with c1:
            st.page_link("app.py", label="Google Form → Naskah Cetak", icon="🖨️")
        with c2:
            st.page_link("pages/2_Soal_ke_Google_Form.py", label="Soal DOCX → Google Form", icon="📝")
    except Exception:
        st.caption("Halaman *Soal → Google Form*: tambahkan `/Soal_ke_Google_Form` di akhir alamat aplikasi.")


def gas_url(key: str) -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    return os.environ.get(key, "").strip()


def link_button(key: str, label: str) -> str:
    url = gas_url(key)
    if url:
        return f'<a class="fp-link" href="{html.escape(url)}" target="_blank" rel="noopener">🔗 {html.escape(label)}</a>'
    return f'<span class="fp-link-missing">Link aplikasi belum diatur — isi {key} di Streamlit Secrets.</span>'


def render_steps(title: str, steps: list) -> None:
    """steps: [(judul, deskripsi_html, html_tambahan_opsional)]"""
    items = []
    for i, step in enumerate(steps, 1):
        judul, desk = step[0], step[1]
        extra = step[2] if len(step) > 2 else ""
        items.append(
            f'<div class="fp-step"><div class="fp-step-no">{i}</div><div>'
            f'<div class="fp-step-title">{judul}</div><div class="fp-step-desc">{desk}</div>{extra}'
            f'</div></div>'
        )
    st.markdown(f'<div class="fp-flow-title">{title}</div><div class="fp-steps">{"".join(items)}</div>',
                unsafe_allow_html=True)
