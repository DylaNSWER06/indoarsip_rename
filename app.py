"""
INDOARSIP RENAME - Sistem Otomatis Penamaan Arsip Digital
Batch File Renaming System based on Excel Reference
"""

import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import time
import base64
from io import BytesIO

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="INDOARSIP RENAME - Sistem Rename Arsip",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# LOGO HELPER
# ============================================================================

def get_logo_base64(logo_path="logo.jpg"):
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

LOGO_B64 = get_logo_base64()

def logo_html(size="2.5rem", fallback_emoji="📦"):
    if LOGO_B64:
        return f'<img src="data:image/png;base64,{LOGO_B64}" style="height:{size};width:{size};object-fit:contain;border-radius:6px;">'
    return f'<span style="font-size:{size};">{fallback_emoji}</span>'

# ============================================================================
# CUSTOM CSS
# ============================================================================

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    * { font-family: 'Inter', sans-serif; }

    .block-container { padding-top: 0.75rem !important; padding-bottom: 1rem !important; }
    .stMarkdown p { margin-bottom: 0.3rem !important; }
    hr { margin: 0.5rem 0 !important; }

    .main-header {
        background: linear-gradient(135deg, #0f2554 0%, #1e3c72 60%, #2a5298 100%);
        padding: 0.9rem 1.5rem; border-radius: 8px; margin-bottom: 0.75rem;
        box-shadow: 0 3px 10px rgba(30,60,114,0.2);
        display: flex; align-items: center; gap: 1rem;
    }
    .main-header h1 {
        color: white; font-size: 1.3rem; font-weight: 700;
        margin: 0 0 0.2rem 0; letter-spacing: 2px; line-height: 1.3;
        padding-top: 1.8rem;
    }
    .main-header p  { color: #c7d7f7; font-size: 0.78rem; margin: 0; font-weight: 400; }
    .header-badge {
        background: rgba(255,255,255,0.18); color: rgba(255,255,255,0.9); font-size: 0.65rem;
        padding: 0.18rem 0.6rem; border-radius: 20px; margin-top: 0.35rem;
        display: inline-block; letter-spacing: 0.8px; border: 1px solid rgba(255,255,255,0.2);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f2554 0%, #1e3c72 100%) !important;
    }
    [data-testid="stSidebar"] * { color: white !important; }
    .sidebar-logo {
        text-align: center; padding: 1rem 0.75rem 0.75rem;
        border-bottom: 1px solid rgba(255,255,255,0.15); margin-bottom: 1rem;
    }
    .sidebar-logo-title { font-size: 0.95rem; font-weight: 700; letter-spacing: 1.5px; margin: 0.4rem 0 0.2rem; color: white; }
    .sidebar-logo-sub   { font-size: 0.68rem; color: #c7d7f7 !important; letter-spacing: 1px; }

    .nav-section-label {
        font-size: 0.65rem; font-weight: 600; letter-spacing: 2px;
        color: #93b4e0 !important; text-transform: uppercase; padding: 0 0.5rem; margin-bottom: 0.4rem;
    }
    .nav-item {
        display: flex; align-items: center; gap: 0.6rem;
        padding: 0.55rem 0.75rem; border-radius: 8px; margin: 0.2rem 0;
    }
    .nav-item.active   { background: rgba(255,255,255,0.18) !important; border-left: 3px solid #60a5fa; }
    .nav-item.inactive { background: rgba(255,255,255,0.05); border-left: 3px solid transparent; }
    .nav-item-icon { font-size: 1rem; }
    .nav-item-text { font-size: 0.82rem; font-weight: 600; }
    .nav-item-sub  { font-size: 0.68rem; color: #93b4e0 !important; }
    .nav-status {
        margin-left: auto; font-size: 0.65rem; padding: 0.15rem 0.45rem;
        border-radius: 10px; font-weight: 600;
    }
    .nav-status.done         { background: #16a34a; color: white !important; }
    .nav-status.pending      { background: rgba(255,255,255,0.15); color: #c7d7f7 !important; }
    .nav-status.active-badge { background: #2a5298; color: white !important; }

    .sidebar-stats {
        background: rgba(255,255,255,0.08); border-radius: 10px; padding: 1rem; margin: 1rem 0;
    }
    .sidebar-stat-row {
        display: flex; justify-content: space-between; align-items: center;
        padding: 0.3rem 0; border-bottom: 1px solid rgba(255,255,255,0.08); font-size: 0.82rem;
    }
    .sidebar-stat-row:last-child { border-bottom: none; }
    .sidebar-stat-label { color: #93b4e0 !important; }
    .sidebar-stat-value { font-weight: 700; color: white !important; }
    .sidebar-stat-value.green { color: #4ade80 !important; }
    .sidebar-stat-value.red   { color: #f87171 !important; }
    .sidebar-divider { border: none; border-top: 1px solid rgba(255,255,255,0.12); margin: 1rem 0; }
    .sidebar-footer  { font-size: 0.72rem; color: #7096c4 !important; text-align: center; padding: 0.5rem; }

    .stTabs [data-baseweb="tab-list"] {
        gap: 6px; background-color: #f0f4fb; padding: 0.3rem;
        border-radius: 8px; border: 1px solid #dce6f7;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: white; border-radius: 6px; padding: 0.45rem 1rem;
        border: 2px solid transparent; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .stTabs [data-baseweb="tab"] p { font-weight: 700 !important; color: #1e3c72 !important; font-size: 0.85rem !important; }
    .stTabs [aria-selected="true"] {
        background-color: #1e3c72 !important; border-color: #0f2554 !important;
        box-shadow: 0 3px 8px rgba(30,60,114,0.3) !important;
    }
    .stTabs [aria-selected="true"] p { color: white !important; font-weight: 700 !important; }

    .section-card {
        background: white; border-radius: 8px; padding: 0.5rem 0.75rem;
        border: 1px solid #e8eef7; box-shadow: 0 1px 4px rgba(0,0,0,0.04); margin-bottom: 0.4rem;
    }
    .section-title { font-size: 0.85rem; font-weight: 700; color: #1e3c72; margin: 0; }

    [data-testid="stFileUploader"] {
        border: 2px solid #c7d7f7 !important;
        border-radius: 8px !important;
        padding: 0.3rem !important;
        background: #f8faff !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #2a5298 !important;
        box-shadow: 0 0 0 3px rgba(42,82,152,0.1) !important;
    }
    .stTextInput input {
        border: 2px solid #c7d7f7 !important;
        border-radius: 8px !important;
        padding: 0.45rem 0.8rem !important;
        font-size: 0.85rem !important;
        background: #f8faff !important;
        transition: border-color 0.2s, box-shadow 0.2s !important;
    }
    .stTextInput input:focus {
        border-color: #2a5298 !important;
        box-shadow: 0 0 0 3px rgba(42,82,152,0.12) !important;
        background: white !important;
    }
    .stRadio > div {
        border: 2px solid #c7d7f7;
        border-radius: 8px;
        padding: 0.4rem 0.75rem;
        background: #f8faff;
    }

    .info-banner {
        background: linear-gradient(135deg, #eff6ff, #dbeafe);
        border: 1px solid #93c5fd; border-left: 3px solid #2a5298;
        border-radius: 6px; padding: 0.5rem 0.9rem; margin-bottom: 0.5rem;
        font-size: 0.78rem; color: #1e40af;
    }

    .alert-box {
        border-radius: 8px; padding: 0.75rem 1rem; margin: 0.4rem 0;
        display: flex; align-items: flex-start; gap: 0.75rem;
        animation: slideIn 0.3s ease;
    }
    .alert-box.error   { background: #fef2f2; border: 1px solid #fca5a5; border-left: 4px solid #ef4444; }
    .alert-box.warning { background: #fffbeb; border: 1px solid #fcd34d; border-left: 4px solid #f59e0b; }
    .alert-icon { font-size: 1rem; margin-top: 0.1rem; flex-shrink: 0; }
    .alert-text { font-size: 0.85rem; font-weight: 600; }
    .alert-text.error   { color: #b91c1c; }
    .alert-text.warning { color: #92400e; }

    .metric-row { display: flex; gap: 0.75rem; margin: 0.75rem 0; }
    .metric-card {
        flex: 1; background: white; border-radius: 10px; padding: 0.9rem 1rem;
        border: 1px solid #e8eef7; box-shadow: 0 2px 8px rgba(0,0,0,0.05); text-align: center;
    }
    .metric-card.total   { border-top: 3px solid #2a5298; }
    .metric-card.match   { border-top: 3px solid #16a34a; }
    .metric-card.nomatch { border-top: 3px solid #dc2626; }
    .metric-label { font-size: 0.72rem; font-weight: 600; color: #6b7280; text-transform: uppercase; letter-spacing: 1px; }
    .metric-value { font-size: 1.7rem; font-weight: 700; color: #1e3c72; margin: 0.15rem 0; }
    .metric-card.match   .metric-value { color: #16a34a; }
    .metric-card.nomatch .metric-value { color: #dc2626; }
    .metric-sub { font-size: 0.72rem; color: #9ca3af; }

    .success-notification {
        background: linear-gradient(135deg, #f0fdf4, #dcfce7);
        border: 1px solid #86efac; border-left: 5px solid #16a34a;
        border-radius: 10px; padding: 1rem 1.5rem; margin: 0.75rem 0;
        display: flex; align-items: center; gap: 1rem;
        box-shadow: 0 4px 12px rgba(22,163,74,0.1);
        animation: slideIn 0.4s ease;
    }
    @keyframes slideIn {
        from { opacity: 0; transform: translateY(-10px); }
        to   { opacity: 1; transform: translateY(0); }
    }
    .success-icon {
        width: 40px; height: 40px; background: #16a34a; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.2rem; flex-shrink: 0; box-shadow: 0 4px 12px rgba(22,163,74,0.3);
    }
    .success-title    { font-size: 0.95rem; font-weight: 700; color: #15803d; margin: 0; }
    .success-subtitle { font-size: 0.82rem; color: #166534; margin: 0.2rem 0 0 0; }

    .download-card {
        background: white; border-radius: 10px; padding: 1.25rem;
        border: 1px solid #e8eef7; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .download-card h4 { color: #1e3c72; font-size: 0.95rem; font-weight: 700; margin-bottom: 0.4rem; }

    .file-item {
        background: #f8faff; border-radius: 8px; padding: 0.6rem 0.9rem;
        border: 1px solid #e8eef7; margin-bottom: 0.4rem;
    }
    .file-name     { font-weight: 600; color: #1e3c72; font-size: 0.88rem; }
    .file-original { font-size: 0.75rem; color: #9ca3af; }

    .stButton button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white; font-weight: 600; padding: 0.5rem 1.5rem;
        border-radius: 7px; border: none;
        box-shadow: 0 3px 8px rgba(30,60,114,0.2);
        transition: all 0.2s ease; letter-spacing: 0.3px;
        font-size: 0.85rem !important;
    }
    .stButton button:hover { box-shadow: 0 5px 12px rgba(30,60,114,0.3); transform: translateY(-1px); }

    .empty-state { text-align: center; padding: 2.5rem 2rem; color: #9ca3af; }
    .empty-state-icon  { font-size: 2.5rem; margin-bottom: 0.75rem; }
    .empty-state-title { font-size: 1rem; font-weight: 600; color: #6b7280; }
    .empty-state-sub   { font-size: 0.85rem; margin-top: 0.4rem; }

    .footer {
        text-align: center; color: #9ca3af;
        padding: 1.5rem 0 0.75rem; border-top: 1px solid #f3f4f6; margin-top: 1.5rem;
    }
    .footer strong { color: #1e3c72; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# SESSION STATE
# ============================================================================

defaults = {
    'validated': False,
    'temp_dir': None,
    'file_list': [],
    'reference_data': None,
    'matched_files': [],
    'unmatched_files': [],
    'rename_mapping': {},
    'file_contents_cache': {},
    'show_download_section': False,
    'rename_done': False,
    'rename_count': 0,
    'active_tab': 0,
    'uploaded_size_mb': 0,
    'last_zip_name': None,
    'last_files_count': 0,
    'last_excel_name': None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================================
# RESET VALIDASI JIKA FILE DIHAPUS
# ============================================================================

def check_and_reset_if_files_removed():
    try:
        zip_uploader   = st.session_state.get("zip_uploader")
        files_uploader = st.session_state.get("files_uploader")
        excel_uploader = st.session_state.get("excel_uploader")
        upload_type    = st.session_state.get("upload_type", "File ZIP Arsip")

        arsip_removed = False
        excel_removed = False

        if upload_type == "File ZIP Arsip":
            try:
                cur_zip = zip_uploader.name if zip_uploader else None
            except Exception:
                cur_zip = None
            if st.session_state.last_zip_name and cur_zip is None:
                arsip_removed = True
            st.session_state.last_zip_name = cur_zip
        else:
            try:
                cur_count = len(files_uploader) if files_uploader else 0
            except Exception:
                cur_count = 0
            if st.session_state.last_files_count > 0 and cur_count == 0:
                arsip_removed = True
            st.session_state.last_files_count = cur_count

        try:
            cur_excel = excel_uploader.name if excel_uploader else None
        except Exception:
            cur_excel = None
        if st.session_state.last_excel_name and cur_excel is None:
            excel_removed = True
        st.session_state.last_excel_name = cur_excel

        if (arsip_removed or excel_removed) and st.session_state.validated:
            st.session_state.update({
                'validated': False,
                'active_tab': 0,
                'matched_files': [],
                'unmatched_files': [],
                'rename_mapping': {},
                'file_contents_cache': {},
                'show_download_section': False,
                'rename_done': False,
            })

    except Exception:
        pass

# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-logo">
        {logo_html(size="48px")}
        <div class="sidebar-logo-title">INDOARSIP RENAME</div>
        <div class="sidebar-logo-sub">ARCHIVE MANAGEMENT</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="nav-section-label">Navigasi</div>', unsafe_allow_html=True)

    tab1_css    = "active" if st.session_state.active_tab == 0 else "inactive"
    tab2_css    = "active" if st.session_state.active_tab == 1 else "inactive"
    tab1_status = "done" if st.session_state.validated else "active-badge"
    tab1_label  = "✓ Selesai" if st.session_state.validated else "Aktif"
    tab2_status = "done" if st.session_state.rename_done else ("active-badge" if st.session_state.validated else "pending")
    tab2_label  = "✓ Selesai" if st.session_state.rename_done else ("Siap" if st.session_state.validated else "Terkunci 🔒")

    st.markdown(f"""
    <div class="nav-item {tab1_css}">
        <span class="nav-item-icon">📋</span>
        <div>
            <div class="nav-item-text">Upload & Validasi</div>
            <div class="nav-item-sub">Upload arsip & referensi</div>
        </div>
        <span class="nav-status {tab1_status}">{tab1_label}</span>
    </div>
    <div class="nav-item {tab2_css}" style="margin-top:0.25rem;">
        <span class="nav-item-icon">✅</span>
        <div>
            <div class="nav-item-text">Preview & Rename</div>
            <div class="nav-item-sub">Proses & download hasil</div>
        </div>
        <span class="nav-status {tab2_status}">{tab2_label}</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.session_state.validated:
        total   = len(st.session_state.file_list)
        matched = len(st.session_state.matched_files)
        nomatch = len(st.session_state.unmatched_files)
        pct     = f"{matched/total*100:.0f}%" if total else "0%"
        rsc     = "green" if st.session_state.rename_done else ""
        rst     = "✓ Selesai" if st.session_state.rename_done else "Belum"

        WARN_MB   = 150
        DANGER_MB = 300
        total_mb  = st.session_state.get('uploaded_size_mb', 0)
        size_pct  = min(total_mb / DANGER_MB * 100, 100)
        bar_color = '#4ade80' if total_mb < WARN_MB else ('#facc15' if total_mb < DANGER_MB else '#f87171')

        if total_mb < WARN_MB:
            size_color  = "green"
            size_status = "✅ Aman"
            size_desc   = "Ukuran file dalam batas aman. Proses rename dapat berjalan normal di Streamlit Cloud."
        elif total_mb < DANGER_MB:
            size_color  = ""
            size_status = "⚡ Hati-hati"
            size_desc   = "Ukuran file mendekati batas aman. Aplikasi mungkin berjalan lebih lambat dari biasanya."
        else:
            size_color  = "red"
            size_status = "⚠️ Berisiko"
            size_desc   = "File terlalu besar untuk diproses di Streamlit Cloud. Aplikasi berisiko crash. Disarankan jalankan secara lokal."

        st.markdown(f"""
        <hr class="sidebar-divider">
        <div class="nav-section-label">Statistik Arsip</div>
        <div class="sidebar-stats">
            <div class="sidebar-stat-row">
                <span class="sidebar-stat-label">Total File</span>
                <span class="sidebar-stat-value">{total}</span>
            </div>
            <div class="sidebar-stat-row">
                <span class="sidebar-stat-label">Cocok</span>
                <span class="sidebar-stat-value green">{matched} ({pct})</span>
            </div>
            <div class="sidebar-stat-row">
                <span class="sidebar-stat-label">Tidak Cocok</span>
                <span class="sidebar-stat-value red">{nomatch}</span>
            </div>
            <div class="sidebar-stat-row">
                <span class="sidebar-stat-label">Status Rename</span>
                <span class="sidebar-stat-value {rsc}">{rst}</span>
            </div>
            <div class="sidebar-stat-row">
                <span class="sidebar-stat-label">Ukuran Upload</span>
                <span class="sidebar-stat-value {size_color}">{total_mb:.1f} MB</span>
            </div>
        </div>

        <div style="margin-top:0.5rem;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:0.3rem;">
                <span style="font-size:0.65rem;color:#93b4e0;">Estimasi Beban Memori</span>
                <span style="font-size:0.65rem;font-weight:700;color:{bar_color};">{size_status}</span>
            </div>
            <div style="background:rgba(255,255,255,0.15);border-radius:6px;height:7px;overflow:hidden;">
                <div style="width:{size_pct:.1f}%;height:100%;background:{bar_color};border-radius:6px;transition:width 0.4s ease;"></div>
            </div>
            <div style="font-size:0.62rem;color:#93b4e0;margin-top:0.25rem;text-align:right;">{total_mb:.1f} MB dari ~300 MB batas aman</div>
        </div>

        <div style="
            background:rgba(255,255,255,0.07);
            border:1px solid rgba(255,255,255,0.15);
            border-left:3px solid {bar_color};
            border-radius:8px;
            padding:0.65rem 0.75rem;
            margin-top:0.6rem;
        ">
            <div style="font-size:0.68rem;font-weight:700;color:white;margin-bottom:0.35rem;">
                ℹ️ Kebijakan Ukuran File
            </div>
            <div style="font-size:0.63rem;color:#c7d7f7;line-height:1.5;">
                {size_desc}
            </div>
            <div style="margin-top:0.4rem;border-top:1px solid rgba(255,255,255,0.1);padding-top:0.4rem;">
                <div style="font-size:0.63rem;color:#93b4e0;line-height:1.6;">
                    🟢 &lt; 150 MB — Aman untuk Cloud<br>
                    🟡 150–300 MB — Hati-hati, bisa lambat<br>
                    🔴 &gt; 300 MB — Gunakan mode <b>Lokal</b>
                </div>
            </div>
            <div style="margin-top:0.4rem;border-top:1px solid rgba(255,255,255,0.1);padding-top:0.4rem;">
                <div style="font-size:0.62rem;color:#7096c4;line-height:1.5;">
                    💡 Tidak ada batas upload yang tetap di Streamlit Cloud, namun file besar menghabiskan RAM server (~1 GB). Jika aplikasi tiba-tiba berhenti, kurangi ukuran file atau jalankan secara lokal.
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <hr class="sidebar-divider">
    <div class="sidebar-footer">INDOARSIP RENAME v2.0<br>Professional Archive Management</div>
    """, unsafe_allow_html=True)

# ============================================================================
# HEADER
# ============================================================================

st.markdown("""
<div class="main-header">
    <div>
        <h1>INDOARSIP RENAME</h1>
        <p>Sistem Otomatis Penamaan Arsip Digital</p>
        <span class="header-badge">✦ PROFESSIONAL ARCHIVE MANAGEMENT</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def extract_zip(zip_file):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file, 'r') as z:
        z.extractall(temp_dir)
    return temp_dir

def get_files_from_directory(directory):
    files = []
    for root, dirs, filenames in os.walk(directory):
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__MACOSX']
        for filename in filenames:
            if (not filename.startswith('.') and not filename.startswith('__')
                    and filename != '.DS_Store' and not root.endswith('__MACOSX')):
                fp = os.path.join(root, filename)
                if os.path.isfile(fp):
                    files.append(fp)
    return files

def extract_code_from_filename(filename):
    import re
    name = os.path.splitext(filename)[0]
    m = re.search(r'_(\d+)$', name)
    if m: return m.group(1)
    m = re.search(r'^(\d+)', name)
    if m: return m.group(1)
    m = re.search(r'(\d+)', name)
    if m: return m.group(1)
    return name

def match_files_with_reference(file_list, reference_values):
    matched, unmatched, rename_map = [], [], {}
    for fp in file_list:
        filename = os.path.basename(fp)
        ext  = os.path.splitext(filename)[1]
        code = extract_code_from_filename(filename)
        found = False
        for ref in reference_values:
            ref_str = str(ref).strip()
            if ref_str.startswith(code):
                matched.append(fp)
                rename_map[fp] = ref_str + ext
                found = True
                break
        if not found:
            unmatched.append(fp)
    return matched, unmatched, rename_map

def create_zip_from_files(file_mapping, original_dir):
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for old_path, new_name in file_mapping.items():
            zf.write(old_path, arcname=new_name)
    buf.seek(0)
    return buf

def create_unmatched_report(unmatched_files):
    data = {
        'Nama File Tidak Cocok': [os.path.basename(f) for f in unmatched_files],
        'Path Lengkap': unmatched_files,
        'Status': ['Tidak Ditemukan di Referensi'] * len(unmatched_files)
    }
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame(data).to_excel(writer, sheet_name='Arsip Tidak Cocok', index=False)
    buf.seek(0)
    return buf

def show_metrics(total, matched, unmatched):
    pm = f"{matched/total*100:.1f}%" if total else "0%"
    pn = f"{unmatched/total*100:.1f}%" if total else "0%"
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card total">
            <div class="metric-label">Total Arsip</div>
            <div class="metric-value">{total}</div>
            <div class="metric-sub">file terdeteksi</div>
        </div>
        <div class="metric-card match">
            <div class="metric-label">Arsip Cocok</div>
            <div class="metric-value">{matched}</div>
            <div class="metric-sub">{pm} dari total</div>
        </div>
        <div class="metric-card nomatch">
            <div class="metric-label">Tidak Cocok</div>
            <div class="metric-value">{unmatched}</div>
            <div class="metric-sub">{pn} dari total</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def alert(msg, kind="error"):
    icon = "❌" if kind == "error" else "⚠️"
    st.markdown(f"""
    <div class="alert-box {kind}">
        <span class="alert-icon">{icon}</span>
        <span class="alert-text {kind}">{msg}</span>
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# AUTO SWITCH TAB via JS
# ============================================================================

st.markdown(f"""
<script>
(function() {{
    var targetTab = {st.session_state.active_tab};
    function tryClick(attempts) {{
        if (attempts <= 0) return;
        var tabs = window.parent.document.querySelectorAll('[data-baseweb="tab"]');
        if (tabs && tabs.length > targetTab) {{
            var tab = tabs[targetTab];
            if (tab.getAttribute('aria-selected') !== 'true') {{
                tab.click();
            }}
        }} else {{
            setTimeout(function() {{ tryClick(attempts - 1); }}, 150);
        }}
    }}
    setTimeout(function() {{ tryClick(8); }}, 120);
}})();
</script>
""", unsafe_allow_html=True)

# ============================================================================
# TABS
# ============================================================================

tab1, tab2 = st.tabs(["📋  Upload & Validasi Arsip", "✅  Preview & Proses Rename"])

# ── TAB 1 ──────────────────────────────────────
with tab1:
    st.markdown("##### 📂 Upload & Validasi Data Arsip")
    st.markdown("""
    <div class="info-banner">
        <strong>Cara kerja:</strong> Kode numerik dari nama file dicocokkan dengan data Excel.
        Contoh: <code>pelanggan_0001.pdf</code> → cocok dengan baris <code>0001-...</code>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown('<div class="section-card"><div class="section-title">1. File Arsip</div></div>', unsafe_allow_html=True)
        upload_type = st.radio("Pilih tipe upload:", ["File ZIP Arsip", "Folder Arsip (Multiple Files)"], key="upload_type")
        if upload_type == "File ZIP Arsip":
            zip_file = st.file_uploader("Upload file ZIP yang berisi arsip", type=['zip'], key="zip_uploader")
        else:
            uploaded_files = st.file_uploader("Upload multiple files arsip", accept_multiple_files=True, key="files_uploader")
            if uploaded_files:
                zips = [f for f in uploaded_files if f.name.lower().endswith('.zip')]
                if zips:
                    st.error(f"❌ File ZIP terdeteksi: {', '.join(f.name for f in zips)}")
                    st.warning("Gunakan opsi **'File ZIP Arsip'** untuk upload ZIP.")
                    uploaded_files = None

    with col2:
        st.markdown('<div class="section-card"><div class="section-title">2. File Referensi Excel</div></div>', unsafe_allow_html=True)
        excel_file       = st.file_uploader("Upload file Excel referensi penamaan", type=['xlsx', 'xls'], key="excel_uploader")
        reference_column = st.text_input("Nama Kolom Referensi", placeholder="Contoh: Nomor_Arsip, Kode_Dokumen ...", key="ref_column")

    check_and_reset_if_files_removed()

    st.markdown("---")

    if st.button("🔍  Validasi & Cek Arsip", use_container_width=True, type="primary"):
        errors = []

        if upload_type == "File ZIP Arsip":
            if not st.session_state.get("zip_uploader"):
                errors.append(("error", "📁 File arsip ZIP belum diupload. Silakan upload file ZIP terlebih dahulu."))
        else:
            files_up = st.session_state.get("files_uploader")
            if not files_up:
                errors.append(("error", "📁 File arsip belum diupload. Silakan upload minimal satu file arsip."))
            else:
                zips = [f.name for f in files_up if f.name.lower().endswith('.zip')]
                if zips:
                    errors.append(("error", f"File ZIP tidak bisa diupload di opsi Multiple Files: {', '.join(zips)}"))

        if not st.session_state.get("excel_uploader"):
            errors.append(("error", "📊 File Excel referensi belum diupload. Silakan upload file Excel (.xlsx/.xls)."))

        if not reference_column or not reference_column.strip():
            errors.append(("warning", "✏️ Nama kolom referensi belum diisi."))

        if errors:
            st.markdown("**Harap lengkapi data berikut sebelum melanjutkan:**")
            for kind, msg in errors:
                alert(msg, kind)
        else:
            progress = st.progress(0, text="Memulai proses validasi...")
            try:
                progress.progress(20, text="📦 Mengekstrak file arsip...")
                time.sleep(0.3)

                if upload_type == "File ZIP Arsip":
                    temp_dir    = extract_zip(zip_file)
                    file_list   = get_files_from_directory(temp_dir)
                    uploaded_mb = zip_file.size / (1024 * 1024)
                    if not file_list:
                        progress.empty()
                        st.error("❌ ZIP kosong atau tidak ada file yang valid!")
                        st.session_state.validated = False
                        st.stop()
                    st.info(f"📦 ZIP berhasil di-extract — ditemukan **{len(file_list)} file**")
                else:
                    temp_dir    = tempfile.mkdtemp()
                    file_list   = []
                    uploaded_mb = 0
                    for uf in uploaded_files:
                        fp = os.path.join(temp_dir, uf.name)
                        with open(fp, 'wb') as fh:
                            fh.write(uf.getbuffer())
                        file_list.append(fp)
                        uploaded_mb += uf.size / (1024 * 1024)
                    st.info(f"📁 **{len(file_list)} file** berhasil diupload.")

                st.session_state.temp_dir = temp_dir

                progress.progress(50, text="📊 Membaca file Excel referensi...")
                time.sleep(0.3)
                df = pd.read_excel(excel_file)

                if reference_column not in df.columns:
                    progress.empty()
                    alert(f"Kolom '{reference_column}' tidak ditemukan! Kolom tersedia: {', '.join(df.columns.tolist())}", "error")
                    st.session_state.validated = False
                    st.stop()

                reference_values = df[reference_column].dropna().astype(str).tolist()

                progress.progress(80, text="🔗 Mencocokkan file dengan referensi...")
                time.sleep(0.3)
                matched, unmatched, rename_map = match_files_with_reference(file_list, reference_values)

                st.session_state.update({
                    'file_list': file_list,
                    'reference_data': df,
                    'matched_files': matched,
                    'unmatched_files': unmatched,
                    'rename_mapping': rename_map,
                    'validated': True,
                    'show_download_section': False,
                    'rename_done': False,
                    'file_contents_cache': {},
                    'uploaded_size_mb': round(uploaded_mb, 2),
                    'active_tab': 1,
                })

                progress.progress(100, text="✅ Validasi selesai! Pindah ke Preview...")
                time.sleep(0.5)
                progress.empty()
                st.rerun()

            except Exception as ex:
                progress.empty()
                st.error(f"❌ Terjadi kesalahan: {ex}")
                st.session_state.validated = False

# ── TAB 2 ──────────────────────────────────────
with tab2:
    st.markdown("##### 🚀 Preview & Eksekusi Rename Arsip")

    if not st.session_state.validated:
        st.markdown("""
        <div class="empty-state">
            <div class="empty-state-icon">🔒</div>
            <div class="empty-state-title">Tab ini masih terkunci</div>
            <div class="empty-state-sub">Selesaikan upload & validasi di tab <strong>Upload & Validasi Arsip</strong> terlebih dahulu.</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="success-notification">
            <div class="success-icon">✔</div>
            <div>
                <p class="success-title">Validasi Berhasil — {len(st.session_state.matched_files)} arsip siap diproses</p>
                <p class="success-subtitle">Periksa detail di bawah lalu klik tombol Mulai Proses Rename.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        show_metrics(len(st.session_state.file_list), len(st.session_state.matched_files), len(st.session_state.unmatched_files))

        if st.session_state.matched_files:
            with st.expander(f"✅ {len(st.session_state.matched_files)} arsip cocok — klik untuk lihat detail"):
                mdf = pd.DataFrame({
                    'No': range(1, len(st.session_state.matched_files)+1),
                    'Nama File Asli': [os.path.basename(f) for f in st.session_state.matched_files],
                    'Kode': [extract_code_from_filename(os.path.basename(f)) for f in st.session_state.matched_files],
                    'Akan Direname Jadi': [st.session_state.rename_mapping[f] for f in st.session_state.matched_files],
                })
                st.dataframe(mdf, use_container_width=True)

        if st.session_state.unmatched_files:
            with st.expander(f"⚠️ {len(st.session_state.unmatched_files)} arsip tidak cocok"):
                st.dataframe(
                    pd.DataFrame({'Nama File': [os.path.basename(f) for f in st.session_state.unmatched_files]}),
                    use_container_width=True
                )

        st.markdown("---")

        if st.session_state.matched_files:
            st.markdown("##### 👁️ Preview Penamaan Arsip")
            prev = pd.DataFrame({
                'No': range(1, len(st.session_state.matched_files)+1),
                'Nama Arsip Lama': [os.path.basename(f) for f in st.session_state.matched_files],
                'Nama Arsip Baru': [st.session_state.rename_mapping[f] for f in st.session_state.matched_files],
                'Status': ['✅ Siap Rename'] * len(st.session_state.matched_files),
            })
            st.dataframe(prev, use_container_width=True)
            st.markdown("---")

            if st.button("🚀  Mulai Proses Rename Arsip", use_container_width=True, type="primary"):
                prog2 = st.progress(0, text="Mempersiapkan file...")
                try:
                    total = len(st.session_state.rename_mapping)
                    for i, (old_path, new_name) in enumerate(st.session_state.rename_mapping.items(), 1):
                        prog2.progress(int(i/total*90), text=f"Memproses file {i} dari {total}...")
                        if new_name not in st.session_state.file_contents_cache:
                            with open(old_path, 'rb') as fh:
                                st.session_state.file_contents_cache[new_name] = fh.read()
                        time.sleep(0.02)
                    prog2.progress(100, text="✅ Semua file berhasil diproses!")
                    time.sleep(0.5)
                    prog2.empty()
                    st.session_state.show_download_section = True
                    st.session_state.rename_done  = True
                    st.session_state.rename_count = total
                except Exception as ex:
                    prog2.empty()
                    st.error(f"❌ Terjadi kesalahan: {ex}")

            if st.session_state.get('rename_done'):
                st.markdown(f"""
                <div class="success-notification">
                    <div class="success-icon">✔</div>
                    <div>
                        <p class="success-title">Proses Rename Berhasil Diselesaikan</p>
                        <p class="success-subtitle">{st.session_state.rename_count} file arsip telah berhasil diproses dan siap untuk diunduh.</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            if st.session_state.get('show_download_section'):
                st.markdown("---")
                st.markdown("##### 📥 Download Hasil Rename")

                dl1, dl2 = st.columns(2, gap="large")
                with dl1:
                    st.markdown("""
                    <div class="download-card">
                        <h4>📦 Download sebagai ZIP</h4>
                        <p style="font-size:0.82rem;color:#6b7280;margin-bottom:0.75rem;">Semua file yang telah direname dikemas dalam satu file ZIP.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    zip_buf = create_zip_from_files(st.session_state.rename_mapping, st.session_state.temp_dir)
                    st.download_button(
                        label="📦 Download ZIP (Semua File)", data=zip_buf,
                        file_name="INDOARSIP_Arsip_Renamed.zip", mime="application/zip",
                        use_container_width=True, type="primary"
                    )
                with dl2:
                    st.markdown("""
                    <div class="download-card">
                        <h4>📄 Download File Individual</h4>
                        <p style="font-size:0.82rem;color:#6b7280;margin-bottom:0.75rem;">Unduh file satu per satu sesuai kebutuhan.</p>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("---")
                st.markdown(f"##### 📋 Daftar File ({len(st.session_state.matched_files)} file)")

                for idx, (old_path, new_name) in enumerate(st.session_state.rename_mapping.items(), 1):
                    c_info, c_btn = st.columns([4, 1])
                    with c_info:
                        st.markdown(f"""
                        <div class="file-item">
                            <div class="file-name">{idx}. {new_name}</div>
                            <div class="file-original">Original: {os.path.basename(old_path)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with c_btn:
                        fc = st.session_state.file_contents_cache.get(new_name)
                        if fc:
                            st.download_button(
                                label="⬇️", data=fc, file_name=new_name,
                                mime="application/octet-stream",
                                key=f"dl_{idx}", use_container_width=True
                            )

                st.markdown("---")
                st.markdown("##### 📊 Laporan File Tidak Cocok")
                if st.session_state.unmatched_files:
                    r1, r2 = st.columns([3, 1])
                    with r1:
                        st.warning(f"⚠️ **{len(st.session_state.unmatched_files)} file** tidak cocok dan tidak akan direname.")
                    with r2:
                        st.download_button(
                            label="📄 Download Laporan",
                            data=create_unmatched_report(st.session_state.unmatched_files),
                            file_name="INDOARSIP_Laporan_Tidak_Cocok.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    st.markdown("""
                    <div class="success-notification">
                        <div class="success-icon">✔</div>
                        <div>
                            <p class="success-title">Semua File Berhasil Dicocokkan</p>
                            <p class="success-subtitle">Tidak ada file yang tersisa tanpa pasangan referensi.</p>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <div class="empty-state-title">Tidak ada arsip yang cocok untuk direname</div>
                <div class="empty-state-sub">Periksa kembali data referensi Excel dan nama file arsip kamu.</div>
            </div>
            """, unsafe_allow_html=True)

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("""
<div class="footer">
    <strong>INDOARSIP RENAME</strong> &nbsp;·&nbsp; Sistem Otomatis Penamaan Arsip Digital<br>
    <span style="font-size:0.78rem;">Layanan Penyimpanan & Manajemen Arsip Profesional</span>
</div>
""", unsafe_allow_html=True)