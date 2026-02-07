"""
INDOARSIP - Sistem Otomatis Penamaan Arsip Digital
Batch File Renaming System based on Excel Reference

Author: Claude AI
Company: INDOARSIP
"""

import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import shutil
from pathlib import Path
from io import BytesIO

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="INDOARSIP - Sistem Rename Arsip",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================================
# CUSTOM CSS - INDOARSIP BRANDING
# ============================================================================

st.markdown("""
<style>
    /* Main branding */
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 2px;
    }
    .main-header p {
        color: #e0e7ff;
        font-size: 1.1rem;
        margin: 0.5rem 0 0 0;
        font-weight: 300;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: white;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        border: 2px solid transparent;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2a5298;
        color: white;
        border-color: #1e3c72;
    }
    
    /* Button styling */
    .stButton button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        font-weight: 600;
        padding: 0.75rem 2rem;
        border-radius: 6px;
        border: none;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    .stButton button:hover {
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        transform: translateY(-2px);
    }
    
    /* Info boxes */
    .info-box {
        background-color: #f0f4f8;
        padding: 1.5rem;
        border-radius: 8px;
        border-left: 4px solid #2a5298;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #f0fdf4;
        border-left-color: #22c55e;
    }
    .warning-box {
        background-color: #fef3c7;
        border-left-color: #f59e0b;
    }
    .error-box {
        background-color: #fef2f2;
        border-left-color: #ef4444;
    }
    
    /* DataFrames */
    .dataframe {
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER SECTION
# ============================================================================

st.markdown("""
<div class="main-header">
    <h1>📦 INDOARSIP</h1>
    <p>Sistem Otomatis Penamaan Arsip Digital</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================

if 'validated' not in st.session_state:
    st.session_state.validated = False
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None
if 'file_list' not in st.session_state:
    st.session_state.file_list = []
if 'reference_data' not in st.session_state:
    st.session_state.reference_data = None
if 'matched_files' not in st.session_state:
    st.session_state.matched_files = []
if 'unmatched_files' not in st.session_state:
    st.session_state.unmatched_files = []
if 'rename_mapping' not in st.session_state:
    st.session_state.rename_mapping = {}
if 'show_individual_files' not in st.session_state:
    st.session_state.show_individual_files = False

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def extract_zip(zip_file):
    """Extract ZIP file to temporary directory"""
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir

def get_files_from_directory(directory):
    """Get all files from directory (including subdirectories)"""
    files = []
    for root, dirs, filenames in os.walk(directory):
        # Skip __MACOSX and other system directories
        dirs[:] = [d for d in dirs if not d.startswith('.') and d != '__MACOSX']
        
        for filename in filenames:
            # Skip hidden files, system files, and macOS metadata
            if (not filename.startswith('.') and 
                not filename.startswith('__') and
                filename != '.DS_Store' and
                not root.endswith('__MACOSX')):
                file_path = os.path.join(root, filename)
                # Double check it's actually a file
                if os.path.isfile(file_path):
                    files.append(file_path)
    return files

def extract_code_from_filename(filename):
    """Extract code from filename (e.g., '0336' from 'file_pelanggan_0336')"""
    # Remove extension
    name_without_ext = os.path.splitext(filename)[0]
    
    # Try to find numeric code patterns
    import re
    
    # Pattern 1: Numbers at the end (e.g., file_pelanggan_0336 -> 0336)
    match = re.search(r'_(\d+)$', name_without_ext)
    if match:
        return match.group(1)
    
    # Pattern 2: Numbers at the beginning (e.g., 0336_document -> 0336)
    match = re.search(r'^(\d+)', name_without_ext)
    if match:
        return match.group(1)
    
    # Pattern 3: Any sequence of digits (fallback)
    match = re.search(r'(\d+)', name_without_ext)
    if match:
        return match.group(1)
    
    # If no code found, return the whole name
    return name_without_ext

def match_files_with_reference(file_list, reference_values):
    """Match filenames with reference values using prefix-based matching"""
    matched = []
    unmatched = []
    rename_map = {}
    
    # Get base filenames without extension
    for file_path in file_list:
        filename = os.path.basename(file_path)
        file_extension = os.path.splitext(filename)[1]
        
        # Extract code from filename
        file_code = extract_code_from_filename(filename)
        
        # Check if code matches any reference value (prefix matching)
        match_found = False
        for ref_val in reference_values:
            ref_str = str(ref_val).strip()
            
            # Method 1: Check if reference starts with the file code
            # Example: file code "0336" matches "0336-PT. CONTAINER MARITIME ACTIVITIES"
            if ref_str.startswith(file_code):
                matched.append(file_path)
                # Create new filename: reference value + original extension
                new_filename = ref_str + file_extension
                rename_map[file_path] = new_filename
                match_found = True
                break
        
        if not match_found:
            unmatched.append(file_path)
    
    return matched, unmatched, rename_map

def create_zip_from_files(file_mapping, original_dir):
    """Create ZIP file from renamed files"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for old_path, new_name in file_mapping.items():
            # Write file with NEW name (not old path basename)
            # old_path = full path to original file
            # new_name = the new filename we want
            zip_file.write(old_path, arcname=new_name)
    zip_buffer.seek(0)
    return zip_buffer

def create_unmatched_report(unmatched_files):
    """Create Excel report for unmatched files"""
    data = {
        'Nama File Tidak Cocok': [os.path.basename(f) for f in unmatched_files],
        'Path Lengkap': unmatched_files,
        'Status': ['Tidak Ditemukan di Referensi'] * len(unmatched_files)
    }
    df = pd.DataFrame(data)
    
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Arsip Tidak Cocok', index=False)
    excel_buffer.seek(0)
    return excel_buffer

# ============================================================================
# TAB STRUCTURE
# ============================================================================

tab1, tab2 = st.tabs(["📋 Upload & Validasi Arsip", "✅ Preview & Proses Rename"])

# ============================================================================
# TAB 1: UPLOAD & VALIDASI ARSIP
# ============================================================================

with tab1:
    st.markdown("### Upload & Validasi Data Arsip")
    st.info("ℹ️ **Cara kerja:** Sistem bakal ngecek nomor di nama file terus cocokkin sama data Excel. Misalnya file `file_pelanggan_0001.pdf` bakal ketemu sama data yang mulai dari `0001-...`")
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 1️⃣ Upload File Arsip")
        upload_type = st.radio(
            "Pilih tipe upload:",
            ["File ZIP Arsip", "Folder Arsip (Multiple Files)"],
            key="upload_type"
        )
        
        if upload_type == "File ZIP Arsip":
            zip_file = st.file_uploader(
                "Upload file ZIP yang berisi arsip",
                type=['zip'],
                key="zip_uploader",
                help="Hanya file ZIP yang bisa diupload di sini"
            )
        else:
            uploaded_files = st.file_uploader(
                "Upload multiple files arsip (BUKAN ZIP)",
                accept_multiple_files=True,
                key="files_uploader",
                help="Upload file individual, bukan file ZIP. Untuk ZIP gunakan opsi di atas."
            )
            
            # Check if user uploaded ZIP in wrong option
            if uploaded_files:
                zip_files_found = [f for f in uploaded_files if f.name.lower().endswith('.zip')]
                if zip_files_found:
                    st.error(f"❌ **File ZIP terdeteksi: {', '.join([f.name for f in zip_files_found])}**")
                    st.warning("⚠️ Untuk upload file ZIP, gunakan opsi **'File ZIP Arsip'** di atas ya!")
                    uploaded_files = None  # Reset uploaded files
    
    with col2:
        st.markdown("#### 2️⃣ Upload File Referensi Excel")
        excel_file = st.file_uploader(
            "Upload file Excel referensi penamaan",
            type=['xlsx', 'xls'],
            key="excel_uploader"
        )
        
        reference_column = st.text_input(
            "Nama Kolom Referensi Arsip",
            placeholder="Contoh: Nomor_Arsip, Kode_Dokumen, dll",
            key="ref_column"
        )
    
    st.markdown("---")
    
    # Validation button
    if st.button("🔍 Validasi & Cek Arsip", use_container_width=True, type="primary"):
        # Validation checks
        errors = []
        
        if upload_type == "File ZIP Arsip" and not zip_file:
            errors.append("File ZIP arsip belum diupload")
        elif upload_type == "Folder Arsip (Multiple Files)":
            if not uploaded_files:
                errors.append("File arsip belum diupload")
            else:
                # Check if user uploaded ZIP files in multiple files mode
                zip_files_found = [f.name for f in uploaded_files if f.name.lower().endswith('.zip')]
                if zip_files_found:
                    errors.append(f"File ZIP tidak bisa diupload di opsi Multiple Files! Gunakan opsi 'File ZIP Arsip' untuk: {', '.join(zip_files_found)}")
        
        if not excel_file:
            errors.append("File Excel referensi belum diupload")
        
        if not reference_column or reference_column.strip() == "":
            errors.append("Nama kolom referensi belum diisi")
        
        if errors:
            st.error("❌ **Validasi Gagal**")
            for error in errors:
                st.markdown(f"- {error}")
        else:
            with st.spinner("Memproses validasi data..."):
                try:
                    # Step 1: Extract files
                    if upload_type == "File ZIP Arsip":
                        temp_dir = extract_zip(zip_file)
                        file_list = get_files_from_directory(temp_dir)
                        
                        if len(file_list) == 0:
                            st.error("❌ **ZIP kosong atau tidak ada file yang valid!**")
                            st.warning("🔍 **Kemungkinan penyebab:**")
                            st.markdown("""
                            - ZIP kosong (tidak ada file)
                            - Semua file adalah hidden files (diawali titik)
                            - File berada dalam folder `__MACOSX` (metadata macOS)
                            - Coba extract manual dulu untuk mengecek isi ZIP
                            """)
                            
                            # Debug info: show what's in temp_dir
                            all_items = []
                            for root, dirs, files in os.walk(temp_dir):
                                for f in files:
                                    all_items.append(os.path.join(root, f))
                            
                            if all_items:
                                with st.expander("🐛 Debug: Lihat semua item yang di-extract (termasuk hidden files)"):
                                    st.code('\n'.join([os.path.basename(item) for item in all_items]))
                            
                            st.session_state.validated = False
                            continue_validation = False
                        else:
                            # Show extracted files info
                            st.info(f"📦 **ZIP berhasil di-extract!** Ditemukan {len(file_list)} file")
                            with st.expander("📂 Lihat file hasil extract dari ZIP"):
                                extracted_df = pd.DataFrame({
                                    'No': list(range(1, len(file_list) + 1)),
                                    'Nama File': [os.path.basename(f) for f in file_list],
                                    'Lokasi': [os.path.relpath(f, temp_dir) for f in file_list],
                                    'Ukuran': [f"{os.path.getsize(f) / 1024:.2f} KB" for f in file_list]
                                })
                                st.dataframe(extracted_df, use_container_width=True)
                            continue_validation = True
                    else:
                        temp_dir = tempfile.mkdtemp()
                        file_list = []
                        for uploaded_file in uploaded_files:
                            file_path = os.path.join(temp_dir, uploaded_file.name)
                            with open(file_path, 'wb') as f:
                                f.write(uploaded_file.getbuffer())
                            file_list.append(file_path)
                        
                        st.info(f"📁 **File berhasil diupload!** Total {len(file_list)} file")
                        continue_validation = True
                    
                    st.session_state.temp_dir = temp_dir
                    
                    if not continue_validation:
                        pass  # Stop here, error already shown
                    else:
                        # Step 2: Read Excel and validate column
                        df = pd.read_excel(excel_file)
                        
                        if reference_column not in df.columns:
                            st.error(f"❌ Kolom '{reference_column}' tidak ditemukan dalam file Excel!")
                            st.info(f"📋 Kolom yang tersedia: {', '.join(df.columns.tolist())}")
                            st.session_state.validated = False
                        else:
                            # Step 3: Get reference values
                            reference_values = df[reference_column].dropna().astype(str).tolist()
                            
                            # Step 4: Match files
                            matched, unmatched, rename_map = match_files_with_reference(
                                file_list, reference_values
                            )
                            
                            # Store in session state
                            st.session_state.file_list = file_list
                            st.session_state.reference_data = df
                            st.session_state.matched_files = matched
                            st.session_state.unmatched_files = unmatched
                            st.session_state.rename_mapping = rename_map
                            st.session_state.validated = True
                            
                            # Display results
                            st.success("✅ **Validasi Berhasil!**")
                            
                            st.markdown("---")
                            st.markdown("### 📊 Ringkasan Validasi Arsip")
                            
                            col_a, col_b, col_c = st.columns(3)
                            
                            with col_a:
                                st.metric(
                                    label="Total Arsip",
                                    value=len(file_list),
                                    delta=None
                                )
                            
                            with col_b:
                                st.metric(
                                    label="Arsip Cocok",
                                    value=len(matched),
                                    delta=f"{(len(matched)/len(file_list)*100):.1f}%" if file_list else "0%",
                                    delta_color="normal"
                                )
                            
                            with col_c:
                                st.metric(
                                    label="Arsip Tidak Cocok",
                                    value=len(unmatched),
                                    delta=f"{(len(unmatched)/len(file_list)*100):.1f}%" if file_list else "0%",
                                    delta_color="inverse"
                                )
                            
                            # Detail information
                            if matched:
                                with st.expander(f"✅ Lihat {len(matched)} arsip yang cocok"):
                                    matched_df = pd.DataFrame({
                                        'No': list(range(1, len(matched) + 1)),
                                        'Nama File Asli': [os.path.basename(f) for f in matched],
                                        'Kode Ekstrak': [extract_code_from_filename(os.path.basename(f)) for f in matched],
                                        'Akan Direname Jadi': [rename_map[f] for f in matched]
                                    })
                                    st.dataframe(matched_df, use_container_width=True)
                            
                            if unmatched:
                                with st.expander(f"⚠️ Lihat {len(unmatched)} arsip yang tidak cocok"):
                                    unmatched_df = pd.DataFrame({
                                        'Nama File': [os.path.basename(f) for f in unmatched]
                                    })
                                    st.dataframe(unmatched_df, use_container_width=True)
                            
                            st.info("✅ Data siap diproses. Lanjut ke tab **Preview & Proses Rename** ya")
                
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    st.session_state.validated = False

# ============================================================================
# TAB 2: PREVIEW & PROSES RENAME
# ============================================================================

with tab2:
    st.markdown("### Preview & Eksekusi Rename Arsip")
    st.markdown("---")
    
    if not st.session_state.validated:
        st.warning("⚠️ **Isi dulu validasi di Tab 1 ya**")
        st.info("📋 Upload file arsip sama file Excel referensi dulu, terus klik tombol **Validasi & Cek Arsip**")
    else:
        # Display summary
        st.markdown("### 📊 Ringkasan Proses Rename")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Arsip", len(st.session_state.file_list))
        with col2:
            st.metric("Akan Direname", len(st.session_state.matched_files))
        with col3:
            st.metric("Tidak Cocok", len(st.session_state.unmatched_files))
        
        st.markdown("---")
        
        # Preview table
        if st.session_state.matched_files:
            st.markdown("### 👁️ Preview Penamaan Arsip")
            
            preview_data = {
                'No': list(range(1, len(st.session_state.matched_files) + 1)),
                'Nama Arsip Lama': [os.path.basename(f) for f in st.session_state.matched_files],
                'Nama Arsip Baru': [st.session_state.rename_mapping[f] for f in st.session_state.matched_files],
                'Status': ['✅ Siap Rename'] * len(st.session_state.matched_files)
            }
            preview_df = pd.DataFrame(preview_data)
            st.dataframe(preview_df, use_container_width=True)
            
            st.markdown("---")
            
            # Rename button
            if st.button("🚀 Mulai Proses Rename Arsip", use_container_width=True, type="primary"):
                with st.spinner("Memproses rename arsip..."):
                    try:
                        # Use the rename_mapping directly (already contains new names)
                        # st.session_state.rename_mapping format:
                        # {old_file_path: new_filename_with_extension}
                        
                        st.success("✅ **Proses Rename Selesai!**")
                        st.balloons()
                        
                        # Download section
                        st.markdown("---")
                        st.markdown("### 📥 Download Hasil Rename")
                        
                        # Create two columns for download options
                        download_col1, download_col2 = st.columns(2)
                        
                        with download_col1:
                            st.markdown("#### 📦 Opsi 1: Download sebagai ZIP")
                            st.info("Semua file jadi satu dalam ZIP")
                            
                            # Create ZIP for renamed files using the correct mapping
                            zip_buffer = create_zip_from_files(
                                st.session_state.rename_mapping,  # This already has old_path -> new_name
                                st.session_state.temp_dir
                            )
                            
                            st.download_button(
                                label="📦 Download ZIP Arsip (Semua File)",
                                data=zip_buffer,
                                file_name="INDOARSIP_Arsip_Renamed.zip",
                                mime="application/zip",
                                use_container_width=True,
                                type="primary"
                            )
                        
                        with download_col2:
                            st.markdown("#### 📄 Opsi 2: Download File Individual")
                            st.info("Download satu-satu sesuai kebutuhan")
                        
                        # Show individual file download list DIRECTLY (no button needed)
                        st.markdown("---")
                        st.markdown("### 📋 Daftar File yang Bisa Didownload")
                        st.caption(f"Total ada {len(st.session_state.matched_files)} file")
                        
                        # Create scrollable container for file list
                        for idx, (old_path, new_name) in enumerate(st.session_state.rename_mapping.items(), 1):
                            with st.container():
                                col_file, col_btn = st.columns([3, 1])
                                
                                with col_file:
                                    st.markdown(f"**{idx}.** `{new_name}`")
                                    st.caption(f"Original: {os.path.basename(old_path)}")
                                
                                with col_btn:
                                    # Read file content
                                    with open(old_path, 'rb') as f:
                                        file_content = f.read()
                                    
                                    st.download_button(
                                        label="⬇️ Download",
                                        data=file_content,
                                        file_name=new_name,
                                        mime="application/octet-stream",
                                        key=f"download_individual_{idx}",
                                        use_container_width=True
                                    )
                            
                            if idx < len(st.session_state.rename_mapping):
                                st.divider()
                        
                        # Excel report section (always available)
                        st.markdown("---")
                        st.markdown("### 📊 Laporan File Tidak Cocok")
                        
                        if st.session_state.unmatched_files:
                            col_report1, col_report2 = st.columns([2, 1])
                            
                            with col_report1:
                                st.warning(f"⚠️ Ada **{len(st.session_state.unmatched_files)} file** yang nggak cocok sama data referensi")
                                st.caption("File-file ini nggak akan direname dan udah dicatat di laporan Excel")
                            
                            with col_report2:
                                # Create Excel report for unmatched files
                                excel_buffer = create_unmatched_report(
                                    st.session_state.unmatched_files
                                )
                                
                                st.download_button(
                                    label="📄 Download Laporan Excel",
                                    data=excel_buffer,
                                    file_name="INDOARSIP_Laporan_Tidak_Cocok.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        else:
                            st.success("✅ Semua file berhasil cocok! Mantap")
                    
                    except Exception as e:
                        st.error(f"❌ Terjada kesalahan saat proses rename: {str(e)}")
        else:
            st.warning("⚠️ Tidak ada arsip yang cocok untuk direname")

# ============================================================================
# FOOTER
# ============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6b7280; padding: 2rem 0;'>
    <p style='margin: 0;'><strong>INDOARSIP</strong> - Sistem Otomatis Penamaan Arsip Digital</p>
    <p style='margin: 0.5rem 0 0 0; font-size: 0.9rem;'>Layanan Penyimpanan & Manajemen Arsip Profesional</p>
</div>
""", unsafe_allow_html=True)