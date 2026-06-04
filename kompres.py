import streamlit as st
import os
import subprocess
import tempfile
import pandas as pd
import pdfplumber

# Pengamanan import comtypes: Hanya di-import jika berjalan di Windows OS
if os.name == 'nt':
    import comtypes.client

# --- FUNGSI UTAMA (BACKEND) ---

def cari_ghostscript():
    if os.name != 'nt': 
        return 'gs'
    try:
        subprocess.run(['gswin64c', '-version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return 'gswin64c'
    except FileNotFoundError:
        pass
    folder_default_gs = r"C:\Program Files\gs"
    if os.path.exists(folder_default_gs):
        for root, dirs, files in os.walk(folder_default_gs):
            if "gswin64c.exe" in files:
                return f'"{os.path.join(root, "gswin64c.exe")}"'
    return None

def kompres_pdf_custom(input_path, output_path, dpi):
    gs_cmd = cari_ghostscript()
    if not gs_cmd:
        return False, "Ghostscript tidak ditemukan."
    
    perintah = (
        f'{gs_cmd} -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH '
        f'-dPDFSETTINGS=/screen '
        f'-dColorImageDownsampleType=/Bicubic -dColorImageResolution={dpi} '
        f'-dGrayImageDownsampleType=/Bicubic -dGrayImageResolution={dpi} '
        f'-dMonoImageDownsampleType=/Bicubic -dMonoImageResolution={dpi} '
        f'-sOutputFile="{output_path}" "{input_path}"'
    )
    try:
        subprocess.run(perintah, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True, "Sukses"
    except subprocess.CalledProcessError as e:
        return False, str(e)

def word_to_pdf(input_path, output_path):
    # JIKA BERJALAN DI WINDOWS (LOKAL)
    if os.name == 'nt':
        try:
            comtypes.CoInitialize()
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            doc.Close(False)
            word.Quit()
            return True, "Sukses"
        except Exception as e:
            try: word.Quit()
            except: pass
            return False, str(e)
    # JIKA BERJALAN DI LINUX (STREAMLIT COMMUNITY CLOUD)
    else:
        try:
            outdir = os.path.dirname(output_path)
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', outdir, input_path
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return True, "Sukses"
        except Exception as e:
            return False, str(e)

def excel_to_pdf(input_path, output_path):
    # JIKA BERJALAN DI WINDOWS (LOKAL)
    if os.name == 'nt':
        try:
            comtypes.CoInitialize()
            excel = comtypes.client.CreateObject('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(input_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
            wb.Close(False)
            excel.Quit()
            return True, "Sukses"
        except Exception as e:
            try: excel.Quit()
            except: pass
            return False, str(e)
    # JIKA BERJALAN DI LINUX (STREAMLIT COMMUNITY CLOUD)
    else:
        try:
            outdir = os.path.dirname(output_path)
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', outdir, input_path
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return True, "Sukses"
        except Exception as e:
            return False, str(e)

# --- DETEKSI BAHASA OTOMATIS ---
bahasa_terdeteksi = "id"
try:
    lang_header = st.context.headers.get("Accept-Language", "en")
    if "id" not in lang_header.lower():
        bahasa_terdeteksi = "en"
except:
    pass

# --- DICTIONARY MULTI-BAHASA ---
KAMUS = {
    "id": {
        "judul_app": "💼 Pro Document Suite",
        "sub_app": "Sistem manajemen dan konversi dokumen terintegrasi dengan performa tinggi.",
        "tab1": "📉 Kompres PDF", "tab2": "📝 Word ➔ PDF", "tab3": "📊 Excel ➔ PDF", "tab4": "📄 PDF ➔ Word", "tab5": "📈 PDF ➔ Excel",
        "side_support": "☕ Apresiasi & Dukungan", "side_desc": "Dukung pengembang agar tetap semangat memperbarui sistem ini!",
        "side_sub": "🔴 Subscribe YouTube", "side_don": "💛 Donasi via Trakteer / Saweria", "side_lang": "🌐 Pilih Bahasa (Language)",
        "comp_title": "Menu Kompresi PDF", "comp_up": "Unggah file PDF yang ingin dikecilkan:", "comp_orig": "📂 Berkas asli:",
        "comp_dpi": "Resolusi (DPI):", "comp_btn": "⚡ Jalankan Kompresi", "comp_spin": "Sistem sedang mengoptimasi...",
        "comp_success": "Selesai dikompres!", "comp_final": "Ukuran Akhir", "comp_dl": "💾 Unduh PDF Hasil",
        "w_title": "Konversi Word (.docx) ke PDF", "w_up": "Unggah dokumen Word Anda:", "w_btn": "⚡ Konversi ke PDF", "w_spin": "Mengonversi dokumen Word...", "w_success": "Konversi Berhasil!", "w_dl": "💾 Unduh File PDF", "w_err": "Gagal konversi. Sistem mendeteksi gangguan software Office.",
        "e_title": "Konversi Excel (.xlsx) ke PDF", "e_up": "Unggah sheet Excel Anda:", "e_btn": "⚡ Konversi ke PDF", "e_spin": "Memproses lembar kerja Excel...", "e_success": "Konversi Berhasil!", "e_dl": "💾 Unduh File PDF", "e_err": "Gagal konversi. Sistem mendeteksi gangguan software Office.",
        "pw_title": "Konversi PDF ke Word (.docx)", "pw_up": "Unggah file PDF untuk dijadikan Word:", "pw_btn": "⚡ Konversi ke Word", "pw_spin": "Mengekstrak teks...", "pw_success": "Konversi Berhasil!", "pw_dl": "💾 Unduh File Word", "pw_err": "Gagal konversi:",
        "pe_title": "Konversi PDF ke Excel (.xlsx)", "pe_up": "Unggah file PDF berisi tabel data:", "pe_btn": "⚡ Konversi ke Excel", "pe_spin": "Mendeteksi tabel data...", "pe_success": "Tabel data berhasil diekstrak!", "pe_dl": "💾 Unduh File Excel", "pe_err": "Gagal mengekstrak data:", "pe_warn": "Tidak dideteksi adanya struktur tabel data numerik."
    },
    "en": {
        "judul_app": "💼 Pro Document Suite",
        "sub_app": "High-performance integrated document conversion and management system.",
        "tab1": "📉 Compress PDF", "tab2": "📝 Word ➔ PDF", "tab3": "📊 Excel ➔ PDF", "tab4": "📄 PDF ➔ Word", "tab5": "📈 PDF ➔ Excel",
        "side_support": "☕ Appreciation & Support", "side_desc": "Support the developer to keep this system running and updated!",
        "side_sub": "🔴 Subscribe YouTube", "side_don": "💛 Donate via SocialBuzz", "side_lang": "🌐 Select Language",
        "comp_title": "PDF Compression Menu", "comp_up": "Upload PDF file to compress:", "comp_orig": "📂 Original file size:",
        "comp_dpi": "Resolution (DPI):", "comp_btn": "⚡ Run Compression", "comp_spin": "System is optimizing...",
        "comp_success": "Successfully compressed!", "comp_final": "Final Size", "comp_dl": "💾 Download Resulting PDF",
        "w_title": "Convert Word (.docx) to PDF", "w_up": "Upload your Word document:", "w_btn": "⚡ Convert to PDF", "w_spin": "Converting Word document...", "w_success": "Conversion Successful!", "w_dl": "💾 Download PDF File", "w_err": "Conversion failed. Office software issue detected.",
        "e_title": "Convert Excel (.xlsx) to PDF", "e_up": "Upload your Excel sheet:", "e_btn": "⚡ Convert to PDF", "e_spin": "Processing Excel sheet...", "e_success": "Conversion Successful!", "e_dl": "💾 Download PDF File", "e_err": "Conversion failed. Office software issue detected.",
        "pw_title": "Convert PDF to Word (.docx)", "pw_up": "Upload PDF file to convert into Word:", "pw_btn": "⚡ Convert to Word", "pw_spin": "Extracting text...", "pw_success": "Conversion Successful!", "pw_dl": "💾 Download Word File", "pw_err": "Conversion failed:",
        "pe_title": "Convert PDF to Excel (.xlsx)", "pe_up": "Upload PDF file containing data tables:", "pe_btn": "⚡ Convert to Excel", "pe_spin": "Detecting data tables...", "pe_success": "Tables successfully extracted!", "pe_dl": "💾 Download Excel File", "pe_err": "Data extraction failed:", "pe_warn": "No numerical table structures were detected in the file."
    }
}

# --- KONFIGURASI HALAMAN STREAMLIT ---
st.set_page_config(page_title="Pro Document Suite", page_icon="💼", layout="centered")

# --- TRICK CSS HYPER-SPECIFICITY (MENGHAPUS TOTAL INFO LIMIT 200MB) ---
st.markdown("""
    <style>
    html body [data-testid="stFileUploaderLimitHint"],
    html body div[data-testid="stFileUploader"] small,
    html body div[data-testid="stFileUploader"] [class*="Limit"],
    html body [data-testid="stFileUploaderDropzone"] + div,
    html body [data-testid="stFileUploaderDropzoneInstructions"] + div,
    html body [data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2),
    html body [data-testid="stFileUploaderDropzoneInstructions"] > span:nth-child(2),
    html body [data-testid="stFileUploaderDropzoneInstructions"] > div:last-child,
    html body div[data-testid="stFileUploader"] > div:last-child {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
        padding: 0px !important;
        margin: 0px !important;
        opacity: 0 !important;
    }
    
    html body [data-testid="stFileUploaderUploadedFiles"] {
        display: block !important;
        visibility: visible !important;
        height: auto !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR NAVIGASI & BAHASA ---
with st.sidebar:
    st.markdown(f"### {KAMUS[bahasa_terdeteksi]['side_lang']}")
    pilihan_bahasa = st.selectbox(
        "Language Selector", 
        ["Auto-Detect", "Bahasa Indonesia", "English"], 
        label_visibility="collapsed"
    )
    
    if pilihan_bahasa == "Bahasa Indonesia":
        lang = "id"
    elif pilihan_bahasa == "English":
        lang = "en"
    else:
        lang = bahasa_terdeteksi
        
    st.divider()
    st.markdown(f"### {KAMUS[lang]['side_support']}")
    st.write(KAMUS[lang]['side_desc'])
    st.divider()
    st.link_button(KAMUS[lang]['side_sub'], url="https://www.youtube.com/@BHG_17", use_container_width=True)
    st.link_button(KAMUS[lang]['side_don'], url="https://sociabuzz.com/tsyndromeg/tribe", use_container_width=True)
    st.divider()
    st.caption("Pro Document Suite v4.1 • Clean Interface")

# --- JUDUL UTAMA ---
st.title(KAMUS[lang]['judul_app'])
st.write(KAMUS[lang]['sub_app'])
st.divider()

# --- MENU NAVIGASI (TABS) ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    KAMUS[lang]['tab1'], KAMUS[lang]['tab2'], KAMUS[lang]['tab3'], KAMUS[lang]['tab4'], KAMUS[lang]['tab5']
])

# ==================== MENU 1: KOMPRES PDF ====================
with tab1:
    st.subheader(KAMUS[lang]['comp_title'])
    up_pdf = st.file_uploader(KAMUS[lang]['comp_up'], type=["pdf"], key="comp_pdf")
    if up_pdf:
        size_awal = len(up_pdf.getvalue()) / (1024 * 1024)
        st.info(f"{KAMUS[lang]['comp_orig']} {size_awal:.2f} MB")
        dpi = st.slider(KAMUS[lang]['comp_dpi'], 60, 200, 110, 10, key="slider_dpi")
        if st.button(KAMUS[lang]['comp_btn'], type="primary", key="btn_comp"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as outf:
                inf.write(up_pdf.getvalue())
                with st.spinner(KAMUS[lang]['comp_spin']):
                    sukses, msg = kompres_pdf_custom(inf.name, outf.name, dpi)
                if sukses:
                    size_akhir = os.path.getsize(outf.name) / (1024 * 1024)
                    st.success(KAMUS[lang]['comp_success'])
                    st.metric(KAMUS[lang]['comp_final'], f"{size_akhir:.2f} MB", f"-{((size_awal-size_akhir)/size_awal)*100:.1f}%")
                    with open(outf.name, "rb") as f:
                        st.download_button(KAMUS[lang]['comp_dl'], f, file_name=f"Compressed_{up_pdf.name}", use_container_width=True)
                else:
                    st.error(f"Gagal: {msg}")

# ==================== MENU 2: WORD TO PDF ====================
with tab2:
    st.subheader(KAMUS[lang]['w_title'])
    up_word = st.file_uploader(KAMUS[lang]['w_up'], type=["docx"], key="w2p")
    if up_word:
        if st.button(KAMUS[lang]['w_btn'], type="primary", key="btn_w2p"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as inf:
                inf.write(up_word.getvalue())
                out_name = inf.name.replace(".docx", ".pdf")
            with st.spinner(KAMUS[lang]['w_spin']):
                sukses, msg = word_to_pdf(inf.name, out_name)
                if sukses:
                    st.success(KAMUS[lang]['w_success'])
                    with open(out_name, "rb") as f:
                        st.download_button(KAMUS[lang]['w_dl'], f, file_name=up_word.name.replace(".docx", ".pdf"), use_container_width=True)
                else:
                    st.error(f"{KAMUS[lang]['w_err']} Info: {msg}")

# ==================== MENU 3: EXCEL TO PDF ====================
with tab3:
    st.subheader(KAMUS[lang]['e_title'])
    up_excel = st.file_uploader(KAMUS[lang]['e_up'], type=["xlsx"], key="e2p")
    if up_excel:
        if st.button(KAMUS[lang]['e_btn'], type="primary", key="btn_e2p"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as inf:
                inf.write(up_excel.getvalue())
                out_name = inf.name.replace(".xlsx", ".pdf")
            with st.spinner(KAMUS[lang]['e_spin']):
                sukses, msg = excel_to_pdf(inf.name, out_name)
                if sukses:
                    st.success(KAMUS[lang]['e_success'])
                    with open(out_name, "rb") as f:
                        st.download_button(KAMUS[lang]['e_dl'], f, file_name=up_excel.name.replace(".xlsx", ".pdf"), use_container_width=True)
                else:
                    st.error(f"{KAMUS[lang]['e_err']} Info: {msg}")

# ==================== MENU 4: PDF TO WORD ====================
with tab4:
    st.subheader(KAMUS[lang]['pw_title'])
    up_pdf_to_w = st.file_uploader(KAMUS[lang]['pw_up'], type=["pdf"], key="p2w")
    if up_pdf_to_w:
        if st.button(KAMUS[lang]['pw_btn'], type="primary", key="btn_p2w"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf:
                inf.write(up_pdf_to_w.getvalue())
                out_name = inf.name.replace(".pdf", ".docx")
            with st.spinner(KAMUS[lang]['pw_spin']):
                try:
                    from pdf2docx import Converter
                    cv = Converter(inf.name)
                    cv.convert(out_name, start=0, end=None)
                    cv.close()
                    st.success(KAMUS[lang]['pw_success'])
                    with open(out_name, "rb") as f:
                        st.download_button(KAMUS[lang]['pw_dl'], f, file_name=up_pdf_to_w.name.replace(".pdf", ".docx"), use_container_width=True)
                except Exception as e:
                    st.error(f"{KAMUS[lang]['pw_err']} {e}")

# ==================== MENU 5: PDF TO EXCEL ====================
with tab5:
    st.subheader(KAMUS[lang]['pe_title'])
    up_pdf_to_e = st.file_uploader(KAMUS[lang]['pe_up'], type=["pdf"], key="p2e")
    if up_pdf_to_e:
        if st.button(KAMUS[lang]['pe_btn'], type="primary", key="btn_p2e"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf:
                inf.write(up_pdf_to_e.getvalue())
                out_name = inf.name.replace(".pdf", ".xlsx")
            with st.spinner(KAMUS[lang]['pe_spin']):
                try:
                    all_tables = []
                    with pdfplumber.open(inf.name) as pdf:
                        for page in pdf.pages:
                            extracted = page.extract_tables()
                            for table in extracted:
                                all_tables.append(pd.DataFrame(table))
                    if all_tables:
                        with pd.ExcelWriter(out_name) as writer:
                            for idx, df in enumerate(all_tables):
                                df.to_excel(writer, sheet_name=f"Page_{idx+1}", index=False, header=False)
                        st.success(KAMUS[lang]['pe_success'])
                        with open(out_name, "rb") as f:
                            st.download_button(KAMUS[lang]['pe_dl'], f, file_name=up_pdf_to_e.name.replace(".pdf", ".xlsx"), use_container_width=True)
                    else:
                        st.warning(KAMUS[lang]['pe_warn'])
                except Exception as e:
                    st.error(f"{KAMUS[lang]['pe_err']} {e}")
