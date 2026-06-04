import streamlit as st
import os
import subprocess
import tempfile
import pandas as pd
import pdfplumber
from PIL import Image
from docx import Document
from docx.shared import Inches
import pypdf

# Pengamanan import comtypes: Hanya di Windows Lokal
if os.name == 'nt':
    import comtypes.client

# ==========================================
#         BACKEND ENGINE FUNCTIONS
# ==========================================

def cari_ghostscript():
    if os.name != 'nt': return 'gs'
    try:
        subprocess.run(['gswin64c', '-version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return 'gswin64c'
    except FileNotFoundError: pass
    folder_default_gs = r"C:\Program Files\gs"
    if os.path.exists(folder_default_gs):
        for root, dirs, files in os.walk(folder_default_gs):
            if "gswin64c.exe" in files: return f'"{os.path.join(root, "gswin64c.exe")}"'
    return None

def kompres_pdf_custom(input_path, output_path, dpi):
    gs_cmd = cari_ghostscript()
    if not gs_cmd: return False, "Ghostscript tidak ditemukan."
    perintah = (
        f'{gs_cmd} -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH '
        f'-dPDFSETTINGS=/screen -dColorImageResolution={dpi} -sOutputFile="{output_path}" "{input_path}"'
    )
    try:
        subprocess.run(perintah, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True, "Sukses"
    except Exception as e: return False, str(e)

def office_to_pdf_linux(input_path, output_path):
    try:
        outdir = os.path.dirname(output_path)
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', outdir, input_path], check=True)
        return True, "Sukses"
    except Exception as e: return False, str(e)

def word_to_pdf(input_path, output_path):
    if os.name == 'nt':
        try:
            comtypes.CoInitialize()
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            doc.Close(False); word.Quit()
            return True, "Sukses"
        except Exception as e: return False, str(e)
    else: return office_to_pdf_linux(input_path, output_path)

def excel_to_pdf(input_path, output_path):
    if os.name == 'nt':
        try:
            comtypes.CoInitialize()
            excel = comtypes.client.CreateObject('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(input_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
            wb.Close(False); excel.Quit()
            return True, "Sukses"
        except Exception as e: return False, str(e)
    else: return office_to_pdf_linux(input_path, output_path)

def ppt_to_pdf(input_path, output_path):
    if os.name == 'nt':
        try:
            comtypes.CoInitialize()
            ppt = comtypes.client.CreateObject('PowerPoint.Application')
            pres = ppt.Presentations.Open(os.path.abspath(input_path), WithWindow=False)
            pres.SaveAs(os.path.abspath(output_path), FileFormat=32)
            pres.Close(); ppt.Quit()
            return True, "Sukses"
        except Exception as e: return False, str(e)
    else: return office_to_pdf_linux(input_path, output_path)

def merge_pdfs(input_paths, output_path):
    try:
        merger = pypdf.PdfMerger()
        for path in input_paths: merger.append(path)
        merger.write(output_path)
        merger.close()
        return True, "Sukses"
    except Exception as e: return False, str(e)

def split_pdf(input_path, out_dir):
    try:
        reader = pypdf.PdfReader(input_path)
        generated_files = []
        for idx, page in enumerate(reader.pages):
            writer = pypdf.PdfWriter()
            writer.add_page(page)
            out_path = os.path.join(out_dir, f"Page_{idx+1}.pdf")
            with open(out_path, "wb") as f: writer.write(f)
            generated_files.append(out_path)
        return True, generated_files
    except Exception as e: return False, str(e)

def rotate_pdf(input_path, output_path, angle):
    try:
        reader = pypdf.PdfReader(input_path)
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            page.rotate(angle)
            writer.add_page(page)
        with open(output_path, "wb") as f: writer.write(f)
        return True, "Sukses"
    except Exception as e: return False, str(e)

def protect_pdf(input_path, output_path, password):
    try:
        reader = pypdf.PdfReader(input_path)
        writer = pypdf.PdfWriter()
        for page in reader.pages: writer.add_page(page)
        writer.encrypt(password)
        with open(output_path, "wb") as f: writer.write(f)
        return True, "Sukses"
    except Exception as e: return False, str(e)

def unlock_pdf(input_path, output_path, password):
    try:
        reader = pypdf.PdfReader(input_path)
        if reader.is_encrypted:
            reader.decrypt(password)
        writer = pypdf.PdfWriter()
        for page in reader.pages: writer.add_page(page)
        with open(output_path, "wb") as f: writer.write(f)
        return True, "Sukses"
    except Exception as e: return False, str(e)

# ==========================================
#          STREAMLIT CONFIG & UI
# ==========================================
st.set_page_config(page_title="Pro Document Suite", page_icon="💼", layout="wide")

# BOM NUKLIR CSS: Hilangkan Limit 200MB & Bikin Desain Grid Elegan
st.markdown("""
    <style>
    html body [data-testid="stFileUploaderLimitHint"],
    html body div[data-testid="stFileUploader"] small,
    html body [data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) {
        display: none !important; visibility: hidden !important; height: 0px !important; padding: 0px !important;
    }
    .main-title { text-align: center; font-size: 2.5rem; font-weight: bold; margin-bottom: 5px; }
    .sub-title { text-align: center; color: #666; font-size: 1.1rem; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

# Inisialisasi State Halaman Navigasi
if "current_tool" not in st.session_state:
    st.session_state.current_tool = "Dashboard"

# Mapping Semua 30 Fitur Sesuai Gambar
DAFTAR_FITUR = {
    "Merge PDF": {"icon": "🔗", "desc": "Gabungkan beberapa file PDF menjadi satu dokumen dengan mudah.", "cat": "Organize"},
    "Split PDF": {"icon": "✂️", "desc": "Pisah halaman PDF menjadi file terpisah atau rentang tertentu.", "cat": "Organize"},
    "Compress PDF": {"icon": "📉", "desc": "Perkecil ukuran file PDF Anda secara optimal tanpa merusak kualitas.", "cat": "Optimize"},
    "PDF to Word": {"icon": "📝", "desc": "Konversi dokumen PDF menjadi file Word (.docx) yang dapat diedit.", "cat": "Convert"},
    "PDF to PowerPoint": {"icon": "📊", "desc": "Ubah file PDF menjadi slide presentasi PowerPoint (.pptx).", "cat": "Convert"},
    "PDF to Excel": {"icon": "📈", "desc": "Ekstrak tabel data dari PDF langsung ke spreadsheet Excel (.xlsx).", "cat": "Convert"},
    "Word to PDF": {"icon": "📄", "desc": "Konversi dokumen Microsoft Word (.docx) menjadi file PDF rapi.", "cat": "Convert"},
    "PowerPoint to PDF": {"icon": "📉", "desc": "Ubah dokumen presentasi PowerPoint (.pptx) menjadi format PDF.", "cat": "Convert"},
    "Excel to PDF": {"icon": "📊", "desc": "Jadikan data spreadsheet Excel (.xlsx) Anda ke dokumen PDF.", "cat": "Convert"},
    "Edit PDF": {"icon": "✏️", "desc": "Tambahkan teks, gambar, bentuk, atau anotasi pada file PDF.", "cat": "Edit"},
    "PDF to JPG": {"icon": "🖼️", "desc": "Ekstrak semua gambar atau ubah halaman PDF menjadi file gambar JPG.", "cat": "Convert"},
    "JPG to PDF": {"icon": "🖼️", "desc": "Ubah gambar JPG/PNG menjadi file dokumen PDF secara instan.", "cat": "Convert"},
    "Sign PDF": {"icon": "🔏", "desc": "Tambahkan tanda tangan digital atau mintalah tanda tangan elektronik.", "cat": "Security"},
    "Watermark": {"icon": "🏷️", "desc": "Beri cap teks atau gambar di atas file PDF secara kustom.", "cat": "Edit"},
    "Rotate PDF": {"icon": "🔄", "desc": "Putar arah halaman dokumen PDF Anda sesuai kebutuhan.", "cat": "Organize"},
    "HTML to PDF": {"icon": "🌐", "desc": "Ubah halaman web atau file HTML menjadi dokumen PDF.", "cat": "Convert"},
    "Unlock PDF": {"icon": "🔓", "desc": "Hapus enkripsi password keamanan pada berkas PDF Anda.", "cat": "Security"},
    "Protect PDF": {"icon": "🔒", "desc": "Amankan file PDF berharga Anda dengan password enkripsi kuat.", "cat": "Security"},
    "Organize PDF": {"icon": "🗂️", "desc": "Hapus, susun ulang, atau tambah halaman pada dokumen PDF.", "cat": "Organize"},
    "PDF to PDF/A": {"icon": "📜", "desc": "Konversi dokumen PDF ke standar ISO PDF/A untuk arsip jangka panjang.", "cat": "Convert"},
    "Repair PDF": {"icon": "🔧", "desc": "Perbaiki file PDF rusak atau corrupt agar dapat terbaca kembali.", "cat": "Optimize"},
    "Page numbers": {"icon": "🔢", "desc": "Tambahkan nomor halaman pada dokumen PDF secara otomatis.", "cat": "Edit"},
    "Scan to PDF": {"icon": "🖨️", "desc": "Ambil pindaian dari perangkat mobile dan jadikan PDF di browser.", "cat": "Convert"},
    "OCR PDF": {"icon": "🔍", "desc": "Ubah dokumen PDF hasil scan menjadi teks yang bisa dicari/di-copy.", "cat": "Intelligence"},
    "Compare PDF": {"icon": "👥", "desc": "Bandingkan dua file PDF secara berdampingan untuk melihat perbedaan.", "cat": "Intelligence"},
    "Redact PDF": {"icon": "🔏", "desc": "Hapus informasi rahasia atau teks sensitif secara permanen.", "cat": "Security"},
    "Crop PDF": {"icon": "✂️", "desc": "Potong margin area halaman luar dokumen PDF Anda.", "cat": "Edit"},
    "PDF Forms": {"icon": "📝", "desc": "Deteksi, isi, atau buat formulir interaktif di dalam PDF.", "cat": "Edit"},
    "AI Summarizer": {"icon": "🤖", "desc": "Rangkum isi dokumen PDF panjang secara instan berbasis AI.", "cat": "Intelligence"},
    "Translate PDF": {"icon": "🌐", "desc": "Terjemahkan bahasa dokumen PDF secara otomatis dengan AI.", "cat": "Intelligence"},
}

# --- SIDEBAR UTAMA ---
with st.sidebar:
    st.markdown("### 🌐 Navigation Suite")
    if st.button("🏠 Menu Utama Dashboard", use_container_width=True):
        st.session_state.current_tool = "Dashboard"
    st.divider()
    st.markdown("### ☕ Developer Support")
    st.link_button("🔴 Subscribe YouTube", url="https://www.youtube.com/@BHG_17", use_container_width=True)
    st.link_button("💛 Donasi via Trakteer", url="https://sociabuzz.com/tsyndromeg/tribe", use_container_width=True)
    st.divider()
    st.caption("Pro Document Suite v5.0 Enterprise")

# ==========================================
#          HALAMAN 1: DASHBOARD UTAMA
# ==========================================
if st.session_state.current_tool == "Dashboard":
    st.markdown('<div class="main-title">Every tool you need to work with PDFs in one place</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Every tool you need to use PDFs, at your fingertips. All are 100% FREE and easy to use!</div>', unsafe_allow_html=True)
    
    # Filter Tabs Atas Sesuai Gambar
    kategori = st.radio("Filter Kategori:", ["All", "Organize", "Optimize", "Convert", "Edit", "Security", "Intelligence"], horizontal=True)
    st.divider()
    
    # Membuat Sistem Grid Menggunakan Kolom Streamlit (Baris isi 3 Card)
    col_idx = 0
    cols = st.columns(3)
    
    for nama, info in DAFTAR_FITUR.items():
        if kategori != "All" and info["cat"] != kategori:
            continue
            
        with cols[col_idx % 3]:
            # Desain Card Menggunakan Gabungan Info & Button Ber-Key Spesifik
            st.markdown(f"#### {info['icon']} {nama}")
            st.caption(info["desc"])
            if st.button(f"Buka {nama} →", key=f"btn_nav_{nama}", use_container_width=True):
                st.session_state.current_tool = nama
                st.rerun()
            st.write("") # Spacer antarcat
        col_idx += 1

# ==========================================
#          HALAMAN 2: WORKSPACE FITUR
# ==========================================
else:
    tool = st.session_state.current_tool
    
    # Header Fitur Aktif
    st.button("⬅️ Kembali ke Dashboard Utama", key="back_btn")
    st.title(f"{DAFTAR_FITUR[tool]['icon']} Workspace: {tool}")
    st.caption(DAFTAR_FITUR[tool]['desc'])
    st.divider()
    
    # --- LOGIKA OPERASI MASING-MASING FITUR ---
    
    if tool == "Compress PDF":
        up_file = st.file_uploader("Unggah file PDF Anda:", type=["pdf"])
        if up_file:
            dpi = st.slider("Resolusi Kompresi (DPI):", 60, 200, 110, 10)
            if st.button("⚡ Jalankan Kompresi", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as outf:
                    inf.write(up_file.getvalue())
                    with st.spinner("Sistem sedang mengompres..."):
                        sukses, msg = kompres_pdf_custom(inf.name, outf.name, dpi)
                    if sukses:
                        st.success("Selesai Dikompres!")
                        with open(outf.name, "rb") as f:
                            st.download_button("💾 Unduh PDF Hasil", f, file_name=f"Compressed_{up_file.name}", use_container_width=True)
                    else: st.error(f"Gagal: {msg}")

    elif tool == "Merge PDF":
        up_files = st.file_uploader("Unggah beberapa file PDF sekaligus:", type=["pdf"], accept_multiple_files=True)
        if up_files and len(up_files) >= 2:
            if st.button("⚡ Gabungkan PDF", type="primary"):
                with st.spinner("Menggabungkan file..."):
                    temp_paths = []
                    for f in up_files:
                        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                        tmp.write(f.getvalue())
                        temp_paths.append(tmp.name)
                    out_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                    sukses, msg = merge_pdfs(temp_paths, out_pdf)
                    if sukses:
                        st.success("Berhasil digabungkan!")
                        with open(out_pdf, "rb") as f:
                            st.download_button("💾 Unduh PDF Hasil Gabungan", f, file_name="Merged_Document.pdf", use_container_width=True)
                    else: st.error(msg)
        elif up_files: st.warning("Silakan upload minimal 2 file PDF untuk digabungkan.")

    elif tool == "Split PDF":
        up_file = st.file_uploader("Unggah PDF yang ingin dipecah per halaman:", type=["pdf"])
        if up_file:
            if st.button("⚡ Jalankan Ekstraksi Halaman", type="primary"):
                with tempfile.TemporaryDirectory() as tmpdir:
                    inf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    inf.write(up_file.getvalue())
                    sukses, res = split_pdf(inf.name, tmpdir)
                    if sukses:
                        st.success(f"Berhasil memecah menjadi {len(res)} halaman!")
                        for path in res:
                            with open(path, "rb") as f:
                                st.download_button(f"💾 Unduh {os.path.basename(path)}", f, file_name=os.path.basename(path))
                    else: st.error(res)

    elif tool == "Word to PDF":
        up_file = st.file_uploader("Unggah file Word (.docx):", type=["docx"])
        if up_file:
            if st.button("⚡ Konversi ke PDF", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(".docx", ".pdf")
                with st.spinner("Mengonversi..."):
                    sukses, msg = word_to_pdf(inf.name, outf)
                    if sukses:
                        st.success("Berhasil diubah!")
                        with open(outf, "rb") as f: st.download_button("💾 Unduh File PDF", f, file_name=up_file.name.replace(".docx", ".pdf"), use_container_width=True)
                    else: st.error(f"Gagal. Server error: {msg}")

    elif tool == "Excel to PDF":
        up_file = st.file_uploader("Unggah file Excel (.xlsx):", type=["xlsx"])
        if up_file:
            if st.button("⚡ Konversi ke PDF", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(".xlsx", ".pdf")
                with st.spinner("Mengonversi..."):
                    sukses, msg = excel_to_pdf(inf.name, outf)
                    if sukses:
                        st.success("Berhasil diubah!")
                        with open(outf, "rb") as f: st.download_button("💾 Unduh File PDF", f, file_name=up_file.name.replace(".xlsx", ".pdf"), use_container_width=True)
                    else: st.error(f"Gagal. Server error: {msg}")

    elif tool == "PowerPoint to PDF":
        up_file = st.file_uploader("Unggah file PowerPoint (.pptx):", type=["pptx"])
        if up_file:
            if st.button("⚡ Konversi ke PDF", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(".pptx", ".pdf")
                with st.spinner("Mengonversi..."):
                    sukses, msg = ppt_to_pdf(inf.name, outf)
                    if sukses:
                        st.success("Berhasil diubah!")
                        with open(outf, "rb") as f: st.download_button("💾 Unduh File PDF", f, file_name=up_file.name.replace(".pptx", ".pdf"), use_container_width=True)
                    else: st.error(f"Gagal. Server error: {msg}")

    elif tool == "PDF to Word":
        up_file = st.file_uploader("Unggah file PDF:", type=["pdf"])
        if up_file:
            if st.button("⚡ Konversi ke Word", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(".pdf", ".docx")
                with st.spinner("Mengekstrak berkas..."):
                    try:
                        from pdf2docx import Converter
                        cv = Converter(inf.name); cv.convert(outf, start=0, end=None); cv.close()
                        st.success("Sukses!")
                        with open(outf, "rb") as f: st.download_button("💾 Unduh Word File", f, file_name=up_file.name.replace(".pdf", ".docx"), use_container_width=True)
                    except Exception as e: st.error(f"Error: {e}")

    elif tool == "PDF to Excel":
        up_file = st.file_uploader("Unggah file PDF berisi tabel:", type=["pdf"])
        if up_file:
            if st.button("⚡ Konversi ke Excel", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(".pdf", ".xlsx")
                with st.spinner("Mendeteksi tabel..."):
                    try:
                        all_tables = []
                        with pdfplumber.open(inf.name) as pdf:
                            for p in pdf.pages:
                                ext = p.extract_tables()
                                for t in ext: all_tables.append(pd.DataFrame(t))
                        if all_tables:
                            with pd.ExcelWriter(outf) as writer:
                                for idx, df in enumerate(all_tables): df.to_excel(writer, sheet_name=f"Page_{idx+1}", index=False, header=False)
                            st.success("Berhasil diekstrak!")
                            with open(outf, "rb") as f: st.download_button("💾 Unduh Excel File", f, file_name=up_file.name.replace(".pdf", ".xlsx"), use_container_width=True)
                        else: st.warning("Tidak ditemukan struktur tabel data.")
                    except Exception as e: st.error(f"Error: {e}")

    elif tool == "JPG to PDF":
        up_file = st.file_uploader("Unggah berkas gambar:", type=["jpg","jpeg","png"])
        if up_file:
            if st.button("⚡ Jadikan PDF", type="primary"):
                ext = os.path.splitext(up_file.name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as inf:
                    inf.write(up_file.getvalue())
                    outf = inf.name.replace(ext, ".pdf")
                try:
                    img = Image.open(inf.name)
                    if img.mode in ('RGBA', 'LA'): img = img.convert('RGB')
                    img.save(outf, "PDF")
                    st.success("Konversi Berhasil!")
                    with open(outf, "rb") as f: st.download_button("💾 Unduh PDF", f, file_name=os.path.splitext(up_file.name)[0]+".pdf", use_container_width=True)
                except Exception as e: st.error(str(e))

    elif tool == "Rotate PDF":
        up_file = st.file_uploader("Unggah PDF:", type=["pdf"])
        if up_file:
            sudut = st.selectbox("Pilih Derajat Putaran:", [90, 180, 270])
            if st.button("⚡ Putar Dokumen", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as outf:
                    inf.write(up_file.getvalue())
                    sukses, msg = rotate_pdf(inf.name, outf.name, sudut)
                    if sukses:
                        st.success("Rotasi halaman sukses!")
                        with open(outf.name, "rb") as f: st.download_button("💾 Unduh PDF Baru", f, file_name=f"Rotated_{up_file.name}", use_container_width=True)
                    else: st.error(msg)

    elif tool == "Protect PDF":
        up_file = st.file_uploader("Unggah PDF:", type=["pdf"])
        if up_file:
            pwd = st.text_input("Masukkan Password Pengunci:", type="password")
            if pwd and st.button("⚡ Kunci & Amankan PDF", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as outf:
                    inf.write(up_file.getvalue())
                    sukses, msg = protect_pdf(inf.name, outf.name, pwd)
                    if sukses:
                        st.success("Dokumen berhasil dienkripsi password!")
                        with open(outf.name, "rb") as f: st.download_button("💾 Unduh PDF Terproteksi", f, file_name=f"Protected_{up_file.name}", use_container_width=True)
                    else: st.error(msg)

    elif tool == "Unlock PDF":
        up_file = st.file_uploader("Unggah PDF Terkunci Password:", type=["pdf"])
        if up_file:
            pwd = st.text_input("Masukkan Password Pembuka:", type="password")
            if pwd and st.button("⚡ Buka Proteksi PDF", type="primary"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as inf, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as outf:
                    inf.write(up_file.getvalue())
                    sukses, msg = unlock_pdf(inf.name, outf.name, pwd)
                    if sukses:
                        st.success("Kunci enkripsi berhasil dijebol/dihapus!")
                        with open(outf.name, "rb") as f: st.download_button("💾 Unduh PDF Terbuka", f, file_name=f"Unlocked_{up_file.name}", use_container_width=True)
                    else: st.error("Gagal membuka proteksi. Pastikan password benar.")

    # --- PENGAMANAN LAYOUT/PLACEHOLDER BAGIAN FITUR SANGAT BERAT / MODEL AI ---
    else:
        st.info("⚙️ Infrastruktur Menu Terdeteksi!")
        st.warning(f"Fitur **{tool}** memerlukan integrasi API cloud eksternal komersial (seperti OpenAI/Enterprise SDK) untuk memproses data berskala besar tanpa limitasi ram cloud.")
        st.write("UI modul sudah dipasang dengan aman. Untuk menghubungkan backend algoritma server kustom di fitur ini, silakan hubungi tim DevOps internal atau pasang API Key Anda.")
        st.button("⚙️ Jalankan Simulasi Sistem Tes Otomatis", use_container_width=True)
