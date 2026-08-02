import streamlit as st
import subprocess
import sys
import os
from pathlib import Path

st.set_page_config(
    page_title="PDF to CSV Pipeline Dashboard",
    page_icon="📊",
    layout="wide"
)

BASE = Path(__file__).parent

st.title("📊 Dashboard Proses Ekstraksi PDF ke CSV")

st.markdown("---")

with st.sidebar:
    st.header("⚙️ Konfigurasi")
    
    st.subheader("📤 Ganti File PDF")
    pdf_replacements = {
        "gaji.pdf": BASE / "gaji.pdf",
        "lembar3.pdf": BASE / "lembar3.pdf",
        "perhitungan.pdf": BASE / "perhitungan.pdf",
    }
    
    for pdf_name, pdf_path in pdf_replacements.items():
        exists = pdf_path.exists()
        if exists:
            size = pdf_path.stat().st_size / 1024
            st.caption(f"📄 {pdf_name} ({size:.1f} KB)")
        else:
            st.caption(f"⚪ {pdf_name} - belum ada")
        
        uploaded = st.file_uploader(
            f"Ganti {pdf_name}",
            type=["pdf"],
            key=f"upload_{pdf_name}",
            label_visibility="collapsed"
        )
        if uploaded is not None:
            with open(pdf_path, "wb") as f:
                f.write(uploaded.getbuffer())
            st.success(f"✅ {pdf_name} berhasil diganti!")
            st.rerun()
    
    st.markdown("---")
    st.subheader("📝 Generate Surat Pernyataan")
    tpp_excel = BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "TPP PNS JULI 2026 SMKN 1 KOBA.xlsm"
    surat_output = BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.docx"
    
    if tpp_excel.exists():
        st.caption(f"📊 Sumber: {tpp_excel.name}")
    else:
        st.caption(f"⚪ File Excel TPP tidak ditemukan")
    
    if st.button("📄 Generate SURAT_PERNYATAAN_PNS.docx", use_container_width=True):
        if not tpp_excel.exists():
            st.error("❌ File Excel TPP tidak ditemukan!")
        else:
            with st.spinner("Sedang membuat surat pernyataan..."):
                try:
                    result = subprocess.run(
                        [sys.executable, str(BASE / "generate_surat_pernyataan.py"), str(tpp_excel), "-o", str(surat_output)],
                        cwd=str(BASE),
                        capture_output=True,
                        text=True
                    )
                    if result.returncode == 0:
                        st.success("✅ Surat pernyataan berhasil dibuat!")
                        if result.stdout:
                            st.code(result.stdout, language=None)
                        if surat_output.exists():
                            with open(surat_output, "rb") as f:
                                st.download_button(
                                    label="⬇️ Download SURAT_PERNYATAAN_PNS.docx",
                                    data=f,
                                    file_name="SURAT_PERNYATAAN_PNS.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key="dl_surat"
                                )
                        st.rerun()
                    else:
                        st.error(f"❌ Gagal membuat surat pernyataan")
                        if result.stderr:
                            st.error(result.stderr)
                        if result.stdout:
                            st.code(result.stdout, language=None)
                except Exception as e:
                    st.error(f"❌ Error: {e}")
    
    if surat_output.exists():
        size = surat_output.stat().st_size / 1024
        st.caption(f"📄 SURAT_PERNYATAAN_PNS.docx ({size:.1f} KB)")
    
    st.markdown("---")
    st.subheader("Pilih File PDF untuk Ekstraksi")
    pdf_files = {
        "gaji.py": BASE / "gaji.pdf",
        "lem.py": BASE / "lembar3.pdf",
        "per.py": BASE / "perhitungan.pdf",
    }
    
    selected_scripts = {}
    for script_name, pdf_path in pdf_files.items():
        exists = pdf_path.exists()
        label = f"{'✅' if exists else '❌'} {pdf_path.name}"
        selected_scripts[script_name] = st.checkbox(
            label, 
            value=exists,
            disabled=not exists,
            help=f"Script: {script_name}, PDF: {pdf_path.name}"
        )
    
    run_merge = st.checkbox("🔄 Jalankan gabung.py (Merge)", value=True, help="Gabungkan semua CSV menjadi gabung.xlsx")
    clean_first = st.checkbox("🧹 Bersihkan file lama", value=True, help="Hapus file output lama sebelum menjalankan")
    
    st.markdown("---")
    st.subheader("📋 Output Files")
    output_files = {
        "gaji.csv": BASE / "gaji.csv",
        "lembar3.csv": BASE / "lembar3.csv",
        "perhitungan.csv": BASE / "perhitungan.csv",
        "perhitungan_ringkasan.csv": BASE / "perhitungan_ringkasan.csv",
        "gabung.xlsx": BASE / "gabung.xlsx",
    }
    
    for name, path in output_files.items():
        if path.exists():
            size = path.stat().st_size / 1024
            st.caption(f"📄 {name} ({size:.1f} KB)")
        else:
            st.caption(f"⚪ {name}")

st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    run_button = st.button("▶️ Jalankan Pipeline", type="primary", use_container_width=True)
with col2:
    run_selected_button = st.button("🔍 Jalankan yang Dipilih", use_container_width=True)
with col3:
    refresh_button = st.button("🔄 Refresh", use_container_width=True)

if refresh_button:
    st.rerun()

if run_button or run_selected_button:
    if not any(selected_scripts.values()):
        st.error("❌ Pilih minimal satu file PDF untuk diekstraksi!")
    else:
        log_container = st.container()
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        logs = []
        
        def add_log(message, level="info"):
            logs.append((level, message))
            with log_container:
                if level == "error":
                    st.error(message)
                elif level == "success":
                    st.success(message)
                elif level == "warning":
                    st.warning(message)
                else:
                    st.info(message)
        
        steps_to_run = []
        if run_button:
            steps_to_run = list(selected_scripts.keys())
        else:
            steps_to_run = [k for k, v in selected_scripts.items() if v]
        
        total_steps = len(steps_to_run) + (1 if run_merge else 0)
        current_step = 0
        
        add_log("🚀 Memulai proses ekstraksi PDF...", "info")
        
        if clean_first:
            add_log("🧹 Membersihkan file lama...", "info")
            output_paths = [
                BASE / "gaji.csv",
                BASE / "lembar3.csv",
                BASE / "perhitungan.csv",
                BASE / "perhitungan_ringkasan.csv",
                BASE / "gabung.xlsx",
            ]
            for f in output_paths:
                if f.exists():
                    f.unlink()
                    add_log(f"   🗑️ Removed: {f.name}", "info")
            add_log("✅ File lama berhasil dibersihkan", "success")
        
        success = True
        for script_name in steps_to_run:
            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text(f"⏳ Menjalankan: {script_name}...")
            
            script_path = BASE / script_name
            if script_name == "gaji.py":
                pdf_arg = str(BASE / "gaji.pdf")
            elif script_name == "lem.py":
                pdf_arg = str(BASE / "lembar3.pdf")
            elif script_name == "per.py":
                pdf_arg = str(BASE / "perhitungan.pdf")
            else:
                pdf_arg = ""
            
            add_log(f"\n{'='*60}", "info")
            add_log(f"🔄 Running: {script_name}", "info")
            add_log(f"{'='*60}", "info")
            
            cmd = [sys.executable, str(script_path), pdf_arg]
            result = subprocess.run(cmd, cwd=str(BASE), capture_output=True, text=True)
            
            if result.stdout:
                for line in result.stdout.strip().split('\n'):
                    add_log(f"   {line}", "info")
            
            if result.returncode != 0:
                add_log(f"❌ ERROR: {script_name} failed with exit code {result.returncode}", "error")
                if result.stderr:
                    add_log(f"   {result.stderr}", "error")
                success = False
                break
            else:
                add_log(f"✅ {script_name} completed successfully", "success")
        
        if success and run_merge:
            current_step += 1
            progress_bar.progress(current_step / total_steps)
            status_text.text("⏳ Menjalankan gabung.py...")
            
            add_log(f"\n{'='*60}", "info")
            add_log("🔄 Running: gabung.py (merge)", "info")
            add_log(f"{'='*60}", "info")
            
            merge_script = BASE / "gabung.py"
            result = subprocess.run([sys.executable, str(merge_script)], cwd=str(BASE), capture_output=True, text=True)
            
            if result.stdout:
                for line in result.stdout.strip().split('\n'):
                    add_log(f"   {line}", "info")
            
            if result.returncode != 0:
                add_log(f"❌ ERROR: gabung.py failed with exit code {result.returncode}", "error")
                if result.stderr:
                    add_log(f"   {result.stderr}", "error")
                success = False
            else:
                add_log("✅ gabung.py completed successfully", "success")
        
        progress_bar.progress(1.0)
        
        if success:
            status_text.text("✅ Pipeline selesai!")
            st.balloons()
            add_log("\n" + "="*60, "success")
            add_log("🎉 Pipeline finished successfully!", "success")
            if (BASE / "gabung.xlsx").exists():
                add_log(f"📊 Output: gabung.xlsx", "success")
            add_log("="*60, "success")
        else:
            status_text.text("❌ Pipeline stopped with errors")
            add_log("\n❌ Pipeline stopped with errors.", "error")

st.markdown("---")
st.subheader("📁 File Output")

output_files = {
    "gaji.csv": BASE / "gaji.csv",
    "lembar3.csv": BASE / "lembar3.csv",
    "perhitungan.csv": BASE / "perhitungan.csv",
    "perhitungan_ringkasan.csv": BASE / "perhitungan_ringkasan.csv",
    "gabung.xlsx": BASE / "gabung.xlsx",
    "SURAT_PERNYATAAN_PNS.docx": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.docx",
}

file_cols = st.columns(5)
for i, (name, path) in enumerate(output_files.items()):
    with file_cols[i]:
        if path.exists():
            size = path.stat().st_size / 1024
            st.metric(label=name, value=f"{size:.1f} KB", delta="✅")
            with open(path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Download",
                    data=f,
                    file_name=name,
                    mime="application/octet-stream",
                    key=f"dl_{name}"
                )
        else:
            st.metric(label=name, value="0 KB", delta="❌ Belum ada")

st.markdown("---")
st.subheader("📝 Log Interaktif")

if 'logs' not in st.session_state:
    st.session_state.logs = []

if st.button("🗑️ Clear Log"):
    st.session_state.logs = []
    st.rerun()

log_placeholder = st.empty()
with log_placeholder.container():
    if st.session_state.logs:
        for level, msg in st.session_state.logs[-50:]:
            if level == "error":
                st.error(msg)
            elif level == "success":
                st.success(msg)
            elif level == "warning":
                st.warning(msg)
            else:
                st.text(msg)
    else:
        st.caption("Belum ada log. Jalankan pipeline untuk melihat log.")
