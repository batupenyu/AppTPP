"""
dashboard.py
------------
Dashboard 1 pintu untuk pipeline run_all.py, ditata dalam beberapa tab
(navbar) supaya tiap halaman singkat dan tidak perlu scroll panjang:

  📤 Upload PDF         -> tambah/ganti file PDF ke folder tujuan
  ⚙️ Jalankan Pipeline  -> checklist proses + tombol Generate + log
  📝 Surat Pernyataan   -> generate SURAT_PERNYATAAN_PNS.docx / P3K.docx
  📁 File Output        -> ringkasan semua file hasil + download

File lama HANYA dihapus setelah file baru berhasil dibuat (pola
backup -> generate -> commit/restore), jadi kalau proses gagal di
tengah jalan, file lama tidak ikut hilang.
"""

import shutil
import subprocess
import sys
import time
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="Dashboard Pipeline TPP",
    page_icon="📊",
    layout="wide",
)

BASE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Peta step -> script -> (pdf input, output file yang dihasilkan)
# ---------------------------------------------------------------------------
STEPS = {
    "gaji.py": {
        "label": "Gaji (gaji.pdf -> gaji.csv)",
        "pdf": BASE / "gaji.pdf",
        "outputs": [BASE / "gaji.csv"],
    },
    "lem.py": {
        "label": "Lembar 3 (lembar3.pdf -> lembar3.csv)",
        "pdf": BASE / "lembar3.pdf",
        "outputs": [BASE / "lembar3.csv"],
    },
    "per.py": {
        "label": "Perhitungan (perhitungan.pdf -> perhitungan.csv + ringkasan)",
        "pdf": BASE / "perhitungan.pdf",
        "outputs": [BASE / "perhitungan.csv", BASE / "perhitungan_ringkasan.csv"],
    },
}
MERGE_OUTPUT = BASE / "gabung.xlsx"

FOLDER_CHOICES = {
    "Folder utama (root)": BASE,
    "_usulan_tpp_smkn1_koba/PNS": BASE / "_usulan_tpp_smkn1_koba" / "PNS",
    "_usulan_tpp_smkn1_koba/PPPK": BASE / "_usulan_tpp_smkn1_koba" / "PPPK",
}

BACKUP_SUFFIX = ".bak_dashboard"

# Sumber Excel & output docx untuk fitur "Generate Surat Pernyataan (SPTJM)".
# generate_surat_pernyataan.py membaca sheet 'sptjm' dari xlsm ini; jenis TPP
# (Negeri Sipil / PPPK) otomatis mengikuti isi Excel, bukan ditebak di dashboard.
SURAT_TARGETS = {
    "PNS": {
        "label": "SURAT_PERNYATAAN_PNS.docx",
        "source": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "TPP PNS JULI 2026 SMKN 1 KOBA.xlsm",
        "output": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.docx",
        "type": "PNS",
        "html_label": "SURAT_PERNYATAAN_PNS.html",
        "html_output": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.html",
    },
    "P3K": {
        "label": "SURAT_PERNYATAAN_P3K.docx",
        "source": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "TPP P3K JULI 2026 SMKN 1 KOBA.xlsm",
        "output": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "SURAT_PERNYATAAN_P3K.docx",
        "type": "PPPK",
        "html_label": "SURAT_PERNYATAAN_P3K.html",
        "html_output": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "SURAT_PERNYATAAN_P3K.html",
    },
}


# ---------------------------------------------------------------------------
# Helper: pola aman "backup lama -> tulis baru -> commit (hapus backup)"
# Kalau gagal, backup dikembalikan sehingga file lama tidak hilang.
# ---------------------------------------------------------------------------
def backup_files(paths):
    backups = {}
    for p in paths:
        if p.exists():
            bak = p.with_name(p.name + BACKUP_SUFFIX)
            shutil.move(str(p), str(bak))
            backups[p] = bak
    return backups


def commit_backups(backups):
    """Generate baru sukses -> baru sekarang file lama betul-betul dihapus."""
    removed = []
    for original, bak in backups.items():
        if bak.exists():
            bak.unlink()
            removed.append(original.name)
    return removed


def restore_backups(backups):
    """Generate gagal -> kembalikan file lama, jangan sampai hilang."""
    for original, bak in backups.items():
        if bak.exists():
            shutil.move(str(bak), str(original))


def save_upload_to(target_path: Path, uploaded_file):
    """Simpan file upload ke target_path dengan pola backup/commit di atas."""
    backup = backup_files([target_path]) if target_path.exists() else {}
    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        with open(target_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        commit_backups(backup)
        return True, None
    except Exception as e:
        restore_backups(backup)
        return False, str(e)


def run_script(script_path: Path, args):
    cmd = [sys.executable, str(script_path)] + [str(a) for a in args]
    return subprocess.run(cmd, cwd=str(BASE), capture_output=True, text=True)


def render_surat_format(jenis, info, fmt, label, output_path: Path):
    """Render tombol Generate + Download untuk satu format (docx/html) pada satu jenis surat."""
    icon = "🌐" if fmt == "html" else "📄"
    if output_path.exists():
        size_kb = output_path.stat().st_size / 1024
        st.caption(f"{icon} {label} sudah ada ({size_kb:.1f} KB)")

    if st.button(
        f"{icon} Generate {label}",
        key=f"gen_surat_{jenis}_{fmt}",
        disabled=not info["source"].exists(),
        use_container_width=True,
    ):
        backups = backup_files([output_path])
        result = run_script(
            BASE / "generate_surat_pernyataan.py",
            [str(info["source"]), "-t", info["type"], "-f", fmt, "-o", str(output_path)],
        )
        if result.returncode == 0 and output_path.exists():
            commit_backups(backups)
            st.success(f"✅ {label} berhasil dibuat.")
            if result.stdout:
                st.code(result.stdout, language=None)
        else:
            restore_backups(backups)
            st.error(f"❌ Gagal membuat {label} (file lama dipertahankan).")
            if result.stdout:
                st.code(result.stdout, language=None)
            if result.stderr:
                st.code(result.stderr, language=None)
        st.rerun()

    if output_path.exists():
        with open(output_path, "rb") as f:
            st.download_button(
                f"⬇️ Download {label}",
                data=f,
                file_name=label,
                key=f"dl_surat_{jenis}_{fmt}",
                use_container_width=True,
            )


# ---------------------------------------------------------------------------
# UI - navbar (tab) supaya tiap halaman singkat
# ---------------------------------------------------------------------------
st.title("📊 Dashboard Pipeline TPP (run_all.py)")

tab_upload, tab_pipeline, tab_surat, tab_output = st.tabs(
    ["📤 Upload PDF", "⚙️ Jalankan Pipeline", "📝 Surat Pernyataan", "📁 File Output"]
)

# === TAB 1: UPLOAD / GANTI FILE PDF ========================================
with tab_upload:
    st.header("Tambah / Ganti File PDF")
    st.caption(
        "File lama otomatis diganti, tapi baru benar-benar dihapus "
        "**setelah** file baru berhasil disimpan."
    )

    with st.form("form_upload_pdf", clear_on_submit=True):
        col_a, col_b = st.columns([1, 1])
        with col_a:
            folder_label = st.selectbox("Folder tujuan", list(FOLDER_CHOICES.keys()))
        with col_b:
            custom_name = st.text_input(
                "Nama file di folder tujuan (kosongkan = pakai nama file upload)",
                placeholder="mis. gaji.pdf",
            )
        uploaded_pdf = st.file_uploader("Pilih file PDF", type=["pdf"])
        submitted = st.form_submit_button("⬆️ Simpan ke folder", type="primary")

        if submitted:
            if uploaded_pdf is None:
                st.error("❌ Pilih file PDF terlebih dahulu.")
            else:
                target_dir = FOLDER_CHOICES[folder_label]
                filename = custom_name.strip() or uploaded_pdf.name
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"
                target_path = target_dir / filename

                existed_before = target_path.exists()
                ok, err = save_upload_to(target_path, uploaded_pdf)
                if ok:
                    if existed_before:
                        st.success(f"✅ File lama '{filename}' diganti dengan yang baru di {folder_label}.")
                    else:
                        st.success(f"✅ File '{filename}' ditambahkan ke {folder_label}.")
                    st.rerun()
                else:
                    st.error(f"❌ Gagal menyimpan file: {err} (file lama tetap dipertahankan)")

    with st.expander("📂 Lihat isi folder tujuan"):
        view_folder_label = st.selectbox(
            "Pilih folder untuk dilihat", list(FOLDER_CHOICES.keys()), key="view_folder"
        )
        view_dir = FOLDER_CHOICES[view_folder_label]
        files = sorted(view_dir.glob("*")) if view_dir.exists() else []
        files = [f for f in files if f.is_file() and not f.name.endswith(BACKUP_SUFFIX)]
        if files:
            for f in files:
                size_kb = f.stat().st_size / 1024
                c1, c2 = st.columns([4, 1])
                c1.caption(f"📄 {f.name} ({size_kb:.1f} KB)")
                if c2.button("🗑️ Hapus", key=f"del_{f}"):
                    f.unlink()
                    st.rerun()
        else:
            st.caption("Folder kosong.")

# === TAB 2: CHECKLIST + GENERATE PIPELINE ==================================
with tab_pipeline:
    st.header("Pilih Proses yang Akan Dijalankan")

    selected_steps = {}
    cols = st.columns(len(STEPS))
    for col, (script_name, info) in zip(cols, STEPS.items()):
        with col:
            pdf_exists = info["pdf"].exists()
            st.caption(f"{'✅' if pdf_exists else '⚪'} {info['pdf'].name}")
            selected_steps[script_name] = st.checkbox(
                info["label"],
                value=pdf_exists,
                disabled=not pdf_exists,
                key=f"chk_{script_name}",
            )

    run_merge = st.checkbox(
        "🔄 Gabungkan hasil (gabung.py -> gabung.xlsx)", value=True, key="chk_merge"
    )
    st.caption(
        "🧹 File CSV/XLSX lama otomatis dihapus, tapi **hanya setelah** file baru "
        "untuk proses yang sama berhasil dibuat. Kalau proses gagal, file lama tidak berubah."
    )

    st.markdown("---")
    generate_clicked = st.button("🚀 Generate", type="primary", use_container_width=True)

    if generate_clicked:
        steps_to_run = [name for name, checked in selected_steps.items() if checked]

        if not steps_to_run:
            st.error("❌ Pilih minimal satu proses untuk dijalankan.")
        else:
            st.markdown("---")
            st.subheader("Proses & Log")
            log_area = st.container()
            progress = st.progress(0.0)
            total = len(steps_to_run) + (1 if run_merge else 0)
            done = 0
            pipeline_ok = True

            for script_name in steps_to_run:
                info = STEPS[script_name]
                with log_area:
                    st.write(f"**▶ Menjalankan {script_name}...**")

                backups = backup_files(info["outputs"])
                result = run_script(info["pdf"].parent / script_name, [str(info["pdf"])])

                with log_area:
                    if result.stdout:
                        st.code(result.stdout, language=None)

                new_outputs_ok = result.returncode == 0 and all(
                    p.exists() for p in info["outputs"]
                )

                if new_outputs_ok:
                    removed = commit_backups(backups)
                    with log_area:
                        st.success(f"✅ {script_name} selesai.")
                        if removed:
                            st.caption(f"🗑️ File lama dihapus: {', '.join(removed)}")
                else:
                    restore_backups(backups)
                    pipeline_ok = False
                    with log_area:
                        st.error(f"❌ {script_name} gagal. File lama dipertahankan.")
                        if result.stderr:
                            st.code(result.stderr, language=None)
                    break

                done += 1
                progress.progress(done / total)

            if pipeline_ok and run_merge:
                with log_area:
                    st.write("**▶ Menjalankan gabung.py (merge)...**")

                # gabung.py sudah menulis via file sementara + os.replace,
                # jadi gabung.xlsx lama otomatis hanya tergantikan setelah
                # file baru selesai ditulis dengan sempurna.
                result = run_script(BASE / "gabung.py", [])

                with log_area:
                    if result.stdout:
                        st.code(result.stdout, language=None)
                    if result.returncode == 0 and MERGE_OUTPUT.exists():
                        st.success("✅ gabung.py selesai. gabung.xlsx diperbarui.")
                    else:
                        pipeline_ok = False
                        st.error("❌ gabung.py gagal.")
                        if result.stderr:
                            st.code(result.stderr, language=None)

                done += 1
                progress.progress(done / total)

            progress.progress(1.0)
            if pipeline_ok:
                st.balloons()
                st.success("🎉 Pipeline selesai! Cek hasilnya di tab 📁 File Output.")
            else:
                st.error("⚠️ Pipeline berhenti karena ada error di atas.")

# === TAB 3: GENERATE SURAT PERNYATAAN (SPTJM) — PNS & P3K ==================
with tab_surat:
    st.header("Generate Surat Pernyataan (SPTJM)")
    st.caption(
        "Dibuat dari sheet 'sptjm' pada file Excel TPP masing-masing. "
        "Jenis TPP (Negeri Sipil / PPPK) otomatis mengikuti isi Excel."
    )

    surat_cols = st.columns(2)
    for col, (jenis, info) in zip(surat_cols, SURAT_TARGETS.items()):
        with col:
            st.subheader(f"{'👤' if jenis == 'PNS' else '🧑‍💼'} {jenis}")
            source_ok = info["source"].exists()
            st.caption(f"{'✅' if source_ok else '⚪'} Sumber: {info['source'].name}")

            render_surat_format(jenis, info, "docx", info["label"], info["output"])
            st.markdown("---")
            render_surat_format(jenis, info, "html", info["html_label"], info["html_output"])

# === TAB 4: RINGKASAN FILE OUTPUT ==========================================
with tab_output:
    st.header("File Output")

    st.subheader("Hasil Ekstraksi & Merge")
    output_files = {
        "gaji.csv": BASE / "gaji.csv",
        "lembar3.csv": BASE / "lembar3.csv",
        "perhitungan.csv": BASE / "perhitungan.csv",
        "perhitungan_ringkasan.csv": BASE / "perhitungan_ringkasan.csv",
        "gabung.xlsx": MERGE_OUTPUT,
    }
    file_cols = st.columns(len(output_files))
    for col, (name, path) in zip(file_cols, output_files.items()):
        with col:
            if path.exists():
                size_kb = path.stat().st_size / 1024
                mtime = time.strftime("%H:%M:%S", time.localtime(path.stat().st_mtime))
                st.metric(label=name, value=f"{size_kb:.1f} KB")
                st.caption(f"🕒 Diperbarui {mtime}")
                with open(path, "rb") as f:
                    st.download_button(
                        "⬇️ Download", data=f, file_name=name, key=f"dl_{name}"
                    )
            else:
                st.caption(f"⚪ {name}")
                st.caption("Belum dibuat")

    st.markdown("---")
    st.subheader("Surat Pernyataan")
    surat_cols2 = st.columns(2)
    for col, (jenis, info) in zip(surat_cols2, SURAT_TARGETS.items()):
        with col:
            for fmt, label_key, out_key in [
                ("docx", "label", "output"),
                ("html", "html_label", "html_output"),
            ]:
                path = info[out_key]
                label = info[label_key]
                if path.exists():
                    size_kb = path.stat().st_size / 1024
                    mtime = time.strftime("%H:%M:%S", time.localtime(path.stat().st_mtime))
                    st.metric(label=label, value=f"{size_kb:.1f} KB")
                    st.caption(f"🕒 Diperbarui {mtime}")
                    with open(path, "rb") as f:
                        st.download_button(
                            "⬇️ Download", data=f, file_name=label, key=f"dl_out_{jenis}_{fmt}"
                        )
                else:
                    st.caption(f"⚪ {label}")
                    st.caption("Belum dibuat")