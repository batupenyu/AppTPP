"""
gui.py
------
Desktop GUI (CustomTkinter) untuk pipeline PythonToCsv_V.4.

Tab:
  1. Upload PDF     - tambah / ganti file PDF
  2. Jalankan       - checklist proses + Generate + log
   3. Surat          - generate SURAT_PERNYATAAN_PNS / P3K (html) + refresh rekap_tpp
  4. File Output    - ringkasan hasil + download + buka folder
"""

import os
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

BASE = Path(__file__).parent

# ---------------------------------------------------------------------------
# Konfigurasi pipeline
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

SURAT_TARGETS = {
    "PNS": {
        "label": "SURAT_PERNYATAAN_PNS.html",
        "source": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "TPP PNS.xlsm",
        "output": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.html",
        "type": "PNS",
        "html_label": "SURAT_PERNYATAAN_PNS.html",
        "html_output": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "SURAT_PERNYATAAN_PNS.html",
    },
    "P3K": {
        "label": "SURAT_PERNYATAAN_P3K.html",
        "source": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "TPP P3K.xlsm",
        "output": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "SURAT_PERNYATAAN_P3K.html",
        "type": "PPPK",
        "html_label": "SURAT_PERNYATAAN_P3K.html",
        "html_output": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "SURAT_PERNYATAAN_P3K.html",
    },
}

USULAN_DIR = BASE / "_usulan_tpp_smkn1_koba"

TPP_FILES = {
    "PNS": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "TPP PNS.xlsm",
    "PPPK": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "TPP P3K.xlsm",
}

OUTPUT_FILES = {
    "gaji.csv": BASE / "gaji.csv",
    "lembar3.csv": BASE / "lembar3.csv",
    "perhitungan.csv": BASE / "perhitungan.csv",
    "perhitungan_ringkasan.csv": BASE / "perhitungan_ringkasan.csv",
    "gabung.xlsx": MERGE_OUTPUT,
}


# ---------------------------------------------------------------------------
# Helper functions (backup / restore / commit)
# ---------------------------------------------------------------------------


def backup_files(paths):
    backups = {}
    for p in paths:
        if p.exists():
            bak = p.with_name(p.name + BACKUP_SUFFIX)
            shutil.copy2(str(p), str(bak))   # copy, bukan move
            backups[p] = bak
    return backups


def commit_backups(backups):
    removed = []
    for original, bak in backups.items():
        if bak.exists():
            bak.unlink()
            removed.append(original.name)
    return removed


def restore_backups(backups):
    for original, bak in backups.items():
        if bak.exists():
            shutil.move(str(bak), str(original))


def run_script(script_path, args):
    cmd = [sys.executable, str(script_path)] + [str(a) for a in args]
    result = subprocess.run(cmd, cwd=str(BASE), capture_output=True, text=True)
    return result


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Dashboard Pipeline TPP")
        self.geometry("1100x750")

        # grid layout 1x1
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tab_view = ctk.CTkTabview(self, width=1100, height=750)
        self.tab_view.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.tab_upload = self.tab_view.add("📤 Upload PDF")
        self.tab_pipeline = self.tab_view.add("⚙️ Jalankan Pipeline")
        self.tab_surat = self.tab_view.add("📝 Surat Pernyataan")
        self.tab_output = self.tab_view.add("📁 File Output")

        self._build_upload_tab()
        self._build_pipeline_tab()
        self._build_surat_tab()
        self._build_output_tab()

    # =====================================================================
    # TAB 1: UPLOAD PDF
    # =====================================================================
    def _build_upload_tab(self):
        tab = self.tab_upload

        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        header = ctk.CTkLabel(tab, text="Tambah / Ganti File PDF", font=ctk.CTkFont(size=18, weight="bold"))
        header.grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))

        form = ctk.CTkFrame(tab)
        form.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(form, text="Folder tujuan:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.upload_folder_var = ctk.StringVar(value=list(FOLDER_CHOICES.keys())[0])
        self.upload_folder_cb = ctk.CTkOptionMenu(form, variable=self.upload_folder_var, values=list(FOLDER_CHOICES.keys()))
        self.upload_folder_cb.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        ctk.CTkLabel(form, text="Nama file (kosong = pakai nama upload):").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.upload_name_entry = ctk.CTkEntry(form, placeholder_text="mis. gaji.pdf")
        self.upload_name_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        self.upload_path_label = ctk.CTkLabel(form, text="Belum pilih file", text_color="gray")
        self.upload_path_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=5)

        btn_frame = ctk.CTkFrame(form, fg_color="transparent")
        btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=10)

        self.uploaded_file_path = None

        def choose_file():
            path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if path:
                self.uploaded_file_path = path
                self.upload_path_label.configure(text=path, text_color="white")

        ctk.CTkButton(btn_frame, text="Pilih File PDF", command=choose_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="⬆️ Simpan ke Folder", command=self._do_upload).pack(side="left")

        # Daftar file di folder tujuan
        list_frame = ctk.CTkFrame(tab)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 20))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(list_frame, text="📂 Isi Folder Tujuan", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.file_listbox = ctk.CTkTextbox(list_frame, wrap="none")
        self.file_listbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self._refresh_file_list()

    def _refresh_file_list(self):
        folder_label = self.upload_folder_var.get()
        folder = FOLDER_CHOICES[folder_label]
        self.file_listbox.delete("0.0", "end")
        if folder.exists():
            files = sorted([f for f in folder.iterdir() if f.is_file() and not f.name.endswith(BACKUP_SUFFIX)])
            if files:
                for f in files:
                    size_kb = f.stat().st_size / 1024
                    self.file_listbox.insert("end", f"{f.name}  ({size_kb:.1f} KB)\n")
            else:
                self.file_listbox.insert("end", "Folder kosong.")
        else:
            self.file_listbox.insert("end", "Folder tidak ditemukan.")

    def _do_upload(self):
        if not self.uploaded_file_path:
            messagebox.showerror("Error", "Pilih file PDF terlebih dahulu.")
            return

        folder_label = self.upload_folder_var.get()
        target_dir = FOLDER_CHOICES[folder_label]
        filename = self.upload_name_entry.get().strip() or Path(self.uploaded_file_path).name
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        target_path = target_dir / filename

        existed_before = target_path.exists()
        backup = backup_files([target_path]) if target_path.exists() else {}
        try:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(self.uploaded_file_path, target_path)
            commit_backups(backup)
            messagebox.showinfo("Sukses", f"File '{filename}' berhasil disimpan.")
            self.uploaded_file_path = None
            self.upload_path_label.configure(text="Belum pilih file", text_color="gray")
            self.upload_name_entry.delete(0, "end")
            self._refresh_file_list()
        except Exception as e:
            restore_backups(backup)
            messagebox.showerror("Gagal", f"Gagal menyimpan file: {e}")

    # =====================================================================
    # TAB 2: PIPELINE
    # =====================================================================
    def _build_pipeline_tab(self):
        tab = self.tab_pipeline

        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(tab, text="Pilih Proses yang Akan Dijalankan", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))

        checks_frame = ctk.CTkFrame(tab)
        checks_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        checks_frame.grid_columnconfigure(tuple(range(len(STEPS))), weight=1)

        self.step_vars = {}
        for idx, (script_name, info) in enumerate(STEPS.items()):
            pdf_exists = info["pdf"].exists()
            var = ctk.StringVar(value="1" if pdf_exists else "0")
            self.step_vars[script_name] = var

            frm = ctk.CTkFrame(checks_frame)
            frm.grid(row=0, column=idx, sticky="nsew", padx=5, pady=5)
            frm.grid_columnconfigure(0, weight=1)

            icon = "✅" if pdf_exists else "⚪"
            ctk.CTkLabel(frm, text=f"{icon} {info['pdf'].name}").grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
            chk = ctk.CTkCheckBox(frm, text=info["label"], variable=var, onvalue="1", offvalue="0")
            if not pdf_exists:
                chk.configure(state="disabled")
            chk.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        merge_frame = ctk.CTkFrame(tab)
        merge_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        self.merge_var = ctk.StringVar(value="1")
        ctk.CTkCheckBox(merge_frame, text="🔄 Gabungkan hasil (gabung.py -> gabung.xlsx)", variable=self.merge_var, onvalue="1", offvalue="0").pack(anchor="w", padx=10, pady=10)

        ctk.CTkButton(tab, text="🚀 Generate", command=self._run_pipeline, height=40).grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 10))

        self.log_text = ctk.CTkTextbox(tab, wrap="word")
        self.log_text.grid(row=4, column=0, sticky="nsew", padx=20, pady=(0, 20))
        tab.grid_rowconfigure(4, weight=1)

    def _append_log(self, text):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def _run_pipeline(self):
        steps_to_run = [name for name, var in self.step_vars.items() if var.get() == "1"]
        run_merge = self.merge_var.get() == "1"

        if not steps_to_run:
            messagebox.showerror("Error", "Pilih minimal satu proses untuk dijalankan.")
            return

        self.log_text.delete("0.0", "end")
        self._append_log("=== Memulai pipeline ===\n")

        def worker():
            total = len(steps_to_run) + (1 if run_merge else 0)
            done = 0
            pipeline_ok = True

            for script_name in steps_to_run:
                info = STEPS[script_name]
                self._append_log(f"\n▶ Menjalankan {script_name}...")

                backups = backup_files(info["outputs"])
                result = run_script(info["pdf"].parent / script_name, [str(info["pdf"])])

                if result.stdout:
                    self._append_log(result.stdout.rstrip())

                new_outputs_ok = result.returncode == 0 and all(p.exists() for p in info["outputs"])

                if new_outputs_ok:
                    removed = commit_backups(backups)
                    self._append_log(f"✅ {script_name} selesai.")
                    if removed:
                        self._append_log(f"🗑️ File lama dihapus: {', '.join(removed)}")
                else:
                    restore_backups(backups)
                    pipeline_ok = False
                    self._append_log(f"❌ {script_name} gagal. File lama dipertahankan.")
                    if result.stderr:
                        self._append_log(result.stderr.rstrip())
                    break

                done += 1

            if pipeline_ok and run_merge:
                self._append_log("\n▶ Menjalankan gabung.py (merge)...")
                result = run_script(BASE / "gabung.py", [])
                if result.stdout:
                    self._append_log(result.stdout.rstrip())
                if result.returncode == 0 and MERGE_OUTPUT.exists():
                    self._append_log("✅ gabung.py selesai. gabung.xlsx diperbarui.")
                else:
                    pipeline_ok = False
                    self._append_log("❌ gabung.py gagal.")
                    if result.stderr:
                        self._append_log(result.stderr.rstrip())

                done += 1

            if pipeline_ok:
                self._append_log("\n🎉 Pipeline selesai!")
                messagebox.showinfo("Selesai", "Pipeline berhasil dijalankan.")
            else:
                self._append_log("\n⚠️ Pipeline berhenti karena ada error.")
                messagebox.showerror("Error", "Pipeline berhenti karena ada error. Cek log.")

        threading.Thread(target=worker, daemon=True).start()

    # =====================================================================
    # TAB 3: SURAT PERNYATAAN
    # =====================================================================
    def _build_surat_tab(self):
        tab = self.tab_surat
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(tab, text="Generate Surat Pernyataan (SPTJM)", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10))

        self.surat_buttons = {}

        for col, (jenis, info) in enumerate(SURAT_TARGETS.items()):
            frame = ctk.CTkFrame(tab)
            frame.grid(row=1, column=col, sticky="nsew", padx=20, pady=(0, 20))
            frame.grid_columnconfigure(0, weight=1)

            icon = "👤" if jenis == "PNS" else "🧑‍💼"
            ctk.CTkLabel(frame, text=f"{icon} {jenis}", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

            source_ok = info["source"].exists()
            status = "✅" if source_ok else "⚪"
            ctk.CTkLabel(frame, text=f"{status} Sumber: {info['source'].name}").grid(row=1, column=0, padx=10, pady=5, sticky="w")

            # html
            html_row = 2
            ctk.CTkLabel(frame, text="Format HTML:").grid(row=html_row, column=0, padx=10, pady=(10, 0), sticky="w")
            gen_html = ctk.CTkButton(frame, text="🌐 Generate", command=lambda j=jenis, i=info: self._gen_surat(j, i, "html"))
            gen_html.grid(row=html_row + 1, column=0, padx=10, pady=5, sticky="ew")
            if not source_ok:
                gen_html.configure(state="disabled")

            # Download buttons (updated after generation)
            dl_frame = ctk.CTkFrame(frame, fg_color="transparent")
            dl_frame.grid(row=html_row + 2, column=0, sticky="ew", padx=10, pady=10)

            def make_dl_cmd(path, name):
                def cmd():
                    if path.exists():
                        save_to = filedialog.asksaveasfilename(initialfile=name, defaultextension="*.*", filetypes=[("All files", "*.*")])
                        if save_to:
                            shutil.copy(path, save_to)
                            messagebox.showinfo("Sukses", f"File disimpan ke:\n{save_to}")
                return cmd

            self.surat_buttons[f"{jenis}_html"] = ctk.CTkButton(dl_frame, text="⬇️ Download HTML", command=make_dl_cmd(info["html_output"], info["html_label"]))
            self.surat_buttons[f"{jenis}_html"].pack(fill="x", pady=2)

        ctk.CTkLabel(tab, text="Refresh Sheet Rekap TPP", font=ctk.CTkFont(size=16, weight="bold")).grid(row=2, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10))
        ctk.CTkLabel(tab, text="Perbarui sheet 'rekap_tpp' pada TPP PNS dan TPP P3K dengan data dari gabung.xlsx (kolom GOLONGAN diabaikan).").grid(row=3, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 10))

        for col, (jenis, path) in enumerate(TPP_FILES.items()):
            frame = ctk.CTkFrame(tab)
            frame.grid(row=4, column=col, sticky="nsew", padx=20, pady=(0, 20))
            frame.grid_columnconfigure(0, weight=1)

            file_ok = path.exists()
            status = "✅" if file_ok else "⚪"
            ctk.CTkLabel(frame, text=f"{status} File: {path.name}").grid(row=0, column=0, padx=10, pady=10, sticky="w")

            def make_refresh_cmd(p=path, j=jenis):
                def cmd():
                    if not p.exists():
                        messagebox.showerror("Gagal", f"File tidak ditemukan: {p}")
                        return
                    backups = backup_files([p])
                    result = run_script(BASE / "refresh_rekap_tpp.py", [j])
                    if result.returncode == 0 and p.exists():
                        commit_backups(backups)
                        messagebox.showinfo("Sukses", f"Rekap TPP {j} berhasil di-refresh.")
                    else:
                        restore_backups(backups)
                        msg = f"Gagal refresh rekap_tpp {j}."
                        if result.stderr:
                            msg += f"\n\nSTDERR:\n{result.stderr}"
                        messagebox.showerror("Gagal", msg)
                return cmd

            ctk.CTkButton(frame, text="🔄 Refresh Rekap TPP", command=make_refresh_cmd()).grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    def _gen_surat(self, jenis, info, fmt):
        output_path = info["html_output"]
        backups = backup_files([output_path])
        result = run_script(BASE / "generate_surat_pernyataan.py", [str(info["source"]), "-t", info["type"], "-o", str(output_path)])
        if result.returncode == 0 and output_path.exists():
            commit_backups(backups)
            messagebox.showinfo("Sukses", f"{output_path.name} berhasil dibuat.")
        else:
            restore_backups(backups)
            msg = f"Gagal membuat {output_path.name}."
            if result.stderr:
                msg += f"\n\nSTDERR:\n{result.stderr}"
            messagebox.showerror("Gagal", msg)

    # =====================================================================
    # TAB 4: FILE OUTPUT
    # =====================================================================
    def _build_output_tab(self):
        tab = self.tab_output
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(tab, text="Ringkasan File Output", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 10))

        self.output_frame = ctk.CTkScrollableFrame(tab)
        self.output_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.output_frame.grid_columnconfigure(1, weight=1)

        self._refresh_output()

        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 20))

        def open_folder():
            if USULAN_DIR.exists():
                os.startfile(str(USULAN_DIR))
            else:
                messagebox.showerror("Error", f"Folder tidak ditemukan: {USULAN_DIR}")

        ctk.CTkButton(btn_frame, text="🖥️ Buka Folder Usulan TPP", command=open_folder).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="🔄 Refresh", command=self._refresh_output).pack(side="left")

    def _refresh_output(self):
        for widget in self.output_frame.winfo_children():
            widget.destroy()

        row = 0

        # Header: Hasil Ekstraksi & Merge
        ctk.CTkLabel(self.output_frame, text="Hasil Ekstraksi & Merge", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 5))
        row += 1

        for name, path in OUTPUT_FILES.items():
            if path.exists():
                size_kb = path.stat().st_size / 1024
                mtime = time.strftime("%H:%M:%S", time.localtime(path.stat().st_mtime))
                ctk.CTkLabel(self.output_frame, text=name).grid(row=row, column=0, sticky="w", padx=5, pady=2)
                ctk.CTkLabel(self.output_frame, text=f"{size_kb:.1f} KB  ({mtime})").grid(row=row, column=1, sticky="w", padx=5, pady=2)
                ctk.CTkButton(self.output_frame, text="⬇️ Download", width=100, command=lambda p=path, n=name: self._download_file(p, n)).grid(row=row, column=2, sticky="e", padx=5, pady=2)
            else:
                ctk.CTkLabel(self.output_frame, text=name).grid(row=row, column=0, sticky="w", padx=5, pady=2)
                ctk.CTkLabel(self.output_frame, text="Belum dibuat", text_color="gray").grid(row=row, column=1, sticky="w", padx=5, pady=2)
                ctk.CTkLabel(self.output_frame, text="").grid(row=row, column=2, sticky="e", padx=5, pady=2)
            row += 1

        # Separator
        ctk.CTkFrame(self.output_frame, height=2, fg_color="gray30").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1

        # Surat Pernyataan
        ctk.CTkLabel(self.output_frame, text="Surat Pernyataan", font=ctk.CTkFont(weight="bold")).grid(row=row, column=0, columnspan=3, sticky="w", padx=5, pady=(10, 5))
        row += 1

        for jenis, info in SURAT_TARGETS.items():
            for fmt, label_key, out_key in [("html", "html_label", "html_output")]:
                path = info[out_key]
                label = info[label_key]
                if path.exists():
                    size_kb = path.stat().st_size / 1024
                    mtime = time.strftime("%H:%M:%S", time.localtime(path.stat().st_mtime))
                    ctk.CTkLabel(self.output_frame, text=label).grid(row=row, column=0, sticky="w", padx=5, pady=2)
                    ctk.CTkLabel(self.output_frame, text=f"{size_kb:.1f} KB  ({mtime})").grid(row=row, column=1, sticky="w", padx=5, pady=2)
                    ctk.CTkButton(self.output_frame, text="⬇️ Download", width=100, command=lambda p=path, n=label: self._download_file(p, n)).grid(row=row, column=2, sticky="e", padx=5, pady=2)
                else:
                    ctk.CTkLabel(self.output_frame, text=label).grid(row=row, column=0, sticky="w", padx=5, pady=2)
                    ctk.CTkLabel(self.output_frame, text="Belum dibuat", text_color="gray").grid(row=row, column=1, sticky="w", padx=5, pady=2)
                    ctk.CTkLabel(self.output_frame, text="").grid(row=row, column=2, sticky="e", padx=5, pady=2)
                row += 1

    def _download_file(self, path, default_name):
        save_to = filedialog.asksaveasfilename(initialfile=default_name, defaultextension="*.*", filetypes=[("All files", "*.*")])
        if save_to:
            shutil.copy(path, save_to)
            messagebox.showinfo("Sukses", f"File disimpan ke:\n{save_to}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
