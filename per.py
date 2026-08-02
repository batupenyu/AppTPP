#!/usr/bin/env python3
"""
pdf_to_csv.py — Ekstrak tabel dari PDF ke CSV, dengan kolom yang tetap
sejajar meskipun tabelnya bersambung ke banyak halaman.

Cara pakai:
    python3 pdf_to_csv.py "nama_file.pdf"
    python3 pdf_to_csv.py "nama_file.pdf" "hasil.csv"

Nama file PDF bisa apa saja / berubah-ubah, tinggal diisi lewat argumen.
Jika nama file output tidak diisi, otomatis dibuat dari nama file PDF
(contoh: laporan.pdf -> laporan.csv).

Kenapa kolom antar halaman sering bergeser:
    Secara default, pdfplumber mendeteksi ulang garis kolom di setiap
    halaman. Untuk tabel besar yang bersambung ke halaman berikutnya,
    hasil deteksi ini bisa sedikit berbeda per halaman (misalnya 28
    kolom di halaman 1 tapi 40 kolom di halaman 2), sehingga saat
    digabung ke CSV, datanya jadi geser.

Cara script ini mengatasinya:
    1. Kumpulkan semua posisi garis vertikal (x) dari SELURUH halaman.
    2. Gunakan posisi garis yang sama itu ("explicit_vertical_lines")
       untuk membaca tabel di setiap halaman, sehingga jumlah dan letak
       kolom pasti konsisten dari halaman pertama sampai terakhir.
    3. Jika PDF ternyata tidak punya garis tabel sama sekali (misalnya
       hasil scan), otomatis fallback ke ekstraksi teks biasa.
"""

import sys
import csv
from pathlib import Path

import pdfplumber
import pandas as pd


def clean_cell(value):
    """Ganti None dengan string kosong dan rapikan whitespace/newline di dalam sel."""
    if value is None:
        return ""
    return " ".join(str(value).split())


def collect_vertical_lines(pdf, tolerance=1):
    """Kumpulkan posisi x semua garis vertikal dari seluruh halaman PDF,
    lalu gabungkan garis yang posisinya sangat berdekatan (dianggap sama)."""
    xs = set()
    for page in pdf.pages:
        for line in page.lines:
            if abs(line["x0"] - line["x1"]) < 0.5:  # garis vertikal
                xs.add(round(line["x0"], 1))
        # beberapa PDF menggambar batas kolom sebagai persegi tipis (rect), bukan line
        for rect in page.rects:
            if rect["width"] < 1.5:
                xs.add(round(rect["x0"], 1))

    if not xs:
        return []

    sorted_xs = sorted(xs)
    merged = [sorted_xs[0]]
    for x in sorted_xs[1:]:
        if x - merged[-1] > tolerance:
            merged.append(x)
    return merged


def extract_pdf_to_csv(pdf_path: str, csv_path: str | None = None):
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {pdf_path}")

    if csv_path is None:
        csv_path = pdf_path.with_suffix(".csv")
    else:
        csv_path = Path(csv_path)

    all_rows = []
    tables_found = 0

    with pdfplumber.open(pdf_path) as pdf:
        vertical_lines = collect_vertical_lines(pdf)

        if vertical_lines:
            table_settings = {
                "vertical_strategy": "explicit",
                "explicit_vertical_lines": vertical_lines,
                "horizontal_strategy": "lines",
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "intersection_tolerance": 3,
            }
        else:
            table_settings = {}  # biarkan pdfplumber pakai default

        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables(table_settings=table_settings)

            for table_index, table in enumerate(tables, start=1):
                tables_found += 1
                all_rows.append([f"--- Halaman {page_number}, Tabel {table_index} ---"])
                for row in table:
                    cleaned_row = [clean_cell(cell) for cell in row]
                    all_rows.append(cleaned_row)
                all_rows.append([])  # baris kosong pemisah antar tabel

    if tables_found == 0:
        print("Tidak ada tabel yang terdeteksi. Mencoba ekstrak teks biasa sebagai gantinya...")
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                all_rows.append([f"--- Halaman {page_number} (teks) ---"])
                for line in text.split("\n"):
                    all_rows.append([line])
                all_rows.append([])

    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerows(all_rows)
    except PermissionError:
        # File tujuan kemungkinan sedang terbuka (misalnya di Excel) atau read-only.
        # Coba simpan dengan nama alternatif supaya proses tidak gagal total.
        alt_path = csv_path.with_name(csv_path.stem + "_baru" + csv_path.suffix)
        print(f"PERINGATAN: Tidak bisa menulis ke '{csv_path}' (izin ditolak).")
        print("Kemungkinan file itu sedang terbuka di Excel/program lain, atau bersifat read-only.")
        print(f"Mencoba menyimpan ke nama alternatif: {alt_path}")
        with open(alt_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerows(all_rows)
        csv_path = alt_path

    print(f"Selesai. {tables_found} tabel ditemukan.")
    print(f"Hasil disimpan di: {csv_path}")
    return csv_path


def to_number(series):
    """Ubah angka format Indonesia ('1.626.000' / '3.500') jadi angka (int),
    dengan menghapus titik ribuan. Nilai kosong/tidak valid jadi 0."""
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace(".", "", regex=False)   # hapus titik ribuan
        .str.replace(",", ".", regex=False)  # jaga-jaga kalau ada koma desimal
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0).astype("int64")


def build_summary(csv_path, summary_path=None, potongan_col_index=12):
    """
    Versi Python dari query Power Query yang diberikan:
      1. Cari baris header asli ('NO, NAMA/NIP, GOL, ...') di dalam CSV mentah.
      2. Buang baris-baris di atasnya (judul, PD/UNIT KERJA, dsb).
      3. Buang baris kosong / baris header yang terulang di tiap halaman
         (dikenali dari kolom GOL yang kosong, '3', atau 'GOL').
      4. Ambil kolom BESARAN TPP (RP) dan kolom potongan kinerja
         (posisi kolom ke-13, sesuai '13=9+11' pada tabel asli -> total
         potongan SKP + presensi), lalu hitung JLH KOTOR = BESARAN - POTONGAN.
      5. Simpan hanya kolom: NAMA/NIP, GOL, JABATAN/ESELON, KELAS JABATAN,
         BESARAN TPP (RP), POTONGAN, JLH KOTOR.

    Catatan: posisi kolom potongan (potongan_col_index, default kolom ke-13
    / index 12) mengikuti struktur tabel TPP ini persis seperti pada query
    Power Query aslinya. Jika suatu saat struktur tabel PDF berbeda, angka
    ini mungkin perlu disesuaikan.
    """
    csv_path = Path(csv_path)
    if summary_path is None:
        summary_path = csv_path.with_name(csv_path.stem + "_ringkasan.csv")
    else:
        summary_path = Path(summary_path)

    # Baca manual dengan csv.reader karena file mentah punya jumlah kolom
    # yang tidak seragam per baris (baris penanda "--- Halaman ... ---" vs
    # baris data tabel), sehingga pd.read_csv biasa akan gagal/parser error.
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    max_cols = max((len(r) for r in rows), default=0)
    padded_rows = [r + [""] * (max_cols - len(r)) for r in rows]
    raw = pd.DataFrame(padded_rows)

    # 1. Cari baris header asli (kolom pertama == "NO")
    header_row_idx = None
    for idx, val in raw[0].items():
        if str(val).strip().upper() == "NO":
            header_row_idx = idx
            break

    if header_row_idx is None:
        print("PERINGATAN: Baris header ('NO, NAMA/NIP, GOL, ...') tidak ditemukan.")
        print("Proses ringkasan dilewati.")
        return None

    header = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = range(data.shape[1])  # kerja dengan posisi kolom (index), bukan nama

    # Pastikan jumlah kolom cukup untuk mengambil kolom potongan
    if data.shape[1] <= potongan_col_index:
        print(f"PERINGATAN: Tabel hanya punya {data.shape[1]} kolom, "
              f"butuh minimal {potongan_col_index + 1}. Proses ringkasan dilewati.")
        return None

    col_nama = 1       # NAMA/NIP
    col_gol = 2        # GOL
    col_jabatan = 3    # JABATAN/ESELON
    col_kelas = 4      # KELAS JABATAN
    col_besaran = 5    # BESARAN TPP (RP)
    col_potongan = potongan_col_index  # total potongan kinerja (13=9+11)

    # 3. Filter baris: GOL tidak kosong, bukan "3" (baris nomor formula),
    #    dan bukan "GOL" (header yang terulang di tiap halaman)
    gol = data[col_gol].astype(str).str.strip()
    mask = (gol != "") & (gol != "3") & (gol.str.upper() != "GOL")
    data = data[mask].copy()

    # 4. Hitung JLH KOTOR
    besaran = to_number(data[col_besaran])
    potongan = to_number(data[col_potongan])
    jlh_kotor = besaran - potongan

    summary = pd.DataFrame({
        "NAMA/NIP": data[col_nama].astype(str).str.strip(),
        "GOL": data[col_gol].astype(str).str.strip(),
        "JABATAN/ESELON": data[col_jabatan].astype(str).str.strip(),
        "KELAS JABATAN": data[col_kelas].astype(str).str.strip(),
        "BESARAN TPP (RP)": besaran,
        "POTONGAN": potongan,
        "JLH KOTOR": jlh_kotor,
    })

    try:
        summary.to_csv(summary_path, index=False, encoding="utf-8-sig")
    except PermissionError:
        alt_path = summary_path.with_name(summary_path.stem + "_baru" + summary_path.suffix)
        print(f"PERINGATAN: Tidak bisa menulis ke '{summary_path}' (izin ditolak).")
        print(f"Mencoba menyimpan ke nama alternatif: {alt_path}")
        summary.to_csv(alt_path, index=False, encoding="utf-8-sig")
        summary_path = alt_path

    print(f"Ringkasan ({len(summary)} baris) disimpan di: {summary_path}")
    return summary_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Cara pakai: python3 pdf_to_csv.py <file.pdf> [output.csv]")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_csv = sys.argv[2] if len(sys.argv) > 2 else None

    raw_csv_path = extract_pdf_to_csv(input_pdf, output_csv)
    build_summary(raw_csv_path)
