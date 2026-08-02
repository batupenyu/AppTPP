"""
extract_tpp_pdf.py
-------------------
Ekstrak tabel "DAFTAR NOMINATIF PEMBAYARAN TPP" dari file PDF (mis. lembar3.pdf)
ke CSV, meniru logika Power Query berikut:

    Source            -> Pdf.Tables(...)
    Table002          -> ambil tabel data
    Promoted Headers  -> baris pertama jadi nama kolom
    Changed Type      -> ubah tipe kolom (angka dibersihkan dari titik ribuan)
    Filtered Rows     -> buang baris "index kolom" (GOL = "3") & baris non-data lain

Cara pakai:
    pip install pdfplumber pandas --break-system-packages
    python extract_tpp_pdf.py "lembar3.pdf" -o "lembar3.csv"

Bisa dipakai untuk semua file sejenis (lembar1.pdf, lembar2.pdf, dst) karena
posisi garis kolom dideteksi otomatis dari halaman pertama, bukan di-hardcode.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

# Nama kolom final (sesuai header PDF, newline diganti spasi/underscore)
COLUMNS = [
    "NO",
    "NAMA_NIP",
    "GOL",
    "JABATAN",
    "KELAS_JABATAN",
    "BESARAN_TPP",
    "BEBAN_KERJA",
    "TEMPAT_BERTUGAS",
    "KONDISI_KERJA",
    "PRESTASI_KERJA",
    "KELANGKAAN_PROFESI",
]

NUMERIC_COLUMNS = [
    "KELAS_JABATAN",
    "BESARAN_TPP",
    "BEBAN_KERJA",
    "TEMPAT_BERTUGAS",
    "KONDISI_KERJA",
    "PRESTASI_KERJA",
    "KELANGKAAN_PROFESI",
]


def detect_column_boundaries(page):
    """
    Deteksi batas vertikal kolom tabel dari sebuah halaman, dengan mencari
    tabel (strategi 'lines') yang punya persis 11 kolom pada baris header.
    Mengembalikan list koordinat X (11+1 garis) atau None jika tidak ketemu.
    """
    for table in page.find_tables(table_settings={"vertical_strategy": "lines",
                                                    "horizontal_strategy": "lines"}):
        if not table.rows:
            continue
        first_row_cells = table.rows[0].cells
        if len(first_row_cells) == len(COLUMNS):
            xs = sorted({round(c[0], 2) for c in first_row_cells} |
                        {round(c[2], 2) for c in first_row_cells})
            if len(xs) == len(COLUMNS) + 1:
                return xs
    return None


def extract_rows(pdf_path):
    """Ekstrak semua baris tabel (mentah, sebelum dibersihkan) dari seluruh halaman."""
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        # cari batas kolom dari halaman pertama yang berhasil dideteksi
        vlines = None
        for page in pdf.pages:
            vlines = detect_column_boundaries(page)
            if vlines:
                break
        if not vlines:
            raise RuntimeError(
                "Tidak berhasil mendeteksi struktur tabel 11 kolom di PDF ini."
            )

        settings = {
            "vertical_strategy": "explicit",
            "explicit_vertical_lines": vlines,
            "horizontal_strategy": "lines",
            "intersection_tolerance": 5,
        }

        for page in pdf.pages:
            for table in page.extract_tables(table_settings=settings):
                for row in table:
                    if len(row) == len(COLUMNS):
                        all_rows.append(row)
    return all_rows


def clean_number(value):
    """'1.626.000' -> 1626000 ; '' / None -> 0"""
    if value is None:
        return 0
    value = str(value).strip()
    if value == "":
        return 0
    value = value.replace(".", "").replace(",", "")
    if not re.fullmatch(r"-?\d+", value):
        return None  # bukan angka -> baris ini bukan data
    return int(value)


def build_dataframe(raw_rows):
    df = pd.DataFrame(raw_rows, columns=COLUMNS)

    # --- buang baris header berulang / judul / baris kosong ---
    df = df[df["NO"].astype(str).str.strip().str.fullmatch(r"\d+")]

    # --- Filtered Rows: setara `each ([GOL] <> "3")` di Power Query,        ---
    # --- yaitu membuang baris index kolom (1,2,3,...,11) yang GOL-nya "3"  ---
    df = df[df["GOL"].astype(str).str.strip() != "3"]

    # --- pisahkan NAMA_NIP (mengandung newline) jadi NAMA & NIP ---
    def split_nama_nip(text):
        parts = str(text).split("\n")
        nama = parts[0].strip() if len(parts) > 0 else ""
        nip = ""
        for p in parts[1:]:
            m = re.search(r"NIP\.?\s*([\dA-Za-z]+)", p)
            if m:
                nip = m.group(1).strip()
        return pd.Series({"NAMA": nama, "NIP": nip})

    df = df.join(df["NAMA_NIP"].apply(split_nama_nip))
    df = df.drop(columns=["NAMA_NIP"])

    # --- Changed Type: bersihkan & ubah kolom angka ---
    for col in NUMERIC_COLUMNS:
        df[col] = df[col].apply(clean_number)

    # baris yang gagal dikonversi ke angka (None) berarti bukan baris data -> buang
    df = df.dropna(subset=NUMERIC_COLUMNS)
    for col in NUMERIC_COLUMNS:
        df[col] = df[col].astype("int64")

    df["NO"] = df["NO"].astype(str).str.strip()
    df["GOL"] = df["GOL"].astype(str).str.strip()
    df["JABATAN"] = df["JABATAN"].astype(str).str.replace("\n", " ", regex=False).str.strip()

    # urutkan ulang kolom
    df = df[["NO", "NAMA", "NIP", "GOL", "JABATAN", "KELAS_JABATAN",
             "BESARAN_TPP", "BEBAN_KERJA", "TEMPAT_BERTUGAS",
             "KONDISI_KERJA", "PRESTASI_KERJA", "KELANGKAAN_PROFESI"]]

    df = df.reset_index(drop=True)
    return df


def main():
    parser = argparse.ArgumentParser(description="Ekstrak tabel TPP dari PDF ke CSV")
    parser.add_argument("pdf_path", help="Path ke file PDF (mis. lembar3.pdf)")
    parser.add_argument("output_path_positional", nargs="?", default=None,
                         help="Path file CSV output (boleh ditulis langsung tanpa -o)")
    parser.add_argument("-o", "--output", default=None,
                         help="Path file CSV output (default: sama nama dengan PDF)")
    args = parser.parse_args()

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        sys.exit(f"File tidak ditemukan: {pdf_path}")

    output_arg = args.output or args.output_path_positional
    output_path = Path(output_arg) if output_arg else pdf_path.with_suffix(".csv")

    raw_rows = extract_rows(pdf_path)
    df = build_dataframe(raw_rows)

    # Paksa Excel membaca NIP sebagai teks (bukan angka), supaya tidak
    # berubah jadi notasi ilmiah (mis. 1,97E+17) walaupun dibuka langsung
    # atau diproses lewat Text to Columns.
    df_out = df.copy()
    df_out["NIP"] = "=\"" + df_out["NIP"].astype(str) + "\""

    df_out.to_csv(output_path, index=False, encoding="utf-8-sig")

    print(f"Berhasil mengekstrak {len(df)} baris data pegawai.")
    print(f"Disimpan ke: {output_path}")
    print("\nRingkasan total (untuk verifikasi terhadap baris 'Total' di PDF):")
    print(f"  BESARAN_TPP     : {df['BESARAN_TPP'].sum():,}".replace(",", "."))
    print(f"  BEBAN_KERJA     : {df['BEBAN_KERJA'].sum():,}".replace(",", "."))
    print(f"  PRESTASI_KERJA  : {df['PRESTASI_KERJA'].sum():,}".replace(",", "."))


if __name__ == "__main__":
    main()
