"""
bayar.py
--------
Ekstrak tabel "DAFTAR NOMINATIF PEMBAYARAN TAMBAHAN PENGHASILAN" dari file PDF
(mis. pembayaran.pdf) ke CSV, meniru pola yang dipakai di lem.py:

    Source            -> Pdf.Tables(...)                (pdfplumber)
    Table             -> ambil tabel data (16 kolom)
    Promoted Headers  -> baris pertama jadi nama kolom
    Changed Type      -> angka dibersihkan dari titik ribuan
    Filtered Rows     -> buang baris header berulang / baris nomor / baris TOTAL

Perbedaan dengan lem.py:
    PDF ini punya kepala bertingkat (multi-row header) dan kolom terakhir
    "JUMLAH BERSIH" dipisah oleh garis vertikal internal, sehingga di beberapa
    halaman pdfplumber mendeteksi 31 kolom. Solusinya sama seperti lem.py:
    deteksi 17 garis batas kolom (16 kolom) dari halaman PERTAMA yang
    memilikinya, lalu pakai explicit_vertical_lines supaya semua halaman
    dibaca dengan jumlah/kolom yang konsisten.

Cara pakai:
    pip install pdfplumber pandas --break-system-packages
    python bayar.py "pembayaran.pdf" -o "pembayaran.csv"

Bisa dipakai untuk file sejenis karena posisi garis kolom dideteksi otomatis
dari halaman pertama, bukan di-hardcode. (Fallback: jika tidak ketemu
17 garis, pakai garis yang terdeteksi otomatis per halaman.)
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

# Nama kolom final (sesuai header PDF, dipipihkan menjadi 1 baris)
COLUMNS = [
    "NO",
    "NAMA_NIP",
    "GOL",
    "JABATAN",
    "KELAS_JABATAN",
    "JUMLAH_TPP_SEBELUM_PAJAK",
    "TUNJANGAN_PLT",
    "VOL_BLN",
    "JUMLAH_KOTOR",
    "PPh_GOL_IX",
    "PPh_GOL_IV",
    "PPh_GOL_III",
    "PPh_GOL_X",
    "PPh_GOL_XI",
    "JUMLAH_PPH",
    "JUMLAH_BERSIH",
]

# Kolom angka RUPIAH / bilangan bulat -> dibersihkan jadi int
NUMERIC_COLUMNS = [
    "KELAS_JABATAN",
    "JUMLAH_TPP_SEBELUM_PAJAK",
    "TUNJANGAN_PLT",
    "VOL_BLN",
    "JUMLAH_KOTOR",
    "JUMLAH_BERSIH",
]

# Kolom PERSENTASE (tarif PPh) -> dibiarkan sebagai teks yang sudah rapi
PERCENT_COLUMNS = [
    "PPh_GOL_IX",
    "PPh_GOL_IV",
    "PPh_GOL_III",
    "PPh_GOL_X",
    "PPh_GOL_XI",
    "JUMLAH_PPH",
]


def detect_column_boundaries(page):
    """
    Cari 17 garis batas vertikal (=> 16 kolom) dari sebuah halaman, dengan
    mendeteksi tabel (strategi 'lines') yang punya baris data 16 kolom
    lengkap (garis kiri & kanan tiap kolom terbentuk). Mengembalikan list
    koordinat X (17 garis) atau None jika tidak ketemu.
    """
    for table in page.find_tables(table_settings={"vertical_strategy": "lines",
                                                    "horizontal_strategy": "lines"}):
        for row in table.rows:
            cells = row.cells
            if len(cells) != len(COLUMNS):
                continue
            good = [c for c in cells if c]
            if not good:
                continue
            xs = sorted({round(c[0], 2) for c in good} |
                        {round(c[2], 2) for c in good})
            if len(xs) == len(COLUMNS) + 1:
                return xs
    return None


def extract_rows(pdf_path):
    """Ekstrak semua baris tabel (mentah) dari seluruh halaman."""
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        # cari 17 garis batas dari halaman pertama yang memilikinya
        vlines = None
        for page in pdf.pages:
            vlines = detect_column_boundaries(page)
            if vlines:
                break

        if vlines:
            settings = {
                "vertical_strategy": "explicit",
                "explicit_vertical_lines": vlines,
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
            }
        else:
            # fallback: biarkan pdfplumber deteksi sendiri per halaman
            settings = {}

        for page in pdf.pages:
            for table in page.extract_tables(table_settings=settings):
                for row in table:
                    if len(row) == len(COLUMNS):
                        all_rows.append(row)
    return all_rows


def clean_number(value):
    """'1.626.000' -> 1626000 ; '' / None -> 0 ; non-angka -> None (bukan baris data)"""
    if value is None:
        return 0
    value = str(value).strip()
    if value == "":
        return 0
    value = value.replace(".", "").replace(",", "")
    if not re.fullmatch(r"-?\d+", value):
        return None  # bukan angka -> baris ini bukan data
    return int(value)


def clean_percent(value):
    """'0.00 %' -> '0.00%' ; '' / None -> '' (dibiarkan sebagai teks)"""
    if value is None:
        return ""
    return str(value).strip().replace(" ", "")


def build_dataframe(raw_rows):
    df = pd.DataFrame(raw_rows, columns=COLUMNS)

    # --- buang baris header berulang / judul / baris nomor kolom ---
    # baris data punya NO berupa angka
    df = df[df["NO"].astype(str).str.strip().str.fullmatch(r"\d+")]

    # --- buang baris TOTAL (NO = "TOTAL" / "TOT") agar tidak masuk data ---
    df = df[~df["NO"].astype(str).str.strip().str.upper().isin(["TOTAL", "TOT", "TOTAL"])]

    # --- pisahkan NAMA_NIP (mengandung newline) jadi NAMA & NIP ---
    def split_nama_nip(text):
        parts = str(text).split("\n")
        nama = parts[0].strip() if parts else ""
        nip = ""
        for p in parts[1:]:
            m = re.search(r"NIP\.?\s*([\dA-Za-z]+)", p)
            if m:
                nip = m.group(1).strip()
        return pd.Series({"NAMA": nama, "NIP": nip})

    df = df.join(df["NAMA_NIP"].apply(split_nama_nip))
    df = df.drop(columns=["NAMA_NIP"])

    # --- ubah kolom angka ---
    for col in NUMERIC_COLUMNS:
        df[col] = df[col].apply(clean_number)

    # baris yang gagal dikonversi ke angka (None) berarti bukan baris data -> buang
    df = df.dropna(subset=NUMERIC_COLUMNS)
    for col in NUMERIC_COLUMNS:
        df[col] = df[col].astype("int64")

    # --- rapikan kolom persentase (biarkan sebagai teks) ---
    for col in PERCENT_COLUMNS:
        df[col] = df[col].apply(clean_percent)

    df["NO"] = df["NO"].astype(str).str.strip()
    df["GOL"] = df["GOL"].astype(str).str.strip()
    df["JABATAN"] = df["JABATAN"].astype(str).str.replace("\n", " ", regex=False).str.strip()

    # urutkan ulang kolom
    df = df[["NO", "NAMA", "NIP", "GOL", "JABATAN", "KELAS_JABATAN",
             "JUMLAH_TPP_SEBELUM_PAJAK", "TUNJANGAN_PLT", "VOL_BLN",
             "JUMLAH_KOTOR", "PPh_GOL_IX", "PPh_GOL_IV", "PPh_GOL_III",
             "PPh_GOL_X", "PPh_GOL_XI", "JUMLAH_PPH", "JUMLAH_BERSIH"]]

    df = df.reset_index(drop=True)
    return df


def main():
    parser = argparse.ArgumentParser(
        description="Ekstrak tabel Pembayaran Tambahan Penghasilan dari PDF ke CSV")
    parser.add_argument("pdf_path", help="Path ke file PDF (mis. pembayaran.pdf)")
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
    # berubah jadi notasi ilmiah (mis. 1,97E+17).
    df_out = df.copy()
    df_out["NIP"] = '="' + df_out["NIP"].astype(str) + '"'

    df_out.to_csv(output_path, index=False, encoding="utf-8-sig")

    print(f"Berhasil mengekstrak {len(df)} baris data pegawai.")
    print(f"Disimpan ke: {output_path}")
    print("\nRingkasan total (untuk verifikasi terhadap baris 'TOTAL' di PDF):")
    print(f"  JUMLAH_TPP_SEBELUM_PAJAK : {df['JUMLAH_TPP_SEBELUM_PAJAK'].sum():,}".replace(",", "."))
    print(f"  JUMLAH_KOTOR             : {df['JUMLAH_KOTOR'].sum():,}".replace(",", "."))
    print(f"  JUMLAH_BERSIH            : {df['JUMLAH_BERSIH'].sum():,}".replace(",", "."))


if __name__ == "__main__":
    main()
