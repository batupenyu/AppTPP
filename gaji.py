"""
gaji.py
-------
Ekstrak tabel "DAFTAR PEMBAYARAN GAJI INDUK PPPK" dari file PDF
(mis. gaji_all.pdf) ke CSV, meniru pola yang dipakai di bayar.py:

    Source            -> extract_words() + clustering kolom        (pdfplumber)
    Band pegawai      -> baris diawali NAMA + STATUS
    Column center     -> deteksi otomatis dari digit halaman pertama
    Changed Type      -> angka dibersihkan dari titik ribuan
    Filtered Rows     -> satukan pecahan angka, buang baris tanpa NAMA

Perbedaan dengan bayar.py:
    PDF ini tidak punya garis tabel, jadi kita pakai ekstraksi berbasis
    kata (extract_words) dengan deteksi band pegawai dan clustering
    pusat kolom dari digit.

Cara pakai:
    pip install pdfplumber pandas --break-system-packages
    python gaji.py "gaji_all.pdf" -o "gaji.csv"
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

DEBUG = False

# Batas bawah header (data mulai ~150)
HEADER_TOP_MAX = 155
TOP_MIN = 150

# Konstanta deteksi band & kolom
NAME_X_MIN = 30
NAME_X_MAX = 130
STATUS_X_MIN = 185
STATUS_X_MAX = 215
NUMERIC_X_MIN = 255
NUMERIC_X_MAX = 830
CLUSTER_GAP = 14
BAND_HEIGHT = 95

# Kolom teks
TEXT_COLUMNS = [
    "NO",
    "NAMA",
    "STATUS",
    "TGL_LAHIR",
    "NIP",
    "NPWP",
    "NO_REKENING",
    "JMLH",
    "Unit_kerja",
]


def _is_digitish(tok):
    s = tok.replace(",", "").replace(".", "").replace("-", "").replace(" ", "")
    return s.isdigit()


def _clean_number(value):
    if value is None:
        return 0
    value = str(value).strip().replace(" ", "")
    if value == "":
        return 0
    value = value.replace(".", "").replace(",", "")
    if not re.fullmatch(r"-?\d+", value):
        return None
    return int(value)


def find_employee_bands(words):
    lines = {}
    for w in words:
        if w["top"] < TOP_MIN:
            continue
        lines.setdefault(round(w["top"], 1), []).append(w)

    if DEBUG:
        all_xs = sorted(set(w["x0"] for w in words if w["top"] >= TOP_MIN))
        print(f"[DEBUG] Unique x0 positions (top>={TOP_MIN}): {all_xs[:30]}")
        name_words = [w for w in words if NAME_X_MIN <= w["x0"] <= NAME_X_MAX and any(c.isalpha() for c in w["text"]) and w["top"] >= TOP_MIN]
        status_words = [w for w in words if STATUS_X_MIN <= w["x0"] <= STATUS_X_MAX and w["text"][:1].isalpha() and w["top"] >= TOP_MIN]
        print(f"[DEBUG] Words in NAME range ({NAME_X_MIN}-{NAME_X_MAX}): {len(name_words)}")
        if name_words:
            for w in name_words[:5]:
                print(f"  NAME: x0={w['x0']:.1f} text='{w['text']}' top={w['top']:.1f}")
        print(f"[DEBUG] Words in STATUS range ({STATUS_X_MIN}-{STATUS_X_MAX}): {len(status_words)}")
        if status_words:
            for w in status_words[:5]:
                print(f"  STATUS: x0={w['x0']:.1f} text='{w['text']}' top={w['top']:.1f}")

    band_starts = []
    for top in sorted(lines):
        ws = lines[top]
        name = [w for w in ws
                if NAME_X_MIN <= w["x0"] <= NAME_X_MAX
                and any(c.isalpha() for c in w["text"])]
        status = [w for w in ws
                  if STATUS_X_MIN <= w["x0"] <= STATUS_X_MAX
                  and w["text"][:1].isalpha()]
        if name and status:
            band_starts.append(top)

    bands = []
    for i, bt in enumerate(band_starts):
        nxt = band_starts[i + 1] if i + 1 < len(band_starts) else bt + BAND_HEIGHT
        end = min(bt + BAND_HEIGHT, nxt)
        bwords = [w for w in words
                  if bt - 1 <= w["top"] < end and w["top"] >= TOP_MIN]
        bands.append((bt, bwords))
    return bands


def build_column_centers(pdf):
    page = pdf.pages[0]
    words = page.extract_words()
    bands = find_employee_bands(words)
    if not bands:
        if DEBUG:
            print("[DEBUG] No bands found. All words on page 1:")
            for w in words[:30]:
                print(f"  x0={w['x0']:.1f} x1={w['x1']:.1f} top={w['top']:.1f} text='{w['text']}'")
        return []

    all_xs = []
    for bt, bwords in bands:
        for w in bwords:
            if not _is_digitish(w["text"]):
                continue
            cx = (w["x0"] + w["x1"]) / 2
            if NUMERIC_X_MIN <= cx <= NUMERIC_X_MAX:
                all_xs.append(cx)

    all_xs.sort()

    clusters = []
    cur = [all_xs[0]]
    for x in all_xs[1:]:
        if x - cur[-1] <= CLUSTER_GAP:
            cur.append(x)
        else:
            clusters.append(cur)
            cur = [x]
    clusters.append(cur)

    centers = [round(sum(c) / len(c), 1) for c in clusters]

    header_words = [w for w in words
                    if w["top"] < HEADER_TOP_MAX
                    and any(ch.isalpha() for ch in w["text"])]

    labeled = []
    for xc in centers:
        bloc = "P" if 255 <= xc <= 560 else "O"
        best, bd = None, 1e9
        for h in header_words:
            hx = (h["x0"] + h["x1"]) / 2
            if bloc == "P" and not (255 <= hx <= 560):
                continue
            if bloc == "O" and not (600 <= hx <= 830):
                continue
            d = abs(hx - xc)
            if d < bd:
                bd, best = d, h
        label = best["text"] if best else f"COL_{xc:.0f}"
        labeled.append((xc, label))
    return labeled


def extract_band(band_top, bwords, column_centers):
    rec = {}

    nos = [w for w in bwords if 30 <= w["x0"] <= 42 and _is_digitish(w["text"])]
    if nos:
        rec["NO"] = min(nos, key=lambda w: w["top"])["text"].strip()

    name_words = [w for w in bwords
                  if NAME_X_MIN <= w["x0"] <= 220
                  and any(c.isalpha() for c in w["text"])
                  and (w["top"] - band_top) < 8]
    name_words.sort(key=lambda w: (round(w["top"]), w["x0"]))
    if name_words:
        rec["NAMA"] = " ".join(w["text"] for w in name_words).strip()

    stat = [w for w in bwords if STATUS_X_MIN <= w["x0"] <= STATUS_X_MAX
            and w["text"][:1].isalpha() and (w["top"] - band_top) < 8]
    if stat:
        rec["STATUS"] = " ".join(w["text"] for w in stat).strip()

    tgl = [w for w in bwords if NAME_X_MIN <= w["x0"] <= 165
           and re.fullmatch(r"\d{1,2}-\d{1,2}-\d{4}", w["text"])]
    if tgl:
        rec["TGL_LAHIR"] = tgl[0]["text"].strip()

    nip_words = [w for w in bwords if NAME_X_MIN <= w["x0"] <= 158
                 and _is_digitish(w["text"])
                 and 12 <= (w["top"] - band_top) <= 32]
    if nip_words:
        nip_words.sort(key=lambda w: w["top"])
        nip_raw = "".join(w["text"].strip() for w in nip_words)
        digits_only = re.sub(r"\D", "", nip_raw)
        rec["NIP"] = digits_only[-18:] if len(digits_only) >= 18 else digits_only

    npwp = [w for w in bwords if NAME_X_MIN <= w["x0"] <= 175
             and _is_digitish(w["text"].replace(".", "").replace("-", ""))
             and 40 <= (w["top"] - band_top) <= 58]
    if npwp:
        npwp.sort(key=lambda w: w["top"])
        rec["NPWP"] = "".join(w["text"].strip() for w in npwp)

    rek = [w for w in bwords if 700 <= w["x0"] <= 800
            and _is_digitish(w["text"])
            and 30 <= (w["top"] - band_top) <= 48]
    if rek:
        rek.sort(key=lambda w: w["top"])
        rec["NO_REKENING"] = "".join(w["text"].strip() for w in rek)

    for xc, label in column_centers:
        toks = [w for w in bwords
                if abs((w["x0"] + w["x1"]) / 2 - xc) <= 15
                and _is_digitish(w["text"])]
        if not toks:
            continue
        toks.sort(key=lambda w: w["top"])
        parts = []
        cur_part = ""
        last_top = None
        for w in toks:
            if last_top is not None and abs(w["top"] - last_top) < 3:
                cur_part += w["text"].strip().replace(" ", "")
            else:
                if cur_part:
                    parts.append(cur_part)
                cur_part = w["text"].strip().replace(" ", "")
            last_top = w["top"]
        if cur_part:
            parts.append(cur_part)
        rec[label] = parts[-1] if parts else ""

    jb_toks = [w for w in bwords
               if 680 <= (w["x0"] + w["x1"]) / 2 <= 720
               and _is_digitish(w["text"])
               and 17.5 <= (w["top"] - band_top) <= 18.5]
    if jb_toks:
        jb_toks.sort(key=lambda w: w["top"])
        jb_parts = []
        cur_part = ""
        last_top = None
        for w in jb_toks:
            if last_top is not None and abs(w["top"] - last_top) < 3:
                cur_part += w["text"].strip().replace(" ", "")
            else:
                if cur_part:
                    jb_parts.append(cur_part)
                cur_part = w["text"].strip().replace(" ", "")
            last_top = w["top"]
        if cur_part:
            jb_parts.append(cur_part)
        rec["PEG"] = jb_parts[-1] if jb_parts else ""

    jmlh_toks = [w for w in bwords
                  if 230 <= (w["x0"] + w["x1"]) / 2 <= 250
                  and _is_digitish(w["text"])
                  and 8.5 <= (w["top"] - band_top) <= 10.5]
    if jmlh_toks:
        jmlh_toks.sort(key=lambda w: w["top"])
        jmlh_parts = []
        cur_part = ""
        last_top = None
        for w in jmlh_toks:
            if last_top is not None and abs(w["top"] - last_top) < 3:
                cur_part += w["text"].strip().replace(" ", "")
            else:
                if cur_part:
                    jmlh_parts.append(cur_part)
                cur_part = w["text"].strip().replace(" ", "")
            last_top = w["top"]
        if cur_part:
            jmlh_parts.append(cur_part)
        rec["JMLH"] = jmlh_parts[-1] if jmlh_parts else ""

    return rec


def extract_all(pdf_path):
    all_recs = []
    page_unit_kerja = {}
    with pdfplumber.open(pdf_path) as pdf:
        column_centers = build_column_centers(pdf)
        if not column_centers:
            sys.exit("Tidak dapat mendeteksi kolom gaji di halaman pertama.")
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            unit_kerja = ""
            if text:
                for line in text.split("\n"):
                    if line.startswith("[") and "DINAS" in line:
                        unit_kerja = line.strip()
                        break
            page_unit_kerja[page_num] = unit_kerja
            words = page.extract_words()
            bands = find_employee_bands(words)
            for bt, bwords in bands:
                rec = extract_band(bt, bwords, column_centers)
                if rec.get("NAMA"):
                    rec["_page"] = page_num
                    all_recs.append(rec)
    return all_recs, page_unit_kerja


def main():
    parser = argparse.ArgumentParser(
        description="Ekstrak Daftar Pembayaran Gaji Induk PPPK ke CSV")
    parser.add_argument("pdf_path")
    parser.add_argument("output_path_positional", nargs="?", default=None)
    parser.add_argument("-o", "--output", default=None)
    parser.add_argument("--debug", action="store_true", help="Tampilkan info debug untuk diagnosis layout")
    args = parser.parse_args()

    global DEBUG
    DEBUG = args.debug

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        sys.exit(f"File tidak ditemukan: {pdf_path}")

    output_arg = args.output or args.output_path_positional
    output_path = Path(output_arg) if output_arg else pdf_path.with_suffix(".csv")

    recs, page_unit_kerja = extract_all(pdf_path)
    if not recs:
        sys.exit("Tidak ada data pegawai yang terdeteksi.")

    df = pd.DataFrame(recs)

    df["Unit_kerja"] = df["_page"].map(page_unit_kerja)
    df = df.drop(columns=["_page"])

    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    angka_cols = [c for c in df.columns if c not in TEXT_COLUMNS]
    for col in angka_cols:
        df[col] = df[col].apply(_clean_number)

    df["gaji_net"] = df["PEG"].apply(lambda v: int(v) * 100)

    df["jumlah_anak"] = df.apply(
        lambda r: str(r.get("STATUS", "").split("-")[0]) + "/" + str(int(r.get("JMLH", 0)) if str(r.get("JMLH", "0")).isdigit() else 0),
        axis=1,
    )

    df_out = df[["NAMA", "NIP", "TGL_LAHIR", "NPWP", "NO_REKENING", "Unit_kerja", "gaji_net", "jumlah_anak"]].copy()
    df_out.columns = [c.upper() for c in df_out.columns]
    for c in ["NIP", "NO_REKENING", "NPWP"]:
        if c in df_out.columns:
            df_out[c] = '="' + df_out[c].astype(str) + '"'

    df_out.to_csv(output_path, index=False, encoding="utf-8-sig", sep=";")

    print(f"Berhasil mengekstrak {len(df)} pegawai.")
    print(f"Disimpan ke: {output_path}")
    print(f"Jumlah kolom: {len(df_out.columns)}")


if __name__ == "__main__":
    main()
