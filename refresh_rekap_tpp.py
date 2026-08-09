#!/usr/bin/env python3
"""
refresh_rekap_tpp.py
--------------------
Refresh sheet 'rekap_tpp' pada TPP PNS.xlsm dan TPP P3K.xlsm
dengan data terbaru dari gabung.xlsx.

Menggunakan Excel COM (win32com) agar VBA, external data connections,
dan drawings tetap terjaga.
"""

import os
import shutil
import sys
import time
from pathlib import Path

import pandas as pd
import win32com.client

BASE = Path(__file__).resolve().parent
GABUNG = BASE / "gabung.xlsx"

TPP_FILES = {
    "PNS": BASE / "_usulan_tpp_smkn1_koba" / "PNS" / "TPP PNS.xlsm",
    "PPPK": BASE / "_usulan_tpp_smkn1_koba" / "PPPK" / "TPP P3K.xlsm",
}


def load_gabung():
    if not GABUNG.exists():
        raise FileNotFoundError(f"gabung.xlsx tidak ditemukan: {GABUNG}")
    df = pd.read_excel(GABUNG)
    if "GOLONGAN" in df.columns:
        df = df.drop(columns=["GOLONGAN"])
    return df


def refresh_rekap_tpp(tpp_path, df):
    if not tpp_path.exists():
        raise FileNotFoundError(f"File TPP tidak ditemukan: {tpp_path}")

    backup_path = tpp_path.with_suffix(".xlsm.bak")
    shutil.copy2(tpp_path, backup_path)

    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(str(tpp_path))
        ws = wb.Worksheets("rekap_tpp")

        ws.Activate()
        xl.ActiveWindow.ScrollRow = 1

        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        header_row = 2
        first_data_row = 3

        if last_row >= first_data_row:
            ws.Range(ws.Rows(first_data_row), ws.Rows(last_row)).Delete()

        headers = [ws.Cells(header_row, c).Value for c in range(1, ws.UsedRange.Columns.Count + 1)]
        df_cols = list(df.columns)

        rekap_headers = [h for h in headers if h is not None]
        if df_cols != rekap_headers:
            raise ValueError(
                f"Kolom gabung.xlsx tidak cocok dengan header rekap_tpp.\n"
                f"gabung: {df_cols}\n"
                f"rekap_tpp: {rekap_headers}"
            )

        nip_col = rekap_headers.index("NIP") + 1

        for r_idx, row in enumerate(df.itertuples(index=False), start=first_data_row):
            for c_idx, value in enumerate(row, start=1):
                if c_idx == nip_col:
                    value = str(value) if value is not None else ""
                    ws.Cells(r_idx, c_idx).Value = "'" + value
                else:
                    ws.Cells(r_idx, c_idx).Value = value

        ws.Columns(nip_col).NumberFormat = "@"

        wb.Save()
        print(f"[OK] {tpp_path.name}: {len(df)} baris ditulis ke rekap_tpp")
    except Exception:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if xl is not None:
            xl.Quit()
        if backup_path.exists():
            shutil.copy2(backup_path, tpp_path)
        raise
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        if backup_path.exists():
            os.remove(backup_path)


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else "all"
    target = target.upper()

    if target == "ALL":
        targets = TPP_FILES.items()
    elif target in TPP_FILES:
        targets = [(target, TPP_FILES[target])]
    else:
        print(f"Usage: python refresh_rekap_tpp.py [PNS|PPPK|all]")
        print(f"       all (default) = refresh kedua file")
        sys.exit(1)

    df = load_gabung()
    print(f"Data gabung.xlsx: {len(df)} baris, {len(df.columns)} kolom")

    for jenis, path in targets:
        print(f"Refreshing {jenis}...")
        refresh_rekap_tpp(path, df)

    print("[OK] Selesai.")


if __name__ == "__main__":
    main()
