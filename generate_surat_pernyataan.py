#!/usr/bin/env python3
"""
generate_surat_pernyataan.py
----------------------------
Membuat dokumen SURAT PERNYATAAN TANGGUNG JAWAB MUTLAK (SPTJM)
untuk pembayaran TPP dari data template yang ada di sheet 'sptjm'
dalam file Excel TPP PNS JULI 2026 SMKN 1 KOBA.xlsm.

Cara pakai:
    python generate_surat_pernyataan.py "TPP PNS JULI 2026 SMKN 1 KOBA.xlsm" \
        -o "SURAT_PERNYATAAN_PNS.docx"

Output:
    File Word (.docx) berisi surat pernyataan dengan format standir
"""

import argparse
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def extract_sptjm_data(xlsx_path):
    """Baca sheet 'sptjm' dan kembalikan dictionary data surat."""
    wb = load_workbook(xlsx_path, data_only=True)
    if "sptjm" not in wb.sheetnames:
        raise ValueError(f"Sheet 'sptjm' tidak ditemukan di {xlsx_path}. "
                         f"Sheet tersedia: {wb.sheetnames}")

    ws = wb["sptjm"]

    data = {
        "nama": "",
        "nip": "",
        "jabatan": "",
        "spm_nomor": "",
        "spm_tanggal": "",
        "jumlah_rupiah": "",
        "jumlah_terbilang": "",
        "bulan": "",
        "tahun": "",
        "lokasi": "",
        "bulan_tahun_tempat": "",
    }

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        for idx, cell in enumerate(row):
            val = str(cell).strip() if cell is not None else ""
            if not val:
                continue

            if "Nama" in val and idx + 1 < len(row) and row[idx + 1]:
                data["nama"] = str(row[idx + 1]).strip()

            if val == "NIP" and idx + 1 < len(row) and row[idx + 1]:
                data["nip"] = str(row[idx + 1]).strip()

            if "Jabatan" in val and idx + 1 < len(row) and row[idx + 1]:
                data["jabatan"] = str(row[idx + 1]).strip()

            if "Nomor" in val and ":" in val:
                after = val.split(":", 1)[1].strip()
                dots = re.findall(r"\.{3,}", after)
                if dots:
                    data["spm_nomor"] = after.split(dots[0])[0].strip()

            if "tanggal" in val.lower():
                parts = re.split(r"tanggal\s*", val, flags=re.IGNORECASE)
                if len(parts) > 1:
                    after = parts[1].strip()
                    dots = re.findall(r"\.{3,}", after)
                    if dots:
                        data["spm_tanggal"] = after.split(dots[0])[0].strip()

            if re.search(r"Rp\.\s*[\d\.]+,-", val):
                m = re.search(r"(Rp\.\s*[\d\.]+,-)", val)
                if m:
                    data["jumlah_rupiah"] = m.group(1).strip()
                m2 = re.search(r"\(([^)]+)\)", val)
                if m2:
                    data["jumlah_terbilang"] = m2.group(1).strip()

            if re.search(r"untuk bulan\s+\w+\s+\d{4}", val, re.IGNORECASE):
                m = re.search(r"untuk bulan\s+(\w+)\s+(\d{4})", val, re.IGNORECASE)
                if m:
                    data["bulan"] = m.group(1).strip()
                    data["tahun"] = m.group(2).strip()

            if re.match(r"^Koba,\s*\w+\s+\d{4}$", val.strip()):
                data["lokasi"] = val.strip()
                parts = val.split(",")
                if len(parts) >= 2:
                    data["bulan_tahun_tempat"] = parts[1].strip()

    return data


def set_paragraph_font(paragraph, font_name="Arial", font_size=11, bold=False):
    """Terapkan font ke semua run di paragraph."""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold


def add_paragraph(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT,
                  font_name="Arial", font_size=11, bold=False,
                  space_after=6, space_before=0):
    """Tambah paragraph dengan gaya yang konsisten."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    return p


def build_surat(data, output_path):
    doc = Document()

    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    add_paragraph(doc, "SURAT PERNYATAAN TANGGUNG JAWAB MUTLAK",
                  align=WD_ALIGN_PARAGRAPH.CENTER, bold=True,
                  font_size=14, space_after=12)

    add_paragraph(doc, "Yang bertanda tangan dibawah ini :", space_after=6)

    add_paragraph(doc, f"Nama\t: {data['nama']}", space_after=3)
    add_paragraph(doc, f"NIP\t: {data['nip']}", space_after=3)
    add_paragraph(doc, f"Jabatan\t: {data['jabatan']}", space_after=12)

    add_paragraph(doc, "Menyatakan dengan sesungguhnya bahwa :", space_after=6)

    nomor_spm = data.get("spm_nomor") or "..........................................."
    tanggal_spm = data.get("spm_tanggal") or ".........................................."
    jumlah = data.get("jumlah_rupiah") or ""
    terbilang = data.get("jumlah_terbilang") or ""
    bulan = data.get("bulan") or ""
    tahun = data.get("tahun") or ""

    if not jumlah and not terbilang:
        jumlah = "Rp. ....................................,-"
        terbilang = "..........................................."

    p1 = (
        f"1.\tPerhitungan yang terdapat dalam SPM Langsung (SPM-LS) Nomor : {nomor_spm} "
        f"tanggal {tanggal_spm} untuk pembayaran Tambahan Penghasilan Pegawai (TPP) "
        f"Negeri Sipil sebesar {jumlah} ({terbilang}) untuk bulan {bulan} {tahun} "
        f"telah dihitung dengan benar berdasarkan dokumen pelaksanaan anggaran "
        f"dan dokumen pendukung lainnya."
    )
    add_paragraph(doc, p1, space_after=6)

    p2 = (
        "2.\tApabila terdapat kesalahan dan kelebihan atas pembayaran, sebagaimana "
        "yang dimaksud pada point 1 (satu), kami bertanggung jawab dan bersedia "
        "untuk menyetorkan kelebihan tersebut ke Kas Daerah."
    )
    add_paragraph(doc, p2, space_after=6)

    p3 = (
        "3.\tDokumen bukti-bukti belanja atas pembayaran tersebut di atas disimpan di "
        "Dinas Pendidikan Provinsi Kepulauan Bangka Belitung (SMK Negeri 1 Koba) "
        "sesuai ketentuan yang berlaku untuk kelengkapan administrasi dan keperluan "
        "pemeriksaan BPK dan/atau aparatur pengawas fungsional lainnya."
    )
    add_paragraph(doc, p3, space_after=12)

    lokasi_tanggal = data.get("lokasi") or "Koba,        Agustus 2026"
    if not lokasi_tanggal.startswith("Koba"):
        lokasi_tanggal = f"Koba,        {lokasi_tanggal}"

    add_paragraph(doc, lokasi_tanggal, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6)
    add_paragraph(doc, "Kepala Sekolah", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=36)

    add_paragraph(doc, data["nama"], align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, space_after=3)
    add_paragraph(doc, f"NIP. {data['nip']}", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=3)

    output_path = Path(output_path)
    doc.save(str(output_path))
    print(f"Surat pernyataan disimpan ke: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Generate SURAT_PERNYATAAN_PNS.docx dari sheet sptjm Excel TPP")
    parser.add_argument("xlsx_path", help="Path ke file Excel TPP (mis. TPP PNS JULI 2026 SMKN 1 KOBA.xlsm)")
    parser.add_argument("-o", "--output", default=None,
                        help="Path output docx (default: SURAT_PERNYATAAN_PNS.docx)")
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx_path)
    if not xlsx_path.exists():
        sys.exit(f"File tidak ditemukan: {xlsx_path}")

    output_path = args.output or xlsx_path.with_name("SURAT_PERNYATAAN_PNS.docx")

    data = extract_sptjm_data(xlsx_path)
    build_surat(data, output_path)

    print(f"  Nama   : {data['nama']}")
    print(f"  NIP    : {data['nip']}")
    print(f"  Jabatan: {data['jabatan']}")
    print(f"  Bulan  : {data['bulan']} {data['tahun']}")
    print(f"  Jumlah : {data['jumlah_rupiah']}")


if __name__ == "__main__":
    main()
