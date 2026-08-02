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
from docx.shared import Pt, Cm, Emu, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


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
                m2 = re.search(r"Rp\.[\d\.]+,\-\s*\(([^)]+)\)", val)
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


def add_tab_paragraph(doc, label, value, tab_pos_cm=4.5,
                      font_name="Arial", font_size=11, bold=False,
                      space_after=3, space_before=0):
    """Tambah paragraph dengan label di-sepakati kanan, colon, lalu value.
    Contoh output:
        Nama       : SYAHRYANTO, S.T.,M.Pd
        NIP        : 197708262006041005
        Jabatan    : Kepala SMK Negeri 1 Koba
    sehingga titik dua (:) align vertikal.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(tab_pos_cm))

    run_label = p.add_run(label)
    run_label.font.name = font_name
    run_label.font.size = Pt(font_size)
    run_label.font.bold = bold

    run_tab = p.add_run("\t")
    run_tab.font.name = font_name
    run_tab.font.size = Pt(font_size)

    run_colon = p.add_run(": ")
    run_colon.font.name = font_name
    run_colon.font.size = Pt(font_size)
    run_colon.font.bold = bold

    run_value = p.add_run(value)
    run_value.font.name = font_name
    run_value.font.size = Pt(font_size)
    run_value.font.bold = bold

    return p


def add_numbered_paragraph(doc, text, font_name="Arial", font_size=11,
                           space_after=12, space_before=0):
    """Tambah numbered paragraph (ordered list) untuk item 1, 2, 3, dst.
    Gunakan style 'List Number' agar Word menangani penomoran otomatis.
    Tambah line break (br) di akhir teks untuk spacing antar item.
    """
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.left_indent = None
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.add_break()

    return p


def add_indented_paragraph(doc, text, indent_px=200,
                           font_name="Arial", font_size=11, bold=False,
                           align=WD_ALIGN_PARAGRAPH.LEFT,
                           space_after=6, space_before=0):
    """Tambah paragraph dengan left indent dalam pixel."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.left_indent = Emu(indent_px * 12700)

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
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    kop_path = Path(__file__).parent / "kop_surat.jpg"
    if kop_path.exists():
        doc.add_picture(str(kop_path), width=Inches(5.8))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        last_paragraph.paragraph_format.space_after = Pt(6)

    add_paragraph(doc, "SURAT PERNYATAAN TANGGUNG JAWAB MUTLAK",
                  align=WD_ALIGN_PARAGRAPH.CENTER, bold=True,
                  font_size=14, space_after=12)

    add_paragraph(doc, "Yang bertanda tangan dibawah ini :", space_after=6)

    add_tab_paragraph(doc, "Nama", data['nama'], space_after=3)
    add_tab_paragraph(doc, "NIP", data['nip'], space_after=3)
    add_tab_paragraph(doc, "Jabatan", data['jabatan'], space_after=12)

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
        f"Perhitungan yang terdapat dalam SPM Langsung (SPM-LS) Nomor : {nomor_spm} "
        f"tanggal {tanggal_spm} untuk pembayaran Tambahan Penghasilan Pegawai (TPP) "
        f"Negeri Sipil sebesar {jumlah} ({terbilang}) untuk bulan {bulan} {tahun} "
        f"telah dihitung dengan benar berdasarkan dokumen pelaksanaan anggaran "
        f"dan dokumen pendukung lainnya."
    )
    add_numbered_paragraph(doc, p1, space_after=12)

    p2 = (
        "Apabila terdapat kesalahan dan kelebihan atas pembayaran, sebagaimana "
        "yang dimaksud pada point 1 (satu), kami bertanggung jawab dan bersedia "
        "untuk menyetorkan kelebihan tersebut ke Kas Daerah."
    )
    add_numbered_paragraph(doc, p2, space_after=12)

    p3 = (
        "Dokumen bukti-bukti belanja atas pembayaran tersebut di atas disimpan di "
        "Dinas Pendidikan Provinsi Kepulauan Bangka Belitung (SMK Negeri 1 Koba) "
        "sesuai ketentuan yang berlaku untuk kelengkapan administrasi dan keperluan "
        "pemeriksaan BPK dan/atau aparatur pengawas fungsional lainnya."
    )
    add_numbered_paragraph(doc, p3, space_after=12)

    lokasi_tanggal = data.get("lokasi") or "Koba,        Agustus 2026"
    if not lokasi_tanggal.startswith("Koba"):
        lokasi_tanggal = f"Koba,        {lokasi_tanggal}"

    add_indented_paragraph(doc, lokasi_tanggal, indent_px=200, space_after=6)
    add_indented_paragraph(doc, "Kepala Sekolah", indent_px=200, space_after=36)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_after = Pt(3)
    p_sig.paragraph_format.space_before = Pt(0)
    p_sig.paragraph_format.left_indent = Emu(200 * 12700)
    run_name = p_sig.add_run(data["nama"])
    run_name.font.name = "Arial"
    run_name.font.size = Pt(11)
    run_name.font.bold = True
    run_name.add_break()
    run_nip = p_sig.add_run(f"NIP. {data['nip']}")
    run_nip.font.name = "Arial"
    run_nip.font.size = Pt(11)

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
