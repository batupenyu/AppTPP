import pandas as pd
import os
import tempfile
import time

perhitungan = pd.read_csv('perhitungan_ringkasan.csv')
perhitungan['nip'] = perhitungan['NAMA/NIP'].str.extract(r'NIP\.(\d+)')
perhitungan['nip'] = perhitungan['nip'].astype(str)
perhitungan = perhitungan.drop(columns=['NAMA/NIP'])

gaji = pd.read_csv('gaji.csv', sep=';', encoding='utf-8-sig')
gaji['NIP'] = gaji['NIP'].str.replace('="', '', regex=False).str.replace('"', '', regex=False)
gaji['NPWP'] = gaji['NPWP'].astype(str).str.replace('="', '', regex=False).str.replace('"', '', regex=False)
gaji['NO_REKENING'] = gaji['NO_REKENING'].astype(str).str.replace('="', '', regex=False).str.replace('"', '', regex=False)
gaji['NIP'] = gaji['NIP'].astype(str)
gaji = gaji.rename(columns={'NIP': 'nip'})

lembar3 = pd.read_csv('lembar3.csv')
lembar3['NIP'] = lembar3['NIP'].str.replace('="', '', regex=False).str.replace('"', '', regex=False)
lembar3['NIP'] = lembar3['NIP'].astype(str)
lembar3 = lembar3.rename(columns={'NIP': 'nip'})

nik = pd.read_excel('nik.xlsx')
nik['nip'] = nik['NIP Pegawai'].astype(str)
nik['NIK Pegawai'] = nik['NIK Pegawai'].astype(str)
nik = nik.drop(columns=['NIP Pegawai'])

utang = pd.read_excel('utang_bpjs.xlsx', dtype={'NIP': str}, usecols=['NIP', 'JANUARI'])
utang = utang.rename(columns={'NIP': 'nip', 'JANUARI': 'januari'})

result = perhitungan.merge(gaji, on='nip', how='left')
result = result.merge(lembar3, on='nip', how='left')
result = result.merge(nik, on='nip', how='left')
result = result.merge(utang, on='nip', how='left')

result = result.fillna('')

for col in ['Nama Pegawai.1', 'Nama Pegawai.2']:
    if col in result.columns:
        result[col] = result[col].astype(str).str.strip()
        result[col] = result[col].replace('nan', '')

result['Nama Pegawai'] = result['Nama Pegawai.1'].str.strip()
mask = result['Nama Pegawai.2'].astype(bool)
result.loc[mask, 'Nama Pegawai'] += ' ' + result.loc[mask, 'Nama Pegawai.2'].str.strip()

widths = {'nip': 17, 'NIK Pegawai': 15, 'NPWP Pegawai': 14}
for col, w in widths.items():
    if col in result.columns:
        result[col] = result[col].astype(str).str.zfill(w)

protected_cols = {'TEMPAT_BERTUGAS', 'KONDISI_KERJA', 'KELANGKAAN_PROFESI'}
cols_to_drop = []
for i in range(len(result.columns)):
    for j in range(i + 1, len(result.columns)):
        if result.columns[j] in protected_cols:
            continue
        if result.iloc[:, i].equals(result.iloc[:, j]):
            cols_to_drop.append(result.columns[j])
result = result.drop(columns=cols_to_drop)

if 'NAMA_x' in result.columns:
    has_comma = result['NAMA_x'].str.contains(',', na=False)
    split = result['NAMA_x'].str.split(',', n=1, expand=True)
    result['GELAR'] = split[0].str.strip()
    result['NAMA'] = split[1].str.strip().fillna('')
    result.loc[~has_comma, 'NAMA'] = result.loc[~has_comma, 'GELAR']
    result.loc[~has_comma, 'GELAR'] = ''
    result.loc[result['NAMA_x'] == '', 'NAMA'] = ''
    result.loc[result['NAMA_x'] == '', 'GELAR'] = ''
    result = result.drop(columns=['NAMA_x'])

drop_cols = ['GELAR', 'NAMA', 'NO', 'NAMA_y', 'Nama Pegawai.1', 'Nama Pegawai.2', 'NPWP Pegawai', 'Tanggal Lahir Pegawai']
result = result.drop(columns=[c for c in drop_cols if c in result.columns])

result.columns = [c.upper() for c in result.columns]

if 'NAMA PEGAWAI' in result.columns:
    cols = ['NAMA PEGAWAI'] + [c for c in result.columns if c != 'NAMA PEGAWAI']
    result = result[cols]

tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir='.')
os.close(tmp_fd)
result.to_excel(tmp_path, index=False)
for attempt in range(5):
    try:
        os.replace(tmp_path, 'gabung.xlsx')
        break
    except PermissionError:
        if attempt == 4:
            print("Error: unable to replace gabung.xlsx — close the file in Excel and try again.")
            raise
        time.sleep(1)
print(f'gabung.xlsx saved: {result.shape[0]} rows x {result.shape[1]} columns')