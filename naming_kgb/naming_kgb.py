import os
import shutil
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime

# Setup paths - use current directory
base_path = Path.cwd()
output_base = base_path / 'kgb'
sheet_name = 'PNS'
# Set pages per person
PAGES_PER_PERSON = 1

# Find Excel and PDF files automatically
excel_files = list(base_path.glob('*.xlsx'))
pdf_files = list(base_path.glob('*.pdf'))

if not excel_files:
    print("ERROR: Tidak ada file Excel (.xlsx) yang ditemukan!")
    exit(1)
if not pdf_files:
    print("ERROR: Tidak ada file PDF yang ditemukan!")
    exit(1)

excel_file = excel_files[0]
pdf_file = pdf_files[0]

print(f"Menggunakan file Excel: {excel_file.name}")
print(f"Menggunakan file PDF: {pdf_file.name}")



# Baca Excel
print("=" * 100)
print("STEP 1: Membaca file Excel")
print("=" * 100)

try:
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
except Exception as e:
    print(f"ERROR: Gagal membaca file Excel: {e}")
    exit(1)

print(f"Data dari Excel berhasil dibaca: {len(df)} record")

# Process data and create grouped structure
grouped_data = {}
for idx, row in df.iterrows():
    nama = row['nama'].strip()
    unit_kerja = row['unit_kerja'].strip()
    tanggal = pd.to_datetime(row['baru_mulai_tanggal'])

    # Extract bulan
    bulan_indo = {
        1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
        5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
        9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }
    bulan = bulan_indo.get(tanggal.month, str(tanggal.month))
    tahun = tanggal.year

    # Group by unit_kerja (direktorat)
    if unit_kerja not in grouped_data:
        grouped_data[unit_kerja] = []
    
    grouped_data[unit_kerja].append({
        'idx': idx,
        'nama': nama,
        'unit_kerja': unit_kerja,
        'direktorat': unit_kerja,
        'bulan': bulan,
        'tahun': tahun,
        'tanggal': tanggal
    })

# Baca PDF
print("\n" + "=" * 100)
print("STEP 2: Membaca file PDF")
print("=" * 100)

try:
    reader = PdfReader(str(pdf_file))
    num_pages = len(reader.pages)
    print(f"File PDF memiliki {num_pages} halaman")
    print(f"Expected records: {len(df)}")
    print(f"Pages per person: {PAGES_PER_PERSON}")
except Exception as e:
    print(f"ERROR: Gagal membaca file PDF: {e}")
    exit(1)

# Buat struktur folder
print("\n" + "=" * 100)
print("STEP 3: Membuat struktur folder berdasarkan Direktorat")
print("=" * 100)

for direktorat in sorted(grouped_data.keys()):
    folder_path = output_base / direktorat
    folder_path.mkdir(parents=True, exist_ok=True)
    print(f"[OK] {folder_path.relative_to(base_path)}")

# Split dan simpan PDF
print("\n" + "=" * 100)
print("STEP 4: Split PDF dan menyimpan file per person")
print("=" * 100)

# Flatten grouped_data untuk akses sequential
all_records = []
for direktorat in sorted(grouped_data.keys()):
    for item in grouped_data[direktorat]:
        all_records.append(item)

all_records.sort(key=lambda x: x['idx'])

# Proses setiap orang (2 halaman per orang)
num_people = len(all_records)
print(f"Total records yang akan diproses: {num_people}")

for person_idx in range(num_people):
    record = all_records[person_idx]
    
    # Calculate page range for this person
    start_page = person_idx * PAGES_PER_PERSON
    end_page = min(start_page + PAGES_PER_PERSON, num_pages)
    
    if start_page >= num_pages:
        print(f"[{person_idx+1:2d}/{num_people}] WARNING: Halaman {start_page} tidak tersedia")
        continue
    
    # Extract pages for this person
    writer = PdfWriter()
    for page_idx in range(start_page, end_page):
        if page_idx < num_pages:
            writer.add_page(reader.pages[page_idx])
    
    filename = f"{record['tahun']}_{record['bulan']}_{record['nama']}_kenaikan_gaji_berkala.pdf"
    
    # Path lengkap
    file_path = output_base / record['direktorat'] / filename
    
    # Simpan file
    try:
        with open(str(file_path), 'wb') as output_file:
            writer.write(output_file)
        print(f"[{person_idx+1:2d}/{num_people}] Saved: {record['nama']} (pages {start_page+1}-{end_page})")
        print(f"           Location: {file_path.relative_to(base_path)}")
    except Exception as e:
        print(f"[{person_idx+1:2d}/{num_people}] ERROR saving {record['nama']}: {e}")
        continue

print("\n" + "=" * 100)
print("STEP 5: Verifikasi hasil")
print("=" * 100)

# Count files
total_files = 0
summary = {}

for direktorat in sorted(grouped_data.keys()):
    folder_path = output_base / direktorat
    if folder_path.exists():
        files = list(folder_path.glob('*.pdf'))
        total_files += len(files)
        summary[direktorat] = len(files)
        print(f"[OK] {direktorat}: {len(files)} file")

print("\n" + "=" * 100)
print("SUMMARY")
print("=" * 100)
print(f"Total file yang dihasilkan: {total_files}")
print(f"Lokasi folder: {output_base}")
print(f"File Excel yang digunakan: {excel_file.name}")
print(f"File PDF yang digunakan: {pdf_file.name}")

if summary:
    print("\nDetail per direktorat:")
    for direktorat, count in summary.items():
        print(f"  - {direktorat}: {count} file")

print("\n" + "=" * 100)
print("SELESAI!")
print("=" * 100)
