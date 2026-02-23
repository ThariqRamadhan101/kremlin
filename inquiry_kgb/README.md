# KGB (Kenaikan Gaji Berkala) Data Inquiry Script

A Python script to inquiry and process KGB (periodic salary increase) data from the eHRM API.

## Overview

The script includes three comprehensive phases:
- **Phase 1**: Employee Data Inquiry - Fetches employee data from eHRM API
- **Phase 2**: Data Filtering - Filters data based on specified criteria  
- **Phase 3**: KGB Data Processing & Document Download - Processes KGB data, calculates salary increases, and downloads supporting documents

## Authentication Setup

### Option 1: Environment Variable (Recommended)

**Windows (Command Prompt):**
```cmd
# Temporary (current session only)
set KGB_AUTH_TOKEN=your_token_here
python kgb_inquiry.py

# Permanent (all sessions)
# 1. Press Windows Key + R, type sysdm.cpl
# 2. Go to Advanced tab → Environment Variables
# 3. Add User variable: KGB_AUTH_TOKEN = your_token_here
```

**Linux/Mac:**
```bash
# Temporary (current session only)
export KGB_AUTH_TOKEN=your_token_here
python kgb_inquiry.py

# Permanent (add to ~/.bashrc or ~/.zshrc)
echo 'export KGB_AUTH_TOKEN=your_token_here' >> ~/.bashrc
source ~/.bashrc
```

### Option 2: Command Line Token

```bash
python kgb_inquiry.py --token YOUR_AUTH_TOKEN
```

## Usage

### Complete Workflow (Default)
The script runs as a unified workflow with all three phases:

```bash
# Basic usage with environment token
python kgb_inquiry.py

# Filter by status only
python kgb_inquiry.py --status PNS
python kgb_inquiry.py --status PPK

# Filter by years of service
python kgb_inquiry.py --tahun-kerja 5

# Filter by months of service
python kgb_inquiry.py --bulan-kerja 6

# Filter by NIP
python kgb_inquiry.py --nip 123456789

# Disable KGB eligibility filter (show all employees)
python kgb_inquiry.py --no-kgb-eligibility

# Combined filters
python kgb_inquiry.py --status PNS --tahun-kerja 5 --bulan-kerja 6

# With explicit token
python kgb_inquiry.py --token YOUR_AUTH_TOKEN --status PNS --tahun-kerja 5
```

### Workflow Process
1. **Phase 1**: Fetches all employee data from API
2. **Phase 2**: Automatically filters the data from Phase 1
3. **Phase 3**: Processes KGB data, calculates new salaries, and downloads documents
4. **Output**: Multi-sheet Excel file with organized data and document folder

## Parameters

The script accepts the following parameters:
- `--token`: Authentication token for API access (optional if KGB_AUTH_TOKEN env var is set)
- `--status`: Filter by employee status (PNS or PPK)
- `--tahun-kerja`: Filter by minimum years of service (integer)
- `--bulan-kerja`: Filter by minimum months of service (integer)
- `--nip`: Filter by specific employee NIP (integer)
- `--no-kgb-eligibility`: Disable KGB eligibility filter (show all employees regardless of 2-year cycle)

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Requirements

- Python 3.7+
- Valid eHRM API authentication token
- Internet connection to access the API

## File Configuration

### 1. Daftar Gaji PNS.xlsx & Daftar Gaji PPPK.xlsx
Required for salary lookup based on golongan and masa kerja for each employee status.

**Expected Format:**
```
TAHUN | GOLONGAN | MK TAHUN | KODE | KODE2 | GAJI
2025  | III/a    | 6       | ...  | ...  | 4801000
2025  | III/b    | 6       | ...  | ...  | 5057000
```

### 2. pejabat_penandatangan.json
Required for signing officer data based on golongan range and employee status.

**Expected Format (JSON):**
```json
{
  "pejabat_penandatangan_data": [
    {
      "golongan_pns": ["I/a", "I/b", "I/c", "I/d", "II/a", "II/b", "II/c", "II/d"],
      "golongan_pppk": ["II", "III", "IV", "VI", "VII", "VIII"],
      "jabatan": "Direktur Jenderal Prasarana Strategis",
      "nama": "Bisma Staniarto, S.T., M.Sc.",
      "nip": "196803191996031001",
      "nomor_surat_pns": "PB.03.05/...",
      "nomor_surat_pppk": "PK.03.05/...",
      "special_case": true
    }
  ]
}
```

### 3. pejabat.json
Contains special case NIP mappings for specific individuals.

### 4. tembusan_mapping.json
Maps unit kerja to tembusan (CC recipients) for document distribution.

### 5. urutan_unit_kerja.json
Defines the ordering of unit kerja for consistent sorting in output files.

**Special Case Handling:**
- **special_case: true**: Indicates this entry should only be matched by NIP, not by golongan range
- **Priority**: Special cases are checked first before regular golongan matching
- **Use Case**: For specific individuals who should sign regardless of their golongan
- **Status Support**: Different golongan ranges for PNS and PPPK employees
- **Document Numbers**: Separate surat numbers for PNS and PPPK

**Alternative Formats (CSV/YAML):**
- **CSV**: Comma-separated values with headers
- **YAML**: YAML format with proper indentation
- **JSON**: Recommended (default format)

**Golongan Ranges:**
- **PNS Golongan I & II**: Kepala Bagian Kepegawaian dan Umum
- **PNS Golongan III & IV**: Sekretaris Direktorat Jenderal Prasarana Strategis
- **PPPK Golongan II, III, IV, VI, VII, VIII**: Direktur Jenderal Prasarana Strategis
- **Special Cases**: Matched by NIP regardless of golongan

## Document Download Feature

### Automatic Document Download
The script automatically downloads Kenaikan Pangkat and Kenaikan Gaji Berkala documents for each employee.

### Download Process:
1. **Source**: Documents are downloaded from arsip paths in API detail responses
2. **URL Pattern**: `https://ehrm.pu.go.id/api/cuti/get-arsip-cuti?arsip={encoded_path}`
3. **Folder Structure**: 
   ```
   documents/
   ├── 001_Employee_Name_1_Kenaikan Pangkat.pdf
   ├── 001_Employee_Name_1_Kenaikan Gaji Berkala.pdf
   ├── 002_Employee_Name_2_Kenaikan Pangkat.pdf
   ├── 002_Employee_Name_2_Kenaikan Gaji Berkala.pdf
   └── ...
   ```

### Features:
- **Ordered Folders**: Folders are numbered (001, 002, etc.) to match Excel order
- **Clean Names**: Employee names are sanitized for folder naming
- **Error Handling**: Failed downloads are logged and tracked
- **Summary Report**: Download summary Excel file with statistics

### Output Files:
- `kgb_processed_data_YYYYMMDD_HHMMSS.xlsx` - Main KGB data with multiple sheets:
  - **PNS**: PNS employee data
  - **PPPK**: PPPK employee data  
  - **Summary_by_Unit_Status_Pejabat**: Summary statistics
- `documents/` - Folder with all downloaded PDFs
- Console output with processing statistics

## Authentication Setup

### Option 1: Environment Variable (Recommended)

**Windows (Command Prompt):**
```cmd
# Temporary (current session only)
set KGB_AUTH_TOKEN=your_token_here
python kgb_inquiry.py

# Permanent (all sessions)
# 1. Press Windows Key + R, type sysdm.cpl
# 2. Go to Advanced tab → Environment Variables
# 3. Add User variable: KGB_AUTH_TOKEN = your_token_here
```

**Linux/Mac:**
```bash
# Temporary (current session only)
export KGB_AUTH_TOKEN=your_token_here
python kgb_inquiry.py

# Permanent (add to ~/.bashrc or ~/.zshrc)
echo 'export KGB_AUTH_TOKEN=your_token_here' >> ~/.bashrc
source ~/.bashrc
```

### Option 2: Command Line Token

```bash
python kgb_inquiry.py --token YOUR_AUTH_TOKEN
```

## Detailed Process Flow

### Phase 1: Employee Data Inquiry
1. **API Authentication**: Script authenticates using Bearer token from environment variable or command line
2. **Employee Search**: Sends POST request to `https://ehrm.pu.go.id/api/pegawai/search` with payload:
   ```json
   {
     "nip": "", 
     "nama": "", 
     "jabatan": "jenderal prasarana strategis",
     "on_page": 1, 
     "page_size": 100000000000,
     "status_asal": "1,5,6", 
     "status_akhir": "1"
   }
   ```
3. **Data Processing**: For each employee:
   - Parse masa kerja (years and months from "X Thn Y Bln" format)
   - Extract pangkat and golongan from strings like "III/a (Penata Tingkat I)"
   - Determine unit kerja from jabatan uraian using priority matching
   - Select jabatan based on priority: STRUKTURAL > PERBENDAHARAAN > FUNGSIONAL > PELAKSANA

### Phase 2: Data Filtering
1. **Apply Filters** (AND logic):
   - **Status**: Exact match (PNS/PPPK)
   - **Tahun Kerja**: Exact match on years
   - **Bulan Kerja**: Exact match on months  
   - **NIP**: Exact match if specified
   - **Unit Kerja**: Must not be empty
   - **KGB Eligibility**: 
     - Even years: Golongan I/b, I/c, I/d, II/b, II/c, II/d with even masa kerja years
     - Odd years: All other golongan with odd masa kerja years
     - PPPK adjustment: +3 years to masa kerja for eligibility calculation

### Phase 3: KGB Data Processing & Document Download
1. **Data Enrichment** for each filtered employee:
   - Fetch latest gaji pokok data from `/api/gaji-pokok/get`
   - Fetch latest pangkat data from `/api/pangkat/get`
   - Get detailed information using detail APIs
   - Determine "acuan terakhir" (latest reference) by comparing dates
   - Calculate new salary based on golongan and masa kerja + 1 year
   - Get pejabat penandatangan based on golongan and status
   - Get tembusan based on unit kerja mapping

2. **Salary Calculations**:
   - **New Masa Kerja**: Current years + 1 (PPPK: +4 years total)
   - **New Salary**: Lookup from Daftar Gaji files based on golongan and new masa kerja
   - **Terbilang**: Convert salary amounts to Indonesian words
   - **Effective Date**: 2 months from now, set to day 1

3. **Document Download**:
   - Download Kenaikan Pangkat and Kenaikan Gaji Berkala PDFs
   - URL pattern: `https://ehrm.pu.go.id/api/cuti/get-arsip-cuti?arsip={encoded_path}`
   - Save as: `001_Employee_Name_Kenaikan Pangkat.pdf`
   - Organized in `documents/` folder with sequential numbering

4. **Output Generation**:
   - **Multi-sheet Excel** with custom column ordering
   - **PNS Sheet**: PNS employees sorted by pejabat, unit kerja, masa kerja, nama
   - **PPPK Sheet**: PPPK employees with same sorting
   - **Summary Sheet**: Statistics by unit kerja, status, and pejabat
   - **Validation**: Automatic workbook integrity check

### Error Handling & Logging
- Comprehensive try-catch blocks with line number reporting
- Graceful handling of missing files, network errors, and API failures
- Detailed console output for debugging and progress tracking
- Automatic fallback to default values when data is missing

## Filtering Logic

- **Status Filter**: Exact match (case-insensitive) on employee status
- **Tahun Kerja Filter**: Employees with years == specified value (exact match)
- **Bulan Kerja Filter**: Employees with months == specified value (exact match)
- **NIP Filter**: Exact match on employee NIP
- **Unit Kerja Filter**: Employees must have non-empty unit_kerja
- **KGB Eligibility Filter**: 
  - **Even years**: Golongan I/b, I/c, I/d, II/b, II/c, II/d with even masa kerja years
  - **Odd years**: All other golongan with odd masa kerja years
  - **PPPK Adjustment**: Additional 3 years added to masa kerja for PPPK employees
- **Combined Filters**: All filters are applied together (AND logic)

## API Endpoint

The script queries: `https://ehrm.pu.go.id/api/pegawai/search`

With payload:
```json
{
  "nip": "",
  "nama": "",
  "jabatan": "jenderal prasarana strategis",
  "on_page": 1,
  "page_size": 100000000000,
  "status_asal": "1,5,6",
  "status_akhir": "1"
}
```

## Security Benefits

- ✅ No tokens stored in files
- ✅ No accidental commits to version control
- ✅ System-level security
- ✅ Compatible with corporate security policies

## Troubleshooting

**Token not found:**
- Ensure `KGB_AUTH_TOKEN` environment variable is set
- Restart terminal after setting permanent variables
- Check for typos in variable name

**API connection issues:**
- Verify internet connection
- Check token validity
- Ensure API endpoint is accessible

**No filtered results:**
- Check if filter values are too restrictive
- Verify data contains expected values
- Try without filters first
