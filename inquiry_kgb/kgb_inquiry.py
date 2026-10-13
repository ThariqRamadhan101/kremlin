#!/usr/bin/env python3
"""
KGB (Kenaikan Gaji Berkala) Data Inquiry Script
Phase 1: Employee Data Inquiry from Directorate General of Strategic Infrastructure

This script fetches employee data from the eHRM API and processes it for KGB analysis.
"""

import requests
import pandas as pd
import json
import os
import sys
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Union
import urllib.parse
import argparse
from dateutil.relativedelta import relativedelta

import config


class KGBInquiry:
    def __init__(
        self,
        auth_token: str,
        status: str = None,
        tahun_kerja: int = None,
        bulan_kerja: int = None,
        nip: str = None,
        kgb_eligibility: bool = True,
    ):
        """
        Initialize KGB Inquiry with authentication and filtering parameters.

        Args:
            auth_token (str): Bearer token for API authentication
            status (str): Filter by employee status (PNS/PPK)
            tahun_kerja (int): Filter by years of service
            bulan_kerja (int): Filter by months of service
            nip (str): Filter by employee NIP
            kgb_eligibility (bool): Filter by KGB eligibility (odd years only)
        """
        self.auth_token = auth_token
        self.status = status
        self.tahun_kerja = tahun_kerja
        self.bulan_kerja = bulan_kerja
        self.nip = nip
        self.kgb_eligibility = kgb_eligibility

        self.base_url = config.BASE_URL
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": self.auth_token,
            "User-Agent": config.USER_AGENT,
        }

    def search_employees(self) -> Dict:
        """
        Search for employees in Directorate General of Strategic Infrastructure.

        Returns:
            Dict: API response containing employee data
        """
        url = f"{self.base_url}{config.SEARCH_PEGAWAI_PATH}"

        payload = {
            "nip": f"{self.nip}" if self.nip else "",
            "nama": "",
            "jabatan": config.SEARCH_JABATAN,
            "on_page": config.SEARCH_ON_PAGE,
            "page_size": config.SEARCH_PAGE_SIZE,
            "status_asal": config.SEARCH_STATUS_ASAL,
            "status_akhir": config.SEARCH_STATUS_AKHIR,
        }

        try:
            print("Fetching employee data from API...")
            response = requests.post(url, headers=self.headers, json=payload)
            response.raise_for_status()

            data = response.json()
            print(
                f"Successfully retrieved data for {len(data.get('list_pegawai', []))} employees"
            )
            return data

        except requests.exceptions.RequestException as e:
            print(f"Error fetching employee data: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"Error parsing API response: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            sys.exit(1)

    def get_jabatan_by_priority(self, jabatan_list: List[Dict]) -> Dict:
        """
        Select jabatan based on priority order.
        Priority: STRUKTURAL > PERBENDAHARAAN > FUNGSIONAL > PELAKSANA

        Args:
            jabatan_list (List[Dict]): List of jabatan dictionaries

        Returns:
            Dict: Selected jabatan based on priority
        """
        if not jabatan_list or len(jabatan_list) == 0:
            return {}

        priority_order = [
            "JABATAN STRUKTURAL",
            "JABATAN PERBENDAHARAAN",
            "JABATAN FUNGSIONAL",
            "JABATAN PELAKSANA",
        ]

        # Search for jabatan in priority order
        for priority_jabatan in priority_order:
            for jabatan in jabatan_list:
                if jabatan.get("nama_jabatan", "") == priority_jabatan:
                    return jabatan

        # If no match found, return the first one
        return jabatan_list[0] if jabatan_list else {}

    def extract_unit_kerja(self, uraian: str) -> str:
        """
        Extract unit kerja from jabatan uraian.

        Args:
            uraian (str): Full jabatan description

        Returns:
            str: Extracted unit kerja name
        """
        target_units = [
            "Sekretariat Direktorat Jenderal Prasarana Strategis",
            "Direktorat Sistem dan Strategi Penyelenggaraan Prasarana Strategis",
            "Direktorat Infrastruktur Dukungan Pendidikan",
            "Direktorat Infrastruktur Dukungan Perekonomian, Peribadatan, Kesehatan, Olahraga dan Sosial Budaya",
            "Direktorat Kepatuhan Intern",
        ]

        # Check for Satuan Kerja Pelaksanaan pattern
        if ", Pelaksanaan Prasarana Strategis" in uraian:
            # Extract region after comma and handle multiple commas
            parts = uraian.split(", Pelaksanaan Prasarana Strategis")
            if len(parts) > 1:
                region = parts[1].strip()
                # Get first part before any additional commas
                region_part = region.split(",")[0].strip()
                return f"Satuan Kerja Pelaksanaan Prasarana Strategis {region_part}"
            return "Lainnya"

        # Check for target units
        for unit in target_units:
            if unit in uraian:
                return unit

        # Check for specific directorate patterns
        if "Direktur Sistem dan Strategi Penyelenggaraan Prasarana Strategis" in uraian:
            return "Direktorat Sistem dan Strategi Penyelenggaraan Prasarana Strategis"
        elif "Direktur Infrastruktur Dukungan Pendidikan" in uraian:
            return "Direktorat Infrastruktur Dukungan Pendidikan"
        elif (
            "Direktur Infrastruktur Dukungan Perekonomian, Peribadatan, Kesehatan, Olahraga dan Sosial Budaya"
            in uraian
        ):
            return "Direktorat Infrastruktur Dukungan Perekonomian, Peribadatan, Kesehatan, Olahraga dan Sosial Budaya"
        elif "Direktur Kepatuhan Intern" in uraian:
            return "Direktorat Kepatuhan Intern"
        elif "Sekretaris Direktorat Jenderal Prasarana Strategis" in uraian:
            return "Sekretariat Direktorat Jenderal Prasarana Strategis"
        elif "Direktur Jenderal Prasarana Strategis" in uraian:
            return "Direktorat Jenderal Prasarana Strategis"

        # Default fallback
        return "Lainnya"

    def parse_pangkat_golongan(self, golongan: str) -> tuple:
        """
        Parse golongan string to extract pangkat and golongan.

        Args:
            golongan (str): String like "IX (Mahir/Ahli Pertama)" or "II/d (Pengatur Tingkat I)"

        Returns:
            tuple: (pangkat, golongan_code)
        """
        try:
            if not golongan:
                return "", ""

            # Extract the main golongan code (before parentheses)
            if "(" in golongan:
                golongan_code = golongan.split("(")[0].strip()
                pangkat = golongan.split("(")[1].split(")")[0].strip()
            else:
                golongan_code = golongan.strip()
                pangkat = ""

            return pangkat, golongan_code
        except (ValueError, IndexError):
            print(f"Error parsing pangkat/golongan: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return "", golongan.strip() if golongan else ""

    def parse_masa_kerja(self, masakerja: str) -> tuple:
        """
        Parse masa kerja string into years and months.

        Args:
            masakerja (str): String like "0 Thn 3 Bln"

        Returns:
            tuple: (years, months)
        """
        try:
            if not masakerja:
                return 0, 0

            parts = masakerja.split()
            years = 0
            months = 0

            for i, part in enumerate(parts):
                if part == "Thn" and i > 0:
                    years = int(parts[i - 1])
                elif part == "Bln" and i > 0:
                    months = int(parts[i - 1])

            return years, months
        except (ValueError, IndexError):
            print(f"Error parsing masa kerja: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return 0, 0

    def filter_employees(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filter employee data based on provided parameters.

        Args:
            df (pd.DataFrame): Original employee data

        Returns:
            pd.DataFrame: Filtered employee data
        """
        filtered_df = df.copy()

        # Filter by status
        if self.status:
            filtered_df = filtered_df[
                filtered_df["status"].str.upper() == self.status.upper()
            ]

        # Filter by tahun kerja
        if self.tahun_kerja is not None:
            filtered_df = filtered_df[
                filtered_df["masa_kerja_tahun"] == self.tahun_kerja
            ]

        # Filter by bulan kerja
        if self.bulan_kerja is not None:
            filtered_df = filtered_df[
                filtered_df["masa_kerja_bulan"] == self.bulan_kerja
            ]

        # Filter by unit kerja (must not be empty)
        filtered_df = filtered_df[
            filtered_df["unit_kerja"].notna() & (
                filtered_df["unit_kerja"] != "")
        ]

        # Filter by KGB eligibility - optional based on parameter
        print(f"Applying KGB eligibility filter: {self.kgb_eligibility}")
        if self.kgb_eligibility:
            # Golongan yang eligible pada tahun GENAP (even years)
            even_year_golongan = config.EVEN_YEAR_GOLONGAN

            # Create eligibility mask
            eligible_mask = False

            # Check even-year golongan with even years
            for golongan in even_year_golongan:
                mask_golongan = filtered_df["golongan"] == golongan
                mask_even_year = filtered_df["masa_kerja_tahun"] % 2 == 0
                eligible_mask = eligible_mask | (
                    mask_golongan & mask_even_year)

            # Check odd-year golongan with odd years
            # All other golongan not in even_year list are eligible on odd years
            mask_not_even = ~filtered_df["golongan"].isin(even_year_golongan)
            mask_odd_year = filtered_df["masa_kerja_tahun"] % 2 == 1
            eligible_mask = eligible_mask | (mask_not_even & mask_odd_year)

            eligible_df = filtered_df[eligible_mask]
        else:
            eligible_df = filtered_df

        # Count KGB eligible employees
        print(f"KGB eligible: {len(eligible_df)}")
        print(f"KGB ineligible: {len(filtered_df) - len(eligible_df)}")
        print(f"Total after filtering: {len(eligible_df)}")
        return eligible_df

    def process_employee_data(self, api_response: Dict) -> pd.DataFrame:
        """
        Process employee data from API response into DataFrame.

        Args:
            api_response (Dict): API response containing employee data

        Returns:
            pd.DataFrame: Processed employee data
        """
        if "list_pegawai" not in api_response:
            print("No employee data found in API response")
            return pd.DataFrame()

        employees = api_response["list_pegawai"]
        # Extract relevant fields for KGB processing
        processed_data = []
        for emp in employees:
            # Parse masa kerja
            masa_kerja_tahun, masa_kerja_bulan = self.parse_masa_kerja(
                emp.get("masakerja", "")
            )

            # Parse pangkat and golongan
            pangkat, golongan_code = self.parse_pangkat_golongan(
                emp.get("golongan", "")
            )

            # Extract unit kerja from jabatan uraian
            unit_kerja = ""
            if "jabatan" in emp and len(emp["jabatan"]) > 0:
                # Select jabatan based on priority
                selected_jabatan = self.get_jabatan_by_priority(emp["jabatan"])
                jabatan = selected_jabatan.get(
                    "uraian", "").split(",")[0].strip()
                uraian = selected_jabatan.get("uraian", "")
                unit_kerja = self.extract_unit_kerja(uraian)

            processed_emp = {
                "id": emp.get("id", ""),
                "nama": emp.get("nama", ""),
                "nip": emp.get("nip", ""),
                "golongan": golongan_code,
                "pangkat": pangkat,
                "masa_kerja_tahun": masa_kerja_tahun,
                "masa_kerja_bulan": masa_kerja_bulan,
                "status": emp.get("status", ""),
                "unit_kerja": unit_kerja,
                "jabatan": jabatan,
            }
            processed_data.append(processed_emp)

        df = pd.DataFrame(processed_data)
        print(f"Processed {len(df)} employee records")
        return df

    def clean_gaji_amount(self, gaji_str: str) -> int:
        """
        Clean gaji string to extract numeric value only.

        Args:
            gaji_str (str): Gaji string like "Rp. 3.743.100" or "Rp 3,743,100"

        Returns:
            int: Clean numeric value
        """
        try:
            if not gaji_str:
                return 0

            # Remove "Rp", "Rp.", dots, spaces, and other non-numeric characters except commas
            cleaned = (
                gaji_str.replace("Rp", "")
                .replace("Rp.", "")
                .replace(".", "")
                .replace(" ", "")
                .replace(",", "")
            )

            # Extract only digits
            numeric_str = "".join(filter(str.isdigit, cleaned))

            return int(numeric_str) if numeric_str else 0
        except ValueError:
            print(f"Error cleaning gaji amount: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return 0

    def terbilang_indonesia(self, angka: int) -> str:
        """
        Convert number to Indonesian words (terbilang).

        Args:
            angka (int): Number to convert

        Returns:
            str: Number in Indonesian words
        """
        if angka == 0:
            return "Nol Rupiah"

        def _terbilang(n: int) -> str:
            if n < 12:
                satuan = [
                    "",
                    "Satu",
                    "Dua",
                    "Tiga",
                    "Empat",
                    "Lima",
                    "Enam",
                    "Tujuh",
                    "Delapan",
                    "Sembilan",
                    "Sepuluh",
                    "Sebelas",
                ]
                return satuan[n]
            elif n < 20:
                return _terbilang(n - 10) + " Belas"
            elif n < 100:
                if n % 10 == 0:
                    return _terbilang(n // 10) + " Puluh"
                else:
                    return _terbilang(n // 10) + " Puluh " + _terbilang(n % 10)
            elif n < 200:
                return "Seratus " + _terbilang(n - 100)
            elif n < 1000:
                if n % 100 == 0:
                    return _terbilang(n // 100) + " Ratus"
                else:
                    return _terbilang(n // 100) + " Ratus " + _terbilang(n % 100)
            elif n < 2000:
                return "Seribu " + _terbilang(n - 1000)
            elif n < 1000000:
                if n % 1000 == 0:
                    return _terbilang(n // 1000) + " Ribu"
                else:
                    return _terbilang(n // 1000) + " Ribu " + _terbilang(n % 1000)
            elif n < 1000000000:
                if n % 1000000 == 0:
                    return _terbilang(n // 1000000) + " Juta"
                else:
                    return _terbilang(n // 1000000) + " Juta " + _terbilang(n % 1000000)
            elif n < 1000000000000:
                if n % 1000000000 == 0:
                    return _terbilang(n // 1000000000) + " Miliar"
                else:
                    return (
                        _terbilang(n // 1000000000)
                        + " Miliar "
                        + _terbilang(n % 1000000000)
                    )
            else:
                if n % 1000000000000 == 0:
                    return _terbilang(n // 1000000000000) + " Triliun"
                else:
                    return (
                        _terbilang(n // 1000000000000)
                        + " Triliun "
                        + _terbilang(n % 1000000000000)
                    )

        return _terbilang(angka) + " Rupiah"

    def get_gaji_berdasarkan_golongan_dan_masa_kerja(
        self, golongan: str, masa_kerja_tahun: int, status: str
    ) -> int:
        """
        Get gaji baru from Daftar Gaji {status}.xlsx based on golongan and masa kerja tahun baru.

        Args:
            golongan (str): Golongan (e.g., "III/a", "IV/b")
            masa_kerja_tahun (int): Masa kerja tahun baru

        Returns:
            int: Gaji baru value
        """
        try:
            # TO DO: Pisahkan antara PNS dan PPPK
            if status == "PNS":
                df = pd.read_excel(config.FILE_DAFTAR_GAJI_PNS)
            elif status == "PPPK":
                df = pd.read_excel(config.FILE_DAFTAR_GAJI_PPPK)
            else:
                return 0
            # Filter by golongan and masa kerja tahun
            matching_rows = df[
                (df["GOLONGAN"] == golongan) & (
                    df["MK TAHUN"] == masa_kerja_tahun)
            ]

            if not matching_rows.empty:
                # Get the first matching row's GAJI value
                gaji = matching_rows.iloc[0]["GAJI"]
                return int(gaji) if pd.notna(gaji) else 0
            else:
                print(
                    f"Error: Gaji tidak ditemukan untuk golongan {golongan}, MK TAHUN {masa_kerja_tahun}"
                )
                print(
                    "Silakan periksa file Daftar Gaji.xlsx dan pastikan data tersedia"
                )
                print("Mengembalikan gaji 0 sebagai fallback")
                return 0

        except FileNotFoundError:
            print("Error: File Daftar Gaji.xlsx tidak ditemukan")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            print(
                "Silakan pastikan file Daftar Gaji.xlsx ada di folder yang sama dengan script"
            )
            sys.exit(1)
        except PermissionError:
            print(f"Error: File Daftar Gaji.xlsx sedang dibuka di aplikasi lain")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            print(
                "Silakan tutup file Excel yang sedang dibuka dan jalankan kembali script"
            )
            sys.exit(1)
        except Exception as e:
            error_msg = str(e)
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            if "Permission denied" in error_msg:
                print(f"Error: File Daftar Gaji.xlsx sedang dibuka di aplikasi lain")
                print(
                    "Silakan tutup file Excel yang sedang dibuka dan jalankan kembali script"
                )
            else:
                print(f"Error: Gagal membaca file Daftar Gaji.xlsx: {e}")
                print(
                    "Silakan periksa format file dan pastikan kolom yang diperlukan tersedia"
                )
            sys.exit(1)

    def download_document(
        self, arsip_path: str, employee_name: str, doc_type: str, index: int
    ) -> str:
        """
        Download document from arsip path and save to organized folder structure.

        Args:
            arsip_path (str): Path from API response
            employee_name (str): Employee name for folder naming
            doc_type (str): Document type ('Kenaikan Pangkat' or 'Kenaikan Gaji Berkala')
            index (int): Index for maintaining Excel order

        Returns:
            str: Path to downloaded file or error message
        """
        try:
            if not arsip_path or arsip_path.strip() == "":
                error_msg = f"No arsip path provided for {doc_type}"
                print(f"     {error_msg}")
                return error_msg

            # Create base documents folder
            base_docs_folder = config.DOCUMENTS_FOLDER
            os.makedirs(base_docs_folder, exist_ok=True)

            # Construct download URL
            encoded_path = urllib.parse.quote(arsip_path)
            download_url = (
                f"{self.base_url}{config.GET_ARSIP_CUTI_PATH}?arsip={encoded_path}"
            )

            # Download file with detailed error handling
            response = requests.get(
                download_url, headers=self.headers, timeout=config.TIMEOUT_SECONDS
            )

            if response.status_code != 200:
                error_msg = f"HTTP {response.status_code}: {response.text[:200]}"
                print(f"     {error_msg}")
                return error_msg

            # Check if response contains actual content
            content_type = response.headers.get("content-type", "")
            content_length = len(response.content)

            if content_length == 0:
                error_msg = f"Empty response from server"
                print(f"     {error_msg}")
                return error_msg

            # Determine filename
            # Clean employee name for folder
            clean_name = "".join(
                c for c in employee_name if c.isalnum() or c in (" ", "-", "_")
            ).strip()
            filename = f"{index:03d}_{clean_name}_{doc_type}.pdf"
            file_path = os.path.join(base_docs_folder, filename)

            # Save file
            with open(file_path, "wb") as f:
                f.write(response.content)

            # Verify file was created
            if os.path.exists(file_path):
                file_size = os.path.getsize(file_path)
                return file_path
            else:
                error_msg = f"File was not created: {file_path}"
                print(f"     {error_msg}")
                return error_msg

        except requests.exceptions.Timeout:
            error_msg = f"Timeout downloading {doc_type} for {employee_name}"
            print(f"     {error_msg}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return error_msg
        except requests.exceptions.ConnectionError:
            error_msg = (
                f"Connection error downloading {doc_type} for {employee_name}"
            )
            print(f"     {error_msg}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return error_msg
        except requests.exceptions.RequestException as e:
            error_msg = (
                f"Request error downloading {doc_type} for {employee_name}: {e}"
            )
            print(f"     {error_msg}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return error_msg
        except OSError as e:
            error_msg = f"File system error for {doc_type}: {e}"
            print(f"     {error_msg}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return error_msg
        except Exception as e:
            error_msg = (
                f"Unexpected error downloading {doc_type} for {employee_name}: {e}"
            )
            print(f"     {error_msg}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            import traceback

            print(f"     Traceback: {traceback.format_exc()}")
            return error_msg

    def download_employee_documents(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Download documents for all employees in the DataFrame.

        Args:
            df (pd.DataFrame): KGB data with document paths

        Returns:
            pd.DataFrame: Updated DataFrame with download status
        """
        print("\n" + "=" * 40)
        print("DOWNLOADING DOCUMENTS - STARTING")
        print("=" * 40)
        print(f"Total employees to process: {len(df)}")

        if df.empty:
            print("No employees to process!")
            return pd.DataFrame()

        download_results = []

        try:
            for i, row in enumerate(df.itertuples(index=False), 1):
                print(f"Processing documents {i}/{len(df)}: {row.nama}")

                # Use arsip paths that are already in the DataFrame
                arsip_pangkat = getattr(row, "arsip_pangkat", "")
                arsip_gaji = getattr(row, "arsip_gaji", "")

                # Download documents if paths exist
                pangkat_path = ""
                gaji_path = ""

                if arsip_pangkat:
                    pangkat_path = self.download_document(
                        arsip_pangkat, row.nama, "Kenaikan Pangkat", i
                    )

                if arsip_gaji:
                    gaji_path = self.download_document(
                        arsip_gaji, row.nama, "Kenaikan Gaji Berkala", i
                    )

                download_results.append(
                    {
                        "nama": row.nama,
                        "pangkat_download": pangkat_path,
                        "gaji_download": gaji_path,
                    }
                )

        except Exception as e:
            print(f"ERROR in download_employee_documents: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            import traceback

            traceback.print_exc()

        # Create download summary DataFrame
        download_df = pd.DataFrame(download_results)

        print(f"\nDocument download completed: {len(df)} employees processed")
        print(f"Download results: {len(download_df)} records")

        return download_df

    def get_pejabat_penandatangan(
        self, golongan: str, nip: str = None, status: str = None
    ) -> Dict:
        """
        Get pejabat penandatangan based on golongan from pejabat_penandatangan.json file.
        Special case handling for specific NIPs.

        Args:
            golongan (str): Golongan (e.g., "III/a", "IV/b")
            nip (str): NIP for special case matching (optional)

        Returns:
            Dict: Dictionary containing jabatan, nama, and nip
        """
        golongan = golongan.split("/")[0]
        import json

        with open(config.FILE_PEJABAT_PENANDATANGAN, "r", encoding="utf-8") as f:
            pejabat_penandatangan_list = json.load(f)

        with open(config.FILE_PEJABAT, "r", encoding="utf-8") as f:
            pejabat_list = json.load(f)

        for pejabat in pejabat_list["pejabat_data"]:
            if nip == pejabat["nip"]:
                for pejabat_penandatangan in pejabat_penandatangan_list[
                    "pejabat_penandatangan_data"
                ]:
                    if pejabat_penandatangan["special_case"] == True:
                        return pejabat_penandatangan

        for pejabat_penandatangan in pejabat_penandatangan_list[
            "pejabat_penandatangan_data"
        ]:
            if status == "PNS":
                if golongan in pejabat_penandatangan["golongan_pns"]:
                    return pejabat_penandatangan
            elif status == "PPPK":
                if golongan in pejabat_penandatangan["golongan_pppk"]:
                    return pejabat_penandatangan

        # Default fallback if no match found
        return {
            "jabatan": "Pejabat Default",
            "nama": "Nama Default",
            "nip": "000000000000000000",
        }

    def get_salutation(self, nip: str) -> str:
        """
        Get gender-based salutation based on NIP (4th digit from end).

        Args:
            nip (str): Employee NIP

        Returns:
            str: "Sdr." for male (1) or "Sdri." for female (2)
        """
        try:
            if len(nip) >= 4:
                gender_digit = nip[-4]
                return "Sdr." if gender_digit == "1" else "Sdri."
            return "Sdr."
        except:
            print(f"Error getting salutation: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return "Sdr."

    def get_gaji_pokok_data(self, employee_id: str) -> Dict:
        """
        Fetch latest gaji pokok data for employee.

        Args:
            employee_id (str): Employee ID

        Returns:
            Dict: Latest gaji pokok data or empty dict
        """
        try:
            payload = {
                "on_page": 1,
                "page_size": config.HISTORY_PAGE_SIZE,
                "id_pegawai": employee_id,
            }

            response = requests.post(
                f"{self.base_url}{config.GAJI_POKOK_GET_PATH}",
                json=payload,
                headers=self.headers,
            )

            if response.status_code == 200:
                data = response.json()
                if "history_gaji_pokok" in data and len(data["history_gaji_pokok"]) > 0:
                    return data["history_gaji_pokok"][0]  # Latest is first
            return {}

        except Exception as e:
            print(f"Error fetching gaji pokok data for {employee_id}: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return {}

    def get_pangkat_data(self, employee_id: str) -> Dict:
        """
        Fetch latest pangkat data for employee.

        Args:
            employee_id (str): Employee ID

        Returns:
            Dict: Latest pangkat data or empty dict
        """
        try:
            payload = {
                "on_page": 1,
                "page_size": config.HISTORY_PAGE_SIZE,
                "id_pegawai": employee_id,
            }

            response = requests.post(
                f"{self.base_url}{config.PANGKAT_GET_PATH}",
                json=payload,
                headers=self.headers,
            )

            if response.status_code == 200:
                data = response.json()
                if "history_pangkat" in data and len(data["history_pangkat"]) > 0:
                    return data["history_pangkat"][0]  # Latest is first
            return {}

        except Exception as e:
            print(f"Error fetching pangkat data for {employee_id}: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return {}

    def get_gaji_pokok_detail(self, gaji_id: str) -> Dict:
        """
        Fetch detailed gaji pokok data using detail API.

        Args:
            gaji_id (str): Gaji pokok ID from previous API call

        Returns:
            Dict: Detailed gaji pokok data or empty dict
        """
        try:
            url = f"{self.base_url}{config.GAJI_POKOK_GETDETAIL_PATH}/{gaji_id}"
            response = requests.post(url, headers=self.headers)

            if response.status_code == 200:
                data = response.json()
                if "gajiPokok" in data:
                    return data["gajiPokok"]
            return {}

        except Exception as e:
            print(f"Error fetching gaji pokok detail for {gaji_id}: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return {}

    def get_pangkat_detail(self, pangkat_id: str) -> Dict:
        """
        Fetch detailed pangkat data using detail API.

        Args:
            pangkat_id (str): Pangkat ID from previous API call

        Returns:
            Dict: Detailed pangkat data or empty dict
        """
        try:
            url = f"{self.base_url}{config.PANGKAT_GETDETAIL_PATH}/{pangkat_id}"

            response = requests.post(url, headers=self.headers)

            if response.status_code == 200:
                data = response.json()
                if "pangkat" in data:
                    return data["pangkat"]
            return {}

        except Exception as e:
            print(f"Error fetching pangkat detail for {pangkat_id}: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            return {}

    def get_tanggal_surat_bulan_tahun(self) -> str:
        """
        Get current date in Indonesian format (bulan tahun only).

        Returns:
            str: Date in format "Januari 2026"
        """
        bulan_dict = {
            1: "Januari",
            2: "Februari",
            3: "Maret",
            4: "April",
            5: "Mei",
            6: "Juni",
            7: "Juli",
            8: "Agustus",
            9: "September",
            10: "Oktober",
            11: "November",
            12: "Desember",
        }

        today = datetime.now()
        bulan = bulan_dict[today.month]
        tahun = today.year

        return f"{bulan} {tahun}"

    def process_kgb_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Process KGB data by fetching additional information for eligible employees.

        Args:
            df (pd.DataFrame): Filtered employee data from Phase 2

        Returns:
            pd.DataFrame: Enhanced KGB data with additional information
        """
        print("\n" + "=" * 40)
        print("PHASE 3: KGB DATA PROCESSING")
        print("=" * 40)

        kgb_data = []
        total_employees = len(df)

        for i, row in enumerate(df.itertuples(index=False), 1):
            print(f"Processing {i}/{total_employees}: {row.nama}")

            # Initialize all acuan terakhir variables with default values
            acuan_terakhir = ""
            acuan_terakhir_masa_kerja_tahun = 0
            acuan_terakhir_masa_kerja_bulan = 0
            acuan_terakhir_tanggal_mulai = ""
            acuan_terakhir_tanggal_sk = ""
            acuan_terakhir_gaji_terbilang = ""
            acuan_terakhir_no_sk = ""
            acuan_terakhir_gaji = 0
            acuan_terakhir_pejabat_penandatangan = ""

            # Initialize all gaji pokok variables with default values
            gaji_pokok_latest_id = ""
            gaji_pokok_terakhir_gaji = 0
            gaji_pokok_terakhir_gaji_terbilang = ""
            gaji_pokok_terakhir_pejabat_penandatangan = ""
            gaji_pokok_terakhir_tanggal_sk = ""
            gaji_pokok_terakhir_tanggal_berlaku = ""
            gaji_pokok_terakhir_masa_kerja_tahun = 0
            gaji_pokok_terakhir_masa_kerja_bulan = 0
            gaji_pokok_terakhir_no_sk = ""

            # Initialize all pangkat variables with default values
            pangkat_latest_id = ""
            pangkat_terakhir_gaji = 0
            pangkat_terakhir_gaji_terbilang = ""
            pangkat_terakhir_no_sk = ""
            pangkat_terakhir_tanggal_sk = ""
            pangkat_terakhir_tanggal_mulai = ""
            pangkat_terakhir_masa_kerja_tahun = 0
            pangkat_terakhir_masa_kerja_bulan = 0

            # Fetch gaji pokok data
            gaji_pokok = self.get_gaji_pokok_data(row.id)
            if gaji_pokok:
                gaji_pokok_latest_id = self.clean_gaji_amount(
                    str(gaji_pokok.get("id", ""))
                )
                gaji_pokok_terakhir_gaji = gaji_pokok.get("gaji_pokok", "")
                gaji_pokok_terakhir_gaji_terbilang = self.terbilang_indonesia(
                    gaji_pokok_terakhir_gaji
                )
                gaji_pokok_terakhir_no_sk = gaji_pokok.get("no_sk", "")
                gaji_pokok_terakhir_tanggal_berlaku = gaji_pokok.get(
                    "tanggal_berlaku", ""
                )
                gaji_pokok_terakhir_tanggal_berlaku = (
                    gaji_pokok_terakhir_tanggal_berlaku.split("T")[0]
                    if gaji_pokok_terakhir_tanggal_berlaku
                    and "T" in gaji_pokok_terakhir_tanggal_berlaku
                    else gaji_pokok_terakhir_tanggal_berlaku
                )
                gaji_pokok_latest_detail_data = self.get_gaji_pokok_detail(
                    gaji_pokok_latest_id
                )
                if gaji_pokok_latest_detail_data:
                    gaji_pokok_terakhir_pejabat_penandatangan = (
                        gaji_pokok_latest_detail_data.get("pejabat", "")
                    )
                    gaji_pokok_terakhir_tanggal_sk = gaji_pokok_latest_detail_data.get(
                        "tanggal_sk", ""
                    )
                    gaji_pokok_terakhir_tanggal_sk = (
                        gaji_pokok_terakhir_tanggal_sk.split("T")[0]
                        if gaji_pokok_terakhir_tanggal_sk
                        and "T" in gaji_pokok_terakhir_tanggal_sk
                        else gaji_pokok_terakhir_tanggal_sk or ""
                    )
                    gaji_pokok_terakhir_masa_kerja_tahun = (
                        gaji_pokok_latest_detail_data.get("tahun_baru", "")
                    )
                    gaji_pokok_terakhir_masa_kerja_bulan = (
                        gaji_pokok_latest_detail_data.get("bulan_baru", "")
                    )
                    arsip_gaji = gaji_pokok_latest_detail_data.get("arsip", "")

            # Fetch pangkat data
            pangkat = self.get_pangkat_data(row.id)
            if pangkat:
                pangkat_latest_id = str(pangkat.get("id", ""))
                pangkat_terakhir_no_sk = pangkat.get("no_sk", "")
                pangkat_terakhir_tanggal_mulai = pangkat.get(
                    "tanggal_mulai", "")
                pangkat_terakhir_tanggal_mulai = (
                    pangkat_terakhir_tanggal_mulai.split("T")[0]
                    if pangkat_terakhir_tanggal_mulai
                    and "T" in pangkat_terakhir_tanggal_mulai
                    else pangkat_terakhir_tanggal_mulai
                )
                pangkat_terakhir_masa_kerja_tahun = pangkat.get(
                    "masa_kerja_tahun", "")
                pangkat_terakhir_masa_kerja_bulan = pangkat.get(
                    "masa_kerja_bulan", "")
                pangkat_latest_detail_data = self.get_pangkat_detail(
                    pangkat_latest_id)
                if pangkat_latest_detail_data:
                    pangkat_terakhir_gaji = self.clean_gaji_amount(
                        pangkat_latest_detail_data.get("gaji", "")
                    )
                    pangkat_terakhir_gaji_terbilang = self.terbilang_indonesia(
                        pangkat_terakhir_gaji
                    )
                    pangkat_terakhir_tanggal_sk = pangkat_latest_detail_data.get(
                        "tanggal_sk", ""
                    )
                    pangkat_terakhir_tanggal_sk = (
                        pangkat_terakhir_tanggal_sk.split("T")[0]
                        if pangkat_terakhir_tanggal_sk
                        and "T" in pangkat_terakhir_tanggal_sk
                        else pangkat_terakhir_tanggal_sk
                    )
                    arsip_pangkat = pangkat_latest_detail_data.get("arsip", "")

            # Determine acuan terakhir (latest reference) between gaji pokok and pangkat
            if gaji_pokok_terakhir_tanggal_berlaku and pangkat_terakhir_tanggal_mulai:
                gaji_date = (
                    gaji_pokok_terakhir_tanggal_berlaku.split("T")[0]
                    if "T" in gaji_pokok_terakhir_tanggal_berlaku
                    else gaji_pokok_terakhir_tanggal_berlaku
                )
                pangkat_date_clean = (
                    pangkat_terakhir_tanggal_mulai.split("T")[0]
                    if "T" in pangkat_terakhir_tanggal_mulai
                    else pangkat_terakhir_tanggal_mulai
                )

                if gaji_date > pangkat_date_clean:
                    acuan_terakhir = f"Gaji Pokok ({gaji_date})"
                    acuan_terakhir_masa_kerja_tahun = (
                        gaji_pokok_terakhir_masa_kerja_tahun
                    )
                    acuan_terakhir_masa_kerja_bulan = (
                        gaji_pokok_terakhir_masa_kerja_bulan
                    )
                    acuan_terakhir_tanggal_mulai = gaji_pokok_terakhir_tanggal_berlaku
                    acuan_terakhir_tanggal_sk = gaji_pokok_terakhir_tanggal_sk
                    acuan_terakhir_gaji_terbilang = gaji_pokok_terakhir_gaji_terbilang
                    acuan_terakhir_no_sk = gaji_pokok_terakhir_no_sk
                    acuan_terakhir_gaji = gaji_pokok_terakhir_gaji
                    acuan_terakhir_pejabat_penandatangan = (
                        gaji_pokok_terakhir_pejabat_penandatangan
                    )

                elif pangkat_date_clean > gaji_date:
                    acuan_terakhir = f"Pangkat ({pangkat_date_clean})"
                    acuan_terakhir_masa_kerja_tahun = pangkat_terakhir_masa_kerja_tahun
                    acuan_terakhir_masa_kerja_bulan = pangkat_terakhir_masa_kerja_bulan
                    acuan_terakhir_tanggal_mulai = pangkat_terakhir_tanggal_mulai
                    acuan_terakhir_tanggal_sk = pangkat_terakhir_tanggal_sk
                    acuan_terakhir_gaji_terbilang = pangkat_terakhir_gaji_terbilang

                    acuan_terakhir_no_sk = pangkat_terakhir_no_sk
                    acuan_terakhir_gaji = pangkat_terakhir_gaji
                    acuan_terakhir_pejabat_penandatangan = ""
                else:
                    # When dates are equal, prioritize Kenaikan Pangkat
                    acuan_terakhir = f"Gaji Pokok - Pangkat ({pangkat_date_clean})"
                    acuan_terakhir_masa_kerja_tahun = (
                        gaji_pokok_terakhir_masa_kerja_tahun
                    )
                    acuan_terakhir_masa_kerja_bulan = (
                        gaji_pokok_terakhir_masa_kerja_bulan
                    )
                    acuan_terakhir_tanggal_mulai = gaji_pokok_terakhir_tanggal_berlaku
                    acuan_terakhir_tanggal_sk = gaji_pokok_terakhir_tanggal_sk
                    acuan_terakhir_gaji_terbilang = gaji_pokok_terakhir_gaji_terbilang
                    acuan_terakhir_no_sk = gaji_pokok_terakhir_no_sk
                    acuan_terakhir_gaji = gaji_pokok_terakhir_gaji
                    acuan_terakhir_pejabat_penandatangan = (
                        gaji_pokok_terakhir_pejabat_penandatangan
                    )

            # Get pejabat penandatangan based on golongan and status
            pejabat_data = self.get_pejabat_penandatangan(
                row.golongan, row.nip, row.status
            )

            # Get tembusan based on unit_kerja
            tembusan = self.get_tembusan_by_unit_kerja(row.unit_kerja)

            # Get gaji baru based on golongan and masa kerja tahun baru
            baru_masa_kerja_tahun = row.masa_kerja_tahun + 1
            # Penyesuaian masa kerja PPPK
            if row.status == "PPPK" and row.golongan in [
                "II",
                "III",
                "IV",
                "VI",
                "VII",
                "VIII",
            ]:
                baru_masa_kerja_tahun += 3
                acuan_terakhir_masa_kerja_tahun += 3

            baru_masa_kerja_bulan = 0
            baru_gaji = self.get_gaji_berdasarkan_golongan_dan_masa_kerja(
                row.golongan, baru_masa_kerja_tahun, row.status
            )
            baru_gaji_terbilang = self.terbilang_indonesia(baru_gaji)

            # Safe conversion to int with None check and fallback
            baru_masa_kerja_tahun_clean = int(
                (str(baru_masa_kerja_tahun or 0).strip() or "0")
            )
            acuan_terakhir_masa_kerja_tahun_clean = int(
                (str(acuan_terakhir_masa_kerja_tahun or 0).strip() or "0")
            )
            acuan_terakhir_selisih_masa_kerja_tahun = (
                baru_masa_kerja_tahun_clean - acuan_terakhir_masa_kerja_tahun_clean
            )

            kgb_employee = {
                "id": row.id,
                "gaji_pokok_id": gaji_pokok_latest_id,
                "pangkat_id": pangkat_latest_id,
                "tanggal_surat": self.get_tanggal_surat_bulan_tahun(),
                "nomor_surat": pejabat_data["nomor_surat_pns"] if row.status == "PNS" else pejabat_data["nomor_surat_pppk"],
                "panggilan": self.get_salutation(row.nip),
                "nama": row.nama,
                "nip": row.nip,
                "pangkat": row.pangkat,
                "golongan": row.golongan,
                "jabatan": row.jabatan,
                "unit_kerja": row.unit_kerja,
                "tembusan": tembusan,
                "masa_kerja_tahun": row.masa_kerja_tahun,
                "masa_kerja_bulan": row.masa_kerja_bulan,
                "gaji_pokok_terakhir_gaji": gaji_pokok_terakhir_gaji,
                "gaji_pokok_terakhir_gaji_terbilang": gaji_pokok_terakhir_gaji_terbilang,
                "gaji_pokok_terakhir_pejabat_penandatangan": gaji_pokok_terakhir_pejabat_penandatangan,
                "gaji_pokok_terakhir_no_sk": gaji_pokok_terakhir_no_sk,
                "gaji_pokok_terakhir_tanggal_sk": gaji_pokok_terakhir_tanggal_sk,
                "gaji_pokok_terakhir_tanggal_berlaku": gaji_pokok_terakhir_tanggal_berlaku,
                "gaji_pokok_terakhir_masa_kerja_tahun": gaji_pokok_terakhir_masa_kerja_tahun,
                "gaji_pokok_terakhir_masa_kerja_bulan": gaji_pokok_terakhir_masa_kerja_bulan,
                "pangkat_terakhir_gaji": pangkat_terakhir_gaji,
                "pangkat_terakhir_gaji_terbilang": pangkat_terakhir_gaji_terbilang,
                "pangkat_terakhir_no_sk": pangkat_terakhir_no_sk,
                "pangkat_terakhir_tanggal_sk": pangkat_terakhir_tanggal_sk,
                "pangkat_terakhir_tanggal_mulai": pangkat_terakhir_tanggal_mulai,
                "pangkat_terakhir_masa_kerja_tahun": pangkat_terakhir_masa_kerja_tahun,
                "pangkat_terakhir_masa_kerja_bulan": pangkat_terakhir_masa_kerja_bulan,
                "baru_gaji": baru_gaji,
                "baru_gaji_terbilang": baru_gaji_terbilang,
                "baru_masa_kerja_tahun": baru_masa_kerja_tahun,
                "baru_masa_kerja_bulan": baru_masa_kerja_bulan,
                "baru_mulai_tanggal": (datetime.now() + relativedelta(months=2))
                .replace(day=1)
                .strftime("%Y-%m-%d"),
                "baru_jabatan_pejabat_penandatangan": pejabat_data["jabatan"],
                "baru_nama_pejabat_penandatangan": pejabat_data["nama"],
                "baru_nip_pejabat_penandatangan": pejabat_data["nip"],
                "status": row.status,
                "acuan_terakhir_masa_kerja_tahun": acuan_terakhir_masa_kerja_tahun,
                "acuan_terakhir_masa_kerja_bulan": acuan_terakhir_masa_kerja_bulan,
                "acuan_terakhir_tanggal_mulai": acuan_terakhir_tanggal_mulai,
                "acuan_terakhir_tanggal_sk": acuan_terakhir_tanggal_sk,
                "acuan_terakhir_gaji_terbilang": acuan_terakhir_gaji_terbilang,
                "acuan_terakhir": acuan_terakhir,
                "acuan_terakhir_no_sk": acuan_terakhir_no_sk,
                "acuan_terakhir_gaji": acuan_terakhir_gaji,
                "acuan_terakhir_pejabat_penandatangan": acuan_terakhir_pejabat_penandatangan,
                "acuan_terakhir_selisih_masa_kerja_tahun": acuan_terakhir_selisih_masa_kerja_tahun,
                "arsip_gaji": arsip_gaji,
                "arsip_pangkat": arsip_pangkat,
                "keterangan": "",
            }

            kgb_data.append(kgb_employee)

        return pd.DataFrame(kgb_data)

    def kgb_data_processing(self, input_df: pd.DataFrame) -> str:
        """
        Execute Phase 3: KGB Data Processing.

        Args:
            input_df (pd.DataFrame): Input data from Phase 2

        Returns:
            str: Path to generated Excel file
        """
        print("=" * 60)
        print("KGB INQUIRY - PHASE 3: KGB DATA PROCESSING")
        print("=" * 60)

        # Process KGB data
        print("Step 1: Processing KGB data...")
        kgb_df = self.process_kgb_data(input_df)

        if kgb_df.empty:
            print("No KGB data to process")
            sys.exit(1)

        print(f"\nKGB processing completed: {len(kgb_df)} employees processed")

        # Download documents
        print("\nStep 2: Downloading documents...")

        try:
            download_df = self.download_employee_documents(kgb_df)
            print("download_employee_documents completed successfully")
        except Exception as e:
            print(f"Error in download_employee_documents: {e}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            import traceback

            traceback.print_exc()
            download_df = pd.DataFrame()

        # Save to Excel with KGB-specific format
        print("\nStep 3: Saving Excel files...")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{config.OUTPUT_FILE_PREFIX}{timestamp}.xlsx"
        output_file = self.save_kgb_excel(kgb_df, filename)

        print("\n" + "=" * 60)
        print("PHASE 3 COMPLETED")
        print("=" * 60)
        print(f"Output file: {output_file}")
        print(f"Documents folder: {config.DOCUMENTS_FOLDER}/")

        return output_file

    def get_unit_kerja_order(self) -> List[str]:
        """
        Load unit kerja ordering from urutan_unit_kerja.json file.

        Returns:
            List[str]: Ordered list of unit kerja names
        """
        try:
            with open(config.FILE_URUTAN_UNIT_KERJA, "r", encoding="utf-8") as f:
                unit_kerja_config = json.load(f)
                return unit_kerja_config.get("unit_kerja_order", [])
        except FileNotFoundError:
            print("Warning: urutan_unit_kerja.json not found")
            print("Using alphabetical order for unit_kerja")
            return []
        except json.JSONDecodeError as e:
            print(f"Warning: Error reading urutan_unit_kerja.json: {e}")
            return []

    def get_tembusan_by_unit_kerja(self, unit_kerja: str) -> str:
        """
        Get tembusan (cc recipient) based on unit_kerja from tembusan_mapping.json file.

        Args:
            unit_kerja (str): Unit kerja name

        Returns:
            str: Tembusan name or empty string if not found
        """
        try:
            with open(config.FILE_TEMBUSAN_MAPPING, "r", encoding="utf-8") as f:
                tembusan_config = json.load(f)
                tembusan_mapping = tembusan_config.get("tembusan_mapping", {})
                return tembusan_mapping.get(unit_kerja, "")
        except FileNotFoundError:
            print("Warning: tembusan_mapping.json not found")
            return ""
        except json.JSONDecodeError as e:
            print(f"Warning: Error reading tembusan_mapping.json: {e}")
            return ""

    def save_kgb_excel(self, df: pd.DataFrame, filename: str) -> str:
        """
        Save KGB processed data to Excel file with specific format.
        Data is separated by status (PNS and PPPK) and sorted accordingly.

        Args:
            df (pd.DataFrame): KGB data to save
            filename (str): Output filename

        Returns:
            str: Path to saved file
        """
        try:
            # Define correct column order with new KGB fields
            column_order = [
                "nama",
                "nip",
                "pangkat",
                "golongan",
                "jabatan",
                "unit_kerja",
                "id",
                "masa_kerja_tahun",
                "masa_kerja_bulan",
                "tanggal_surat",
                "nomor_surat",
                "panggilan",
                "tembusan",
                "gaji_pokok_terakhir_gaji",
                "gaji_pokok_terakhir_gaji_terbilang",
                "gaji_pokok_terakhir_pejabat_penandatangan",
                "gaji_pokok_terakhir_no_sk",
                "gaji_pokok_terakhir_tanggal_sk",
                "gaji_pokok_terakhir_tanggal_berlaku",
                "gaji_pokok_terakhir_masa_kerja_tahun",
                "gaji_pokok_terakhir_masa_kerja_bulan",
                "pangkat_terakhir_gaji",
                "pangkat_terakhir_gaji_terbilang",
                "pangkat_terakhir_no_sk",
                "pangkat_terakhir_tanggal_sk",
                "pangkat_terakhir_tanggal_mulai",
                "pangkat_terakhir_masa_kerja_tahun",
                "pangkat_terakhir_masa_kerja_bulan",
                "baru_gaji",
                "baru_gaji_terbilang",
                "baru_masa_kerja_tahun",
                "baru_masa_kerja_bulan",
                "baru_mulai_tanggal",
                "baru_jabatan_pejabat_penandatangan",
                "baru_nama_pejabat_penandatangan",
                "baru_nip_pejabat_penandatangan",
                "status",
                "acuan_terakhir_masa_kerja_tahun",
                "acuan_terakhir_masa_kerja_bulan",
                "acuan_terakhir_tanggal_mulai",
                "acuan_terakhir_tanggal_sk",
                "acuan_terakhir_gaji_terbilang",
                "acuan_terakhir_gaji",
                "acuan_terakhir",
                "acuan_terakhir_no_sk",
                "acuan_terakhir_pejabat_penandatangan",
                "acuan_terakhir_selisih_masa_kerja_tahun",
                "keterangan",
            ]

            # Reorder DataFrame columns
            df_ordered = df[column_order].copy()

            # Load unit kerja ordering from JSON file
            unit_kerja_order = self.get_unit_kerja_order()

            # Create a mapping for custom unit_kerja sorting
            if unit_kerja_order:
                unit_kerja_mapping = {
                    uk: idx for idx, uk in enumerate(unit_kerja_order)
                }
            else:
                # Fallback: use alphabetical order if JSON not available
                unique_units = sorted(df_ordered["unit_kerja"].unique())
                unit_kerja_mapping = {uk: idx for idx,
                                      uk in enumerate(unique_units)}

            with pd.ExcelWriter(filename, engine="openpyxl") as writer:
                # Utility from openpyxl used for column sizing
                from openpyxl.utils import get_column_letter

                # Separate data by status and save to different sheets
                for status in ["PNS", "PPPK"]:
                    status_df = df_ordered[df_ordered["status"]
                                           == status].copy()

                    if not status_df.empty:
                        # Add sorting keys for custom unit_kerja order
                        status_df["unit_kerja_sort_key"] = status_df["unit_kerja"].map(
                            lambda x: unit_kerja_mapping.get(
                                x, len(unit_kerja_mapping))
                        )

                        # Sort by: baru_nama_pejabat_penandatangan (A-Z), unit_kerja (custom order),
                        # baru_masa_kerja_tahun (high-low), baru_masa_kerja_bulan (high-low), nama (A-Z)
                        status_df = status_df.sort_values(
                            by=[
                                "baru_nama_pejabat_penandatangan",
                                "unit_kerja_sort_key",
                                "baru_masa_kerja_tahun",
                                "baru_masa_kerja_bulan",
                                "nama",
                            ],
                            ascending=[True, True, False, False, True],
                        )

                        # Remove temporary sorting column
                        status_df = status_df.drop(
                            "unit_kerja_sort_key", axis=1)

                        # Reset index for cleaner output
                        status_df = status_df.reset_index(drop=True)

                        # Save to sheet named after status
                        status_df.to_excel(
                            writer, sheet_name=status, index=False)
                        print(f"Sheet '{status}': {len(status_df)} records")

                # Create summary
                df_comb = df.copy()
                df_comb["status"] = df_comb["status"].fillna("")
                df_comb["baru_nama_pejabat_penandatangan"] = df_comb["baru_nama_pejabat_penandatangan"].fillna(
                    "")

                summary_combined = (
                    df_comb.groupby(
                        ["unit_kerja", "status", "baru_nama_pejabat_penandatangan", "baru_masa_kerja_tahun", "baru_masa_kerja_bulan"]).size()
                    .reset_index(name="Jumlah_Pegawai")
                )

                # Format Masa Kerja column
                summary_combined["Masa Kerja"] = summary_combined.apply(
                    lambda r: f"{int(r['baru_masa_kerja_tahun'])} Tahun {int(r['baru_masa_kerja_bulan'])} Bulan",
                    axis=1,
                )

                # Add sorting key and sort by unit order, unit name, status, then pejabat
                summary_combined["unit_kerja_sort_key"] = summary_combined["unit_kerja"].map(
                    lambda x: unit_kerja_mapping.get(
                        x, len(unit_kerja_mapping))
                )
                summary_combined = summary_combined.sort_values(
                    by=[
                        "unit_kerja_sort_key",
                        "unit_kerja",
                        "status",
                        "baru_nama_pejabat_penandatangan",
                        "baru_masa_kerja_tahun",
                        "baru_masa_kerja_bulan",
                    ],
                    ascending=[True, True, True, True, False, False],
                ).reset_index(drop=True)
                summary_combined.insert(0, "No", summary_combined.index + 1)

                total_people_combined = int(
                    summary_combined["Jumlah_Pegawai"].sum()) if not summary_combined.empty else 0
                header_combined = f"{total_people_combined} orang - Rincian per Unit Kerja berdasarkan Status dan Pejabat Penandatangan"

                sheet_comb_name = "Summary_by_Unit_Status_Pejabat"
                # Ensure sheet name complies with Excel limit (<= 31 chars)
                if len(sheet_comb_name) > 31:
                    sheet_comb_name = sheet_comb_name[:31]
                    print(
                        f"Warning: sheet name truncated to '{sheet_comb_name}' to comply with Excel limits")
                if sheet_comb_name in writer.book.sheetnames:
                    idx = writer.book.sheetnames.index(sheet_comb_name)
                    existing = writer.book.worksheets[idx]
                    writer.book.remove(existing)

                ws_comb = writer.book.create_sheet(sheet_comb_name)
                ws_comb.merge_cells(
                    start_row=1, start_column=1, end_row=1, end_column=6)
                ws_comb["A1"] = header_combined
                ws_comb.append(["No", "Unit Kerja", "Status",
                               "Pejabat Penandatangan", "Jumlah Pegawai", "Masa Kerja"])
                for _, r in summary_combined.iterrows():
                    ws_comb.append([
                        int(r["No"]),
                        r["unit_kerja"],
                        r["status"],
                        r["baru_nama_pejabat_penandatangan"],
                        int(r["Jumlah_Pegawai"]),
                        r["Masa Kerja"],
                    ])
                widths_comb = [5, 60, 15, 40, 15, 20]
                for i, w in enumerate(widths_comb, start=1):
                    ws_comb.column_dimensions[get_column_letter(i)].width = w

            # Validate workbook by attempting to load it with openpyxl
            try:
                from openpyxl import load_workbook
                _ = load_workbook(filename)
                print(f"Workbook validated: {filename}")
            except Exception as e:
                print(f"Warning: Failed to validate saved workbook: {e}")
                # Re-raise to make the failure visible
                raise

            print(f"KGB data successfully saved to {filename}")
            return filename

        except Exception as e:
            print(f"Error saving KGB Excel: {e}")
            print(f"Error type: {type(e).__name__}")
            print(
                f"Error occurred at line: {sys.exc_info()[2].tb_lineno if sys.exc_info()[2] else 'Unknown'}"
            )
            print(f"Full traceback:\n{traceback.format_exc()}")
            sys.exit(1)

    def run_all_phases(self) -> str:
        """
        Execute all phases sequentially: Phase 1, Phase 2, and Phase 3.

        Returns:
            str: Path to generated Excel file
        """
        print("=" * 60)
        print("KGB INQUIRY - COMPLETE WORKFLOW (PHASE 1 + 2 + 3)")
        print("=" * 60)

        # Phase 1: Fetch and process data
        print("\n" + "=" * 40)
        print("PHASE 1: EMPLOYEE DATA INQUIRY")
        print("=" * 40)

        api_response = self.search_employees()
        df = self.process_employee_data(api_response)

        if df.empty:
            print("No employee data to process")
            sys.exit(1)

        print(f"Phase 1 completed: {len(df)} employee records processed")

        # Phase 2: Filter data
        print("\n" + "=" * 40)
        print("PHASE 2: DATA FILTERING")
        print("=" * 40)

        # Display filtering criteria
        print("Filtering Criteria:")
        if self.status:
            print(f"  - Status: {self.status}")
        if self.tahun_kerja is not None:
            print(f"  - Tahun Kerja: {self.tahun_kerja}")
        if self.bulan_kerja is not None:
            print(f"  - Bulan Kerja: {self.bulan_kerja}")
        print("  - Unit Kerja: harus terisi")
        print("  - KGB Eligibility: Tahun kerja harus per 2 tahun")
        if not any(
            [self.status, self.tahun_kerja is not None, self.bulan_kerja is not None]
        ):
            print("  - No filters applied (showing all data)")

        # Apply filters
        filtered_df = self.filter_employees(df)
        # For debugging, limit the data
        print(f"\nOriginal records: {len(df)}")
        print(f"After unit kerja filter: {len(filtered_df)}")

        # Phase 3: KGB Data Processing with Document Download
        print("\n" + "=" * 40)
        print("PHASE 3: KGB DATA PROCESSING + DOCUMENT DOWNLOAD")
        print("=" * 40)

        # Use run_phase3 which includes document download
        output_file = self.kgb_data_processing(filtered_df)

        return output_file


def main():
    """Main function to run the KGB inquiry script."""
    parser = argparse.ArgumentParser(description="KGB Data Inquiry Script")
    parser.add_argument(
        "--token",
        help="Authentication token for API access (optional if KGB_AUTH_TOKEN env var is set)",
    )
    parser.add_argument(
        "--status",
        choices=["PNS", "PPK"],
        help="Filter by employee status (PNS or PPK)",
    )
    parser.add_argument("--tahun-kerja", type=int,
                        help="Filter by years of service")
    parser.add_argument("--bulan-kerja", type=int,
                        help="Filter by months of service")
    parser.add_argument("--nip", type=int, help="Filter by NIP")
    parser.add_argument(
        "--no-kgb-eligibility",
        action="store_true",
        help="Disable KGB eligibility filter (odd years only)",
    )
    args = parser.parse_args()

    # Get token from command line argument or config (env var with constant fallback)
    token = args.token or config.AUTH_TOKEN

    if not token:
        print(" Error: No authentication token provided!")
        print("Please provide token using either:")
        print("  1. --token argument: python kgb_inquiry.py --token your_token_here")
        print("  2. Environment variable: set KGB_AUTH_TOKEN=your_token_here")
        sys.exit(1)

    # Initialize and run unified workflow
    try:
        inquiry = KGBInquiry(
            auth_token=token,
            status=args.status,
            tahun_kerja=args.tahun_kerja,
            bulan_kerja=args.bulan_kerja,
            nip=args.nip,
            kgb_eligibility=not args.no_kgb_eligibility,
        )

        # Run unified workflow: Phase 1 + Phase 2 + Phase 3 (with document download)
        output_file = inquiry.run_all_phases()

    except KeyboardInterrupt:
        print("\n\nScript interrupted by user")
        import traceback

        print(
            f"Error occurred at line: {traceback.extract_tb(sys.exc_info()[2])[-1].lineno}"
        )
        print(f"Full traceback: {traceback.format_exc()}")
        sys.exit(1)
    except Exception as e:
        import traceback

        print(f"Unexpected error: {e}")
        print(f"Full traceback: {traceback.format_exc()}")
        sys.exit(1)


if __name__ == "__main__":
    main()
