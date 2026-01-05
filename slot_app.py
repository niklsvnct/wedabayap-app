"""
================================================================================
WEDABAY AIRPORT ABSENCE CENTER - ENTERPRISE PERSONNEL MONITORING SYSTEM
================================================================================
Version: 2.1.0 (FULL INTEGRATED BUILD)
Architecture: Clean Architecture with MVC Pattern
================================================================================
"""

import streamlit as st
import pandas as pd
from datetime import time, datetime, timedelta
import io
import xlsxwriter
import streamlit.components.v1 as components
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field
from enum import Enum
from abc import ABC, abstractmethod
import json
import hashlib
from functools import lru_cache
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import base64
import uuid
import base64 # <--- Pastikan ini ada

# Tambahkan fungsi ini untuk membaca file gambar lokal
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        return None

# ================================================================================
# SECTION 1: CONFIGURATION & CONSTANTS LAYER
# ================================================================================

class AppConstants:
    """
    Global application constants following immutable configuration pattern.
    Centralizes all magic strings and numbers for maintainability.
    """
    
    # Application Metadata
    APP_TITLE = "WedaBay Airport Absence Center"
    APP_VERSION = "2.1.0"
    APP_ICON = "✈️"
    COMPANY_NAME = "WedaBay Aviation Services"
    
    # Column Names (Database Schema)
    COL_PERSON_NAME = 'Person Name'
    COL_EVENT_TIME = 'Event Time'
    COL_EMPLOYEE_NAME = 'Nama Karyawan'
    COL_DATE = 'Tanggal'
    COL_STATUS = 'Keterangan'
    
    # Time Thresholds
    LATE_THRESHOLD = time(7, 5, 0)
    EARLY_ARRIVAL = time(6, 0, 0)
    
    # Cache Configuration
    CACHE_TTL_SECONDS = 10
    MAX_CACHE_ENTRIES = 100
    
    # UI Configuration
    SIDEBAR_STATE = "expanded"
    LAYOUT_MODE = "wide"
    CARDS_PER_ROW = 4
    
    # Data Refresh Intervals
    AUTO_REFRESH_INTERVAL = 30  # seconds
    
    # Export Settings
    EXCEL_ENGINE = 'xlsxwriter'
    DATE_FORMAT = '%Y-%m-%d'
    TIME_FORMAT = '%H:%M'
    DATETIME_FORMAT = '%Y-%m-%d %H:%M:%S'


class TimeRanges(Enum):
    MORNING = ('Pagi', '03:00:00', '9:00:00')
    BREAK_OUT = ('Siang_1', '11:29:00', '12:29:00')
    BREAK_IN = ('Siang_2', '12:31:00', '15:00:00')
    EVENING = ('Sore', '17:00:00', '23:59:59')

    def __init__(self, label: str, start: str, end: str):
        self.label = label
        self.start_time = start
        self.end_time = end

    @property
    def time_window(self) -> Tuple[str, str]:
        """Returns the time window as a tuple."""
        return (self.start_time, self.end_time)
    

class AttendanceStatus(Enum):
    """
    Enumeration for attendance status types with styling metadata.
    Follows Single Responsibility Principle for status management.
    """
    FULL_DUTY = ("FULL DUTY", "#4cd137", "status-present")
    PARTIAL_DUTY = ("PARTIAL DUTY", "#fbc531", "status-partial")
    ABSENT = ("ABSENT / NO SHOW", "#e84118", "status-absent")
    PERMIT = ("PERMIT", "#9c88ff", "status-permit")
    LATE = ("DELAYED", "#e84118", "status-late")
    
    def __init__(self, display_text: str, color: str, css_class: str):
        self.display_text = display_text
        self.color = color
        self.css_class = css_class


@dataclass
class DivisionConfig:
    """
    Data class representing a division's configuration.
    Immutable configuration following Data Transfer Object pattern.
    """
    name: str
    color: str
    icon: str
    code: str
    members: List[str] = field(default_factory=list)
    description: str = ""
    priority: int = 0
    
    def __hash__(self):
        """Make dataclass hashable for caching purposes."""
        return hash((self.name, self.code))


class DataSourceConfig:
    """
    Configuration for external data sources.
    Centralizes all external URLs and connection parameters.
    """
    
    # Google Sheets Data Sources
    ATTENDANCE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"
    
    STATUS_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?gid=511860805&single=true&output=csv"
    
    # Google Forms
    REPORT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeopdaE-lyOtFd2TUr5C3K2DWE3Syt2PaKoXMp0cmWKIFnijw/viewform?usp=header"
    
    # API Endpoints (for future expansion)
    API_BASE_URL = "https://api.wedabay.airport/v1"
    WEBSOCKET_URL = "wss://ws.wedabay.airport/notifications"


class DivisionRegistry:
    """
    Registry pattern for managing division configurations.
    Provides centralized access to all division data with validation.
    """
    
    _divisions: Dict[str, DivisionConfig] = {}
    
    @classmethod
    def register(cls, division: DivisionConfig) -> None:
        """Register a new division configuration."""
        # MODIFIED: Overwrite existing keys to prevent Duplicate Error on Reload
        cls._divisions[division.name] = division
    
    @classmethod
    def get(cls, name: str) -> Optional[DivisionConfig]:
        """Retrieve division configuration by name."""
        return cls._divisions.get(name)
    
    @classmethod
    def get_all(cls) -> Dict[str, DivisionConfig]:
        """Get all registered divisions."""
        return cls._divisions.copy()
    
    @classmethod
    def find_by_member(cls, member_name: str) -> Optional[DivisionConfig]:
        """Find division by member name."""
        for division in cls._divisions.values():
            if member_name in division.members:
                return division
        return None
    
    @classmethod
    def get_all_members(cls) -> List[str]:
        """Get all employee names across all divisions."""
        members = []
        
        # 1. Pastikan urutan divisi berdasarkan Priority (1, 2, 3...)
        sorted_divisions = sorted(cls._divisions.values(), key=lambda x: x.priority)
        
        for division in sorted_divisions:
            members.extend(division.members)
            
        # 2. Hapus nama ganda TAPI JANGAN di-sort A-Z. 
        return list(dict.fromkeys(members))


# Initialize Division Registry with actual data
def initialize_divisions():
    """
    Initialize all division configurations with CUSTOM SORT ORDER defined by User.
    """
    divisions_data = [
        DivisionConfig("SPT & SPV", "#FFD700", "👨‍✈️", "SPT & SPV", description="SPT & SPV", priority=1, 
                       members=["Patra Anggana", "Su Adam", "Budiman Arifin", "Rifaldy Ilham Bhagaskara", "Marwan S Halid", "Budiono"]),
        
        DivisionConfig("TLB", "#00A8FF", "🔧", "TLB", description="Teknik Listrik Bandara", priority=2, 
                       members=["M. Ansori", "Bayu Pratama Putra Katuwu", "Yoga Nugraha Putra Pasaribu", "Junaidi Taib", "Muhammad Rizal Amra", "Rusli Dj"]),
        
        DivisionConfig("TBL", "#0097e6", "📦", "TBL", description="Teknik Bangunan Dan Landasan", priority=3, 
                       members=["Venesia Aprilia Ineke", "Muhammad Naufal Ramadhan", "Yuzak Gerson Puturuhu", "Muhamad Alief Wildan", "Gafur Hamisi", "Jul Akbar M. Nur", "Sarni Massiri", "Adrianto Laundang", "Wahyudi Ismail"]),
        
        DivisionConfig("TRANS APRON", "#e1b12c", "🚌", "APR", description="Trans Apron", priority=4, 
                       members=["Marichi Gita Rusdi", "Ilham Rahim", "Abdul Mu Iz Simal", "Dwiki Agus Saputro", "Moh. Sofyan", "Faisal M. Kadir", "Amirudin Rustam", "Faturrahman Kaunar", "Wawan Hermawan", "Rahmat Joni", "Nur Ichsan"]),
        
        DivisionConfig("ATS", "#44bd32", "📡", "ATS", description="Air Traffic Services", priority=5, 
                       members=["Nurul Tanti", "Firlon Paembong", "Irwan Rezky Setiawan", "Yusuf Arviansyah", "Nurdahlia Is. Folaimam", "Ghaly Rabbani Panji Indra", "Ikhsan Wahyu Vebriyan", "Rizki Mahardhika Ardi Tigo", "Nikolaus Vincent Quirino"]),
        
        DivisionConfig("ADM COMPLIANCE", "#8c7ae6", "📋", "ADM", description="Administration & Compliance", priority=6, 
                       members=["Yessicha Aprilyona Siregar", "Gabriela Margrith Louisa Klavert", "Aldi Saptono"]),
        
        DivisionConfig("TRANSLATOR", "#00cec9", "🎧", "TRN", description="Translation Services", priority=7, 
                       members=["Wilyam Candra", "Norika Joselyn Modnissa"]),
        
        DivisionConfig("AVSEC", "#c23616", "🛡️", "SEC", description="Aviation Security", priority=8, 
                       members=["Andrian Maranatha", "Toni Nugroho Simarmata", "Muhamad Albi Ferano", "Andreas Charol Tandjung", "Sabadia Mahmud", "Rusdin Malagapi", "Muhamad Judhytia Winli", "Wahyu Samsudin", "Fientje Elisabeth Joseph", "Anglie Fitria Desiana Mamengko", "Dwi Purnama Bimasakti", "Windi Angriani Sulaeman", "Megawati A. Rauf"]),
        
        DivisionConfig("GROUND HANDLING", "#e17055", "🚜", "GND", description="Ground Handling Operations", priority=9, 
                       members=["Yuda Saputra", "Tesalonika Gratia Putri Toar", "Esi Setia Ningseh", "Ardiyanto Kalatjo", "Febrianti Tikabala"]),
        
        DivisionConfig("HELICOPTER", "#6c5ce7", "🚁", "HEL", description="Helicopter Operations", priority=10, 
                       members=["Agung Sabar S. Taufik", "Recky Irwan R. A Arsyad", "Farok Abdul", "Achmad Rizky Ariz", "Yus Andi", "Muh. Noval Kipudjena"]),
        
        DivisionConfig("AMC & TERMINAL", "#0984e3", "🏢", "AMC", description="Airport Movement Control & Terminal", priority=11, 
                       members=["Risky Sulung", "Muchamad Nur Syaifulrahman", "Muhammad Tunjung Rohmatullah", "Sunarty Fakir", "Albert Papuling", "Gibhran Fitransyah Yusri", "Muhdi R Tomia", "Riski Rifaldo Theofilus Anu", "Eko"]),
        
        DivisionConfig("SAFETY OFFICER", "#fd79a8", "🦺", "SFT", description="Safety Operations", priority=12, 
                       members=["Hildan Ahmad Zaelani", "Abdurahim Andar"]),
        
        DivisionConfig("PKP-PK", "#fab1a0", "🚒", "RES", description="Fire & Rescue Services", priority=13, 
                       members=["Andreas Aritonang", "Achmad Alwan Asyhab", "Doni Eka Satria", "Yogi Prasetya Eka Winandra", "Akhsin Aditya Weza Putra", "Fardhan Ahmad Tajali", "Maikel Renato Syafaruddin", "Saldi Sandra", "Hamzah M. Ali Gani", "Marfan Mandar", "Julham Keya", "Aditya Sugiantoro Abbas", "Muhamad Usman", "M Akbar D Patty", "Daniel Freski Wangka", "Fandi M.Naser", "Agung Fadjriansyah Ano", "Deni Hendri Bobode", "Muhammad Rifai", "Idrus Arsad"])
    ]
    for division in divisions_data:
        DivisionRegistry.register(division)

# ================================================================================
# SECTION 2: DATA ACCESS LAYER (REPOSITORY PATTERN)
# ================================================================================

class UserRepository:
    """
    Mengelola data user (Login & Auth).
    """
    def __init__(self):
        # Default Admin Account & User
        self.default_users = {
            "admin": {"password": "hancok1234", "role": "admin", "name": "System Administrator", "status": "Active"},
            "user": {"password": "user123", "role": "user", "name": "General Viewer", "status": "Active"}
        }
        
    def get_user(self, username):
        return self.default_users.get(username)

    def get_all_users(self):
        return self.default_users

    def update_user_status(self, username, status):
        if username in self.default_users:
            self.default_users[username]['status'] = status
            return True
        return False
        
    def add_user(self, username, password, name, role="user"):
        if username in self.default_users:
            return False
        self.default_users[username] = {
            "password": password, 
            "role": role, 
            "name": name, 
            "status": "Pending"
        }
        return True


class DataRepository(ABC):
    """
    Abstract base class for data repositories.
    Follows Repository Pattern for data access abstraction.
    """
    
    @abstractmethod
    def fetch(self) -> Optional[pd.DataFrame]:
        """Fetch data from source."""
        pass
    
    @abstractmethod
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate fetched data."""
        pass
    
    @abstractmethod
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform raw data to application format."""
        pass


class AttendanceRepository(DataRepository):
    """
    Repository for attendance data management.
    Handles data fetching, validation, and transformation.
    """
    
    def __init__(self, url: str):
        self.url = url
        self._cache: Optional[pd.DataFrame] = None
        self._cache_time: Optional[datetime] = None
    
    @st.cache_data(ttl=AppConstants.CACHE_TTL_SECONDS)
    def fetch(_self) -> Optional[pd.DataFrame]:
        """
        Fetch attendance data from Google Sheets.
        Uses Streamlit caching for performance optimization.
        """
        try:
            df = pd.read_csv(_self.url)
            
            # Standardize column names
            df.columns = df.columns.str.strip()
            
            if not _self.validate(df):
                st.error("❌ Attendance data validation failed")
                return None
            
            return _self.transform(df)
            
        except Exception as e:
            st.error(f"❌ Failed to fetch attendance data: {str(e)}")
            return None
    
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate that required columns exist."""
        required_columns = [AppConstants.COL_PERSON_NAME, AppConstants.COL_EVENT_TIME]
        return all(col in df.columns for col in required_columns)
    
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Transform raw attendance data.
        Applies cleaning, type conversion, and feature engineering.
        """
        # Clean employee names
        df[AppConstants.COL_PERSON_NAME] = (
            df[AppConstants.COL_PERSON_NAME]
            .astype(str)
            .str.strip()
        )
        
        # Remove empty records
        df = df[df[AppConstants.COL_PERSON_NAME] != ''].copy()
        
        # Parse timestamps
        df[AppConstants.COL_EVENT_TIME] = pd.to_datetime(
            df[AppConstants.COL_EVENT_TIME],
            errors='coerce'
        )
        
        # Extract date and time components
        df['Tanggal'] = df[AppConstants.COL_EVENT_TIME].dt.date
        df['Waktu'] = df[AppConstants.COL_EVENT_TIME].dt.time
        df['Jam'] = df[AppConstants.COL_EVENT_TIME].dt.hour
        df['Menit'] = df[AppConstants.COL_EVENT_TIME].dt.minute
        
        # Add day of week
        df['Hari'] = df[AppConstants.COL_EVENT_TIME].dt.day_name()
        
        return df


class StatusRepository(DataRepository):
    """
    Repository for employee status (permits, leaves) management.
    """
    
    def __init__(self, url: str):
        self.url = url
    
    @st.cache_data(ttl=AppConstants.CACHE_TTL_SECONDS)
    def fetch(_self) -> Optional[pd.DataFrame]:
        """Fetch status data from Google Sheets."""
        try:
            df = pd.read_csv(_self.url)
            df = df.rename(columns=lambda x: x.strip())
            
            if not _self.validate(df):
                st.warning("⚠️ Status data validation failed")
                return None
            
            return _self.transform(df)
            
        except Exception as e:
            st.warning(f"⚠️ Failed to fetch status data: {str(e)}")
            return None
    
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate status data structure."""
        required = [AppConstants.COL_EMPLOYEE_NAME, AppConstants.COL_DATE, AppConstants.COL_STATUS]
        return all(col in df.columns for col in required)
    
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform status data."""
        # Clean names
        df[AppConstants.COL_EMPLOYEE_NAME] = (
            df[AppConstants.COL_EMPLOYEE_NAME]
            .astype(str)
            .str.strip()
        )
        
        # Parse dates
        df[AppConstants.COL_DATE] = pd.to_datetime(
            df[AppConstants.COL_DATE],
            format='mixed',
            dayfirst=False,
            errors='coerce'
        ).dt.date # Ensure date format consistency
        
        df['Tanggal_Str'] = pd.to_datetime(df[AppConstants.COL_DATE]).dt.strftime(AppConstants.DATE_FORMAT)
        
        # Standardize status text
        df[AppConstants.COL_STATUS] = (
            df[AppConstants.COL_STATUS]
            .astype(str)
            .str.strip()
            .str.upper()
        )
        
        return df


# ================================================================================
# SECTION 3: BUSINESS LOGIC LAYER (SERVICE CLASSES)
# ================================================================================

class AuthService:
    """
    Menangani logic Login, Logout, dan Session Check.
    """
    def __init__(self):
        self.repo = UserRepository()

    def login(self, username, password):
        user = self.repo.get_user(username)
        if user:
            if user['password'] == password:
                if user['status'] != 'Active':
                    return False, "Akun belum diaktifkan oleh Admin."
                return True, user
            else:
                return False, "Password salah."
        return False, "Username tidak ditemukan."


class TimeService:
    """
    Service class for time-related business logic.
    Centralizes all time calculations and comparisons.
    """
    
    @staticmethod
    def is_late(time_str: Optional[str], threshold: time = AppConstants.LATE_THRESHOLD) -> bool:
        """
        Check if a time string represents a late arrival.
        """
        if not time_str:
            return False
        
        try:
            check_time = datetime.strptime(time_str, AppConstants.TIME_FORMAT).time()
            return check_time > threshold
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def is_early(time_str: Optional[str], threshold: time = AppConstants.EARLY_ARRIVAL) -> bool:
        """Check if arrival is early."""
        if not time_str:
            return False
        
        try:
            check_time = datetime.strptime(time_str, AppConstants.TIME_FORMAT).time()
            return check_time < threshold
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def calculate_duration(start_time: str, end_time: str) -> Optional[timedelta]:
        """
        Calculate duration between two times.
        """
        try:
            start = datetime.strptime(start_time, AppConstants.TIME_FORMAT)
            end = datetime.strptime(end_time, AppConstants.TIME_FORMAT)
            
            # Handle overnight shifts
            if end < start:
                end += timedelta(days=1)
            
            return end - start
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def format_duration(duration: timedelta) -> str:
        """Format timedelta as human-readable string."""
        if not duration:
            return "N/A"
        
        hours = duration.seconds // 3600
        minutes = (duration.seconds % 3600) // 60
        
        return f"{hours}h {minutes}m"
    
    @staticmethod
    def get_time_range_label(check_time: time) -> str:
        """Determine which time range a given time falls into."""
        for time_range in TimeRanges:
            start = time.fromisoformat(time_range.start_time)
            end = time.fromisoformat(time_range.end_time)
            
            if start <= check_time <= end:
                return time_range.label
        
        return "UNKNOWN"


class AttendanceService:
    """
    Core business logic service for attendance processing.
    """
    
    def __init__(self, attendance_repo: AttendanceRepository, status_repo: StatusRepository):
        self.attendance_repo = attendance_repo
        self.status_repo = status_repo
        self.time_service = TimeService()
    
    def get_attendance_for_date(self, target_date: datetime.date) -> Optional[pd.DataFrame]:
        df = self.attendance_repo.fetch()
        if df is None: return None
        return df[df['Tanggal'] == target_date].copy()
    
    def get_status_for_date(self, target_date: datetime.date) -> Dict[str, str]:
        df = self.status_repo.fetch()
        if df is None: return {}
        date_str = target_date.strftime(AppConstants.DATE_FORMAT)
        df_filtered = df[df['Tanggal_Str'] == date_str]
        if df_filtered.empty: return {}
        return pd.Series(
            df_filtered[AppConstants.COL_STATUS].values,
            index=df_filtered[AppConstants.COL_EMPLOYEE_NAME]
        ).to_dict()
    
    def extract_time_ranges(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        LOGIKA SMART RANGE (NO GAPS) - REVISI JUMAT STRICT:
        """
        if df.empty: return pd.DataFrame()

        df_clean = df.dropna(subset=[AppConstants.COL_PERSON_NAME, 'Tanggal'])
        if df_clean.empty: return pd.DataFrame()

        # Pastikan Waktu dalam format datetime yang benar
        df_clean['Waktu_Obj'] = pd.to_datetime(df_clean[AppConstants.COL_EVENT_TIME]).dt.time

        grouped = df_clean.groupby([AppConstants.COL_PERSON_NAME, 'Tanggal'])
        
        def process_group(group):
            result = {
                'Pagi': '',
                'Siang_1': '',
                'Siang_2': '',
                'Sore': ''
            }
            
            # 1. Cek Hari (0=Senin, 4=Jumat, 6=Minggu)
            current_date = group['Tanggal'].iloc[0]
            is_friday = (current_date.weekday() == 4)

            # 2. Tentukan Batas Waktu (Thresholds)
            if is_friday:
                # --- JUMAT ---
                limit_pagi    = time(12, 0, 0) # Batas akhir pagi
                limit_siang1  = time(13, 0, 0) # Batas akhir siang 1 (mulai siang 2)
                limit_siang2  = time(14, 0, 0) # Batas akhir siang 2 (STRICT)
                start_sore    = time(17, 0, 0) # Mulai sore
            else:
                # --- HARI BIASA ---
                limit_pagi    = time(11, 30, 0)
                limit_siang1  = time(12, 30, 0)
                limit_siang2  = time(16, 0, 0)
                start_sore    = time(16, 0, 0)

            # Sortir waktu
            sorted_group = group.sort_values(AppConstants.COL_EVENT_TIME)
            
            for _, row in sorted_group.iterrows():
                t = row['Waktu_Obj'] 
                val_str = row[AppConstants.COL_EVENT_TIME].strftime(AppConstants.TIME_FORMAT)
                
                # --- LOGIKA PEMBAGIAN WAKTU ---
                
                # 1. PAGI
                if t < limit_pagi:
                    if result['Pagi'] == '': 
                        result['Pagi'] = val_str
                
                # 2. SIANG 1
                elif limit_pagi <= t < limit_siang1:
                    if result['Siang_1'] == '':
                        result['Siang_1'] = val_str
                        
                # 3. SIANG 2 (Berhenti tepat di limit_siang2)
                elif limit_siang1 <= t <= limit_siang2: # Pakai <= agar 14:00 pas masuk
                    if result['Siang_2'] == '':
                        result['Siang_2'] = val_str
                        
                # 4. SORE
                elif t >= start_sore:
                    result['Sore'] = val_str # Overwrite (ambil terakhir)

            return pd.Series(result)

        if grouped.ngroups == 0: return pd.DataFrame()

        result_df = grouped.apply(process_group).reset_index()
        result_df.rename(columns={AppConstants.COL_PERSON_NAME: AppConstants.COL_EMPLOYEE_NAME}, inplace=True)
        
        return result_df

    def build_complete_report(self, target_date: datetime.date) -> Tuple[pd.DataFrame, Dict[str, str]]:
        """
        Builds the master dataframe merging attendance times with employee list.
        """
        df_attendance = self.get_attendance_for_date(target_date)
        status_dict = self.get_status_for_date(target_date)
        
        if df_attendance is not None and not df_attendance.empty:
            df_times = self.extract_time_ranges(df_attendance)
        else:
            df_times = pd.DataFrame()
        
        all_employees = DivisionRegistry.get_all_members()
        df_all = pd.DataFrame({AppConstants.COL_EMPLOYEE_NAME: all_employees})
        
        if not df_times.empty:
            df_final = pd.merge(df_all, df_times, on=AppConstants.COL_EMPLOYEE_NAME, how='left')
        else:
            df_final = df_all.copy()
            for col in ['Pagi', 'Siang_1', 'Siang_2', 'Sore']:
                df_final[col] = ''

        df_final.fillna('', inplace=True)
        return df_final, status_dict

    def calculate_metrics(self, df: pd.DataFrame, status_dict: Dict[str, str]) -> Dict[str, Any]:
        """
        Calculates daily statistics (Present, Absent, Late, etc.)
        """
        total_employees = len(df)
        present_count = 0; permit_count = 0; absent_count = 0; late_count = 0
        late_list = []; permit_list = []; absent_list = []; partial_list = []
        
        for idx, row in df.iterrows():
            name = row[AppConstants.COL_EMPLOYEE_NAME]
            times = [row.get('Pagi', ''), row.get('Siang_1', ''), 
                     row.get('Siang_2', ''), row.get('Sore', '')]
            empty_count = sum(1 for t in times if t == '')
            manual_status = status_dict.get(name, "")
            
            if manual_status:
                permit_count += 1
                permit_list.append((name, manual_status))
            elif empty_count == 4:
                absent_count += 1
                absent_list.append(name)
            else:
                present_count += 1
                morning_time = row.get('Pagi', '')
                if morning_time and self.time_service.is_late(morning_time):
                    late_count += 1
                    late_list.append((name, morning_time))
                if empty_count > 0:
                    partial_list.append((name, empty_count))
        
        return {
            'total': total_employees, 'present': present_count, 'permit': permit_count,
            'absent': absent_count, 'late': late_count, 'late_list': late_list,
            'permit_list': permit_list, 'absent_list': absent_list, 'partial_list': partial_list,
            'attendance_rate': (present_count / total_employees * 100) if total_employees > 0 else 0,
            'punctuality_rate': ((present_count - late_count) / present_count * 100) if present_count > 0 else 0
        }

class AnalyticsService:
    """
    Service for advanced analytics and reporting.
    Provides statistical analysis and trend detection.
    """
    
    def __init__(self, attendance_repo: AttendanceRepository):
        self.attendance_repo = attendance_repo
    
    def get_weekly_trends(self, end_date: datetime.date, weeks: int = 4) -> pd.DataFrame:
        """
        Get attendance trends over multiple weeks.
        """
        df = self.attendance_repo.fetch()
        if df is None:
            return pd.DataFrame()
        
        start_date = end_date - timedelta(weeks=weeks)
        mask = (df['Tanggal'] >= start_date) & (df['Tanggal'] <= end_date)
        df_period = df[mask].copy()
        
        # Group by week
        df_period['Week'] = df_period[AppConstants.COL_EVENT_TIME].dt.isocalendar().week
        df_period['Year'] = df_period[AppConstants.COL_EVENT_TIME].dt.year
        
        weekly_stats = df_period.groupby(['Year', 'Week']).agg({
            AppConstants.COL_PERSON_NAME: 'nunique',
            AppConstants.COL_EVENT_TIME: 'count'
        }).reset_index()
        
        weekly_stats.columns = ['Year', 'Week', 'Unique_Employees', 'Total_Events']
        
        return weekly_stats
    
    def get_division_statistics(self, target_date: datetime.date) -> Dict[str, Dict]:
        """
        Calculate statistics per division.
        """
        df = self.attendance_repo.fetch()
        if df is None:
            return {}
        
        df_day = df[df['Tanggal'] == target_date]
        stats = {}
        
        for division_name, division_config in DivisionRegistry.get_all().items():
            members = division_config.members
            df_div = df_day[df_day[AppConstants.COL_PERSON_NAME].isin(members)]
            
            present = df_div[AppConstants.COL_PERSON_NAME].nunique()
            total = len(members)
            
            stats[division_name] = {
                'total': total,
                'present': present,
                'absent': total - present,
                'rate': (present / total * 100) if total > 0 else 0,
                'color': division_config.color,
                'icon': division_config.icon
            }
        
        return stats
    
    def detect_anomalies(self, df: pd.DataFrame, threshold_hours: int = 12) -> List[Dict]:
        """
        Detect anomalous attendance patterns.
        """
        anomalies = []
        
        for name in df[AppConstants.COL_PERSON_NAME].unique():
            person_data = df[df[AppConstants.COL_PERSON_NAME] == name].sort_values(AppConstants.COL_EVENT_TIME)
            
            if len(person_data) > 1:
                times = person_data[AppConstants.COL_EVENT_TIME].tolist()
                
                for i in range(len(times) - 1):
                    gap = (times[i+1] - times[i]).total_seconds() / 3600
                    
                    if gap > threshold_hours:
                        anomalies.append({
                            'employee': name,
                            'type': 'LARGE_GAP',
                            'gap_hours': gap,
                            'time1': times[i],
                            'time2': times[i+1]
                        })
        
        return anomalies


# ================================================================================
# SECTION 4: EXPORT & REPORTING LAYER
# ================================================================================

class ExcelExporter:
    def __init__(self):
        self.workbook = None
        self.formats = {}

    def _init_formats(self, workbook):
        """Helper to initialize formats only once per workbook"""
        self.fmt_head = workbook.add_format({'bold': True, 'fg_color': '#4caf50', 'font_color': 'white', 'border': 1, 'align': 'center'})
        self.fmt_norm = workbook.add_format({'border': 1, 'align': 'center'})
        self.fmt_miss = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center'}) 
        self.fmt_full = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'align': 'center'}) 
        self.fmt_late = workbook.add_format({'font_color': 'red', 'bold': True, 'border': 1, 'align': 'center'})

    def _write_sheet_content(self, ws, df: pd.DataFrame, status_dict: Dict[str, str]):
        """
        Internal logic to write a single sheet. 
        Reused by both single date and range reports.
        """
        # Headers
        headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore', 'Keterangan']
        ws.write_row(0, 0, headers, self.fmt_head)
        ws.set_column(0, 0, 30) # Lebar kolom Nama
        ws.set_column(1, 5, 15) # Lebar kolom Waktu & Ket
        
        # Writing Logic
        for idx, row in df.iterrows():
            row_num = idx + 1
            nm = row[AppConstants.COL_EMPLOYEE_NAME]
            pagi = row.get('Pagi', '')
            siang1 = row.get('Siang_1', '')
            siang2 = row.get('Siang_2', '')
            sore = row.get('Sore', '')
            
            manual_stat = status_dict.get(nm, "")
            times = [pagi, siang1, siang2, sore]
            empty = sum(1 for t in times if t == '')
            
            # Tulis Nama & Keterangan Normal Dulu
            ws.write(row_num, 0, nm, self.fmt_norm)
            ws.write(row_num, 5, manual_stat, self.fmt_norm)
            
            # LOGIKA WARNA
            if manual_stat:
                # Izin/Sakit: Kosongkan waktu dengan format normal
                for i in range(4): ws.write(row_num, i+1, "", self.fmt_norm)
                
            elif empty == 4:
                # Alpha (4 Bolong): Isi waktu dengan format FULL KUNING
                for i in range(4): ws.write(row_num, i+1, "", self.fmt_full)
                
            else:
                # Hadir (Cek satu per satu)
                # 1. Pagi
                if pagi == '': 
                    ws.write(row_num, 1, "", self.fmt_miss) 
                else:
                    if TimeService.is_late(pagi): 
                        ws.write(row_num, 1, pagi, self.fmt_late) 
                    else: 
                        ws.write(row_num, 1, pagi, self.fmt_norm) 
                
                # 2. Istirahat & Pulang
                rest_times = [siang1, siang2, sore]
                for i, t in enumerate(rest_times):
                    col_idx = i + 2
                    if t == '': 
                        ws.write(row_num, col_idx, "", self.fmt_miss) 
                    else: 
                        ws.write(row_num, col_idx, t, self.fmt_norm)

    def create_attendance_report(self, df: pd.DataFrame, status_dict: Dict[str, str], date: datetime.date, metrics: Any = None):
        """Creates a single sheet report"""
        output = io.BytesIO()
        self.workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._init_formats(self.workbook)
        
        sheet_name = date.strftime('%d-%b') # e.g., 29-Nov
        ws = self.workbook.add_worksheet(sheet_name)
        
        self._write_sheet_content(ws, df, status_dict)
        
        self.workbook.close()
        output.seek(0)
        return output

    def create_range_report(self, data_map: Dict[datetime.date, Tuple[pd.DataFrame, Dict]]):
        """
        Creates a multi-sheet Excel report for a date range.
        data_map: Dictionary where Key = Date, Value = (DataFrame, StatusDict)
        """
        output = io.BytesIO()
        self.workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._init_formats(self.workbook)

        # Sort dates to ensure tabs are in order
        sorted_dates = sorted(data_map.keys())

        for date in sorted_dates:
            df, status_dict = data_map[date]
            # Sheet name cannot handle special chars or be too long
            sheet_name = date.strftime('%d-%b') 
            ws = self.workbook.add_worksheet(sheet_name)
            self._write_sheet_content(ws, df, status_dict)

        self.workbook.close()
        output.seek(0)
        return output

# ================================================================================
# SECTION 5: UI STYLING LAYER
# ================================================================================

class ThemeManager:
    """
    Centralized theme and styling management.
    """
    
    @staticmethod
    def apply_global_styles():
        """Apply comprehensive CSS styling."""
        st.markdown("""
        <style>
            /* ========== FONT IMPORTS ========== */
            @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Inter:wght@300;400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
            
            /* ========== CSS VARIABLES ========== */
            :root {
                --bg-primary: #0b0e14;
                --bg-secondary: #151922;
                --bg-tertiary: #1e2530;
                --accent-blue: #00a8ff;
                --accent-cyan: #00cec9;
                --accent-purple: #8c7ae6;
                --alert-red: #e84118;
                --alert-amber: #fbc531;
                --success-green: #4cd137;
                --text-primary: #f5f6fa;
                --text-secondary: #7f8fa6;
                --text-muted: #546e7a;
                --border-color: #2f3640;
                --border-glow: rgba(0, 168, 255, 0.3);
                --shadow-sm: 0 2px 4px rgba(0,0,0,0.3);
                --shadow-md: 0 4px 12px rgba(0,0,0,0.4);
                --shadow-lg: 0 8px 24px rgba(0,0,0,0.5);
                --shadow-glow: 0 0 20px rgba(0, 168, 255, 0.2);
                --transition-fast: 0.2s ease;
                --transition-normal: 0.3s ease;
                --transition-slow: 0.5s ease;
            }
            
            /* ========== BASE STYLES ========== */
            .stApp {
                background: var(--bg-primary);
                background-image: 
                    radial-gradient(circle at 20% 10%, rgba(0, 168, 255, 0.05) 0%, transparent 50%),
                    radial-gradient(circle at 80% 90%, rgba(140, 122, 230, 0.05) 0%, transparent 50%);
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
                color: var(--text-primary);
            }
            
            /* ========== SCROLLBAR STYLING ========== */
            ::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            
            ::-webkit-scrollbar-track {
                background: var(--bg-secondary);
            }
            
            ::-webkit-scrollbar-thumb {
                background: var(--accent-blue);
                border-radius: 4px;
            }
            
            ::-webkit-scrollbar-thumb:hover {
                background: var(--accent-cyan);
            }
            
            /* ========== SIDEBAR STYLING ========== */
            section[data-testid="stSidebar"] {
                background: linear-gradient(180deg, #0f1218 0%, #1a1f2e 100%);
                border-right: 1px solid var(--border-color);
                box-shadow: 4px 0 12px rgba(0,0,0,0.3);
            }
            
            section[data-testid="stSidebar"] .stMarkdown {
                padding: 0.5rem 0;
            }
            
            /* ========== TYPOGRAPHY ========== */
            h1, h2, h3, h4, h5, h6 {
                font-family: 'Rajdhani', sans-serif;
                text-transform: uppercase;
                letter-spacing: 1.5px;
                font-weight: 700;
            }
            
            .brand-title {
                font-family: 'Rajdhani', sans-serif;
                font-size: 3.5rem;
                font-weight: 700;
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan), var(--accent-purple));
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
                margin-bottom: 0;
                text-shadow: 0 0 30px var(--border-glow);
                animation: titlePulse 3s ease-in-out infinite;
            }
            
            @keyframes titlePulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.8; }
            }
            
            .brand-subtitle {
                font-family: 'JetBrains Mono', monospace;
                color: var(--text-secondary);
                font-size: 0.85rem;
                letter-spacing: 4px;
                text-transform: uppercase;
                margin-bottom: 30px;
                border-bottom: 2px solid var(--border-color);
                padding-bottom: 20px;
                position: relative;
            }
            
            .brand-subtitle::after {
                content: '';
                position: absolute;
                bottom: -2px;
                left: 0;
                width: 100px;
                height: 2px;
                background: var(--accent-blue);
                animation: lineExpand 2s ease-in-out infinite;
            }
            
            @keyframes lineExpand {
                0%, 100% { width: 100px; }
                50% { width: 200px; }
            }
            
            /* ========== METRICS STYLING ========== */
            div[data-testid="stMetric"] {
                background: linear-gradient(135deg, rgba(21, 25, 34, 0.8), rgba(30, 37, 48, 0.6));
                border: 1px solid var(--border-color);
                border-left: 4px solid var(--accent-blue);
                border-radius: 8px;
                padding: 20px;
                transition: all var(--transition-normal);
                position: relative;
                overflow: hidden;
            }
            
            div[data-testid="stMetric"]::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(45deg, transparent, rgba(0, 168, 255, 0.05), transparent);
                transform: translateX(-100%);
                transition: transform var(--transition-slow);
            }
            
            div[data-testid="stMetric"]:hover {
                border-color: var(--accent-cyan);
                box-shadow: var(--shadow-glow);
                transform: translateY(-4px);
            }
            
            div[data-testid="stMetric"]:hover::before {
                transform: translateX(100%);
            }
            
            div[data-testid="stMetricLabel"] {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 600;
                color: var(--text-secondary);
                font-size: 0.9rem;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            
            div[data-testid="stMetricValue"] {
                font-family: 'JetBrains Mono', monospace;
                font-size: 2.5rem;
                color: var(--text-primary);
                font-weight: 700;
                text-shadow: 0 2px 4px rgba(0,0,0,0.3);
            }
            
            div[data-testid="stMetricDelta"] {
                font-family: 'Inter', sans-serif;
                font-size: 0.85rem;
            }
            
            /* ========== CARD COMPONENTS ========== */
            .card {
                background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
                border: 1px solid var(--border-color);
                border-radius: 10px;
                margin-bottom: 20px;
                transition: all var(--transition-normal);
                position: relative;
                overflow: hidden;
                box-shadow: var(--shadow-md);
            }
            
            .card::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 3px;
                background: linear-gradient(90deg, var(--accent-blue), var(--accent-cyan));
                transform: scaleX(0);
                transform-origin: left;
                transition: transform var(--transition-normal);
            }
            
            .card:hover {
                border-color: var(--accent-blue);
                box-shadow: var(--shadow-lg), var(--shadow-glow);
                transform: translateY(-4px);
            }
            
            .card:hover::before {
                transform: scaleX(1);
            }
            
            .card-header {
                padding: 18px;
                display: flex;
                align-items: center;
                background: rgba(255, 255, 255, 0.03);
                border-bottom: 1px dashed var(--border-color);
                position: relative;
            }
            
            .card-header::after {
                content: '';
                position: absolute;
                bottom: 0;
                left: 0;
                width: 60px;
                height: 1px;
                background: var(--accent-cyan);
            }
            
            .card-body {
                padding: 18px;
            }
            
            .card-name {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 700;
                font-size: 1.15rem;
                margin: 0;
                color: var(--text-primary);
                line-height: 1.3;
            }
            
            .card-code {
                font-family: 'JetBrains Mono', monospace;
                font-size: 0.75rem;
                color: var(--text-secondary);
                letter-spacing: 1.5px;
                font-weight: 600;
            }
            
            /* ========== FLIGHT LOG STYLES ========== */
            .flight-row {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 10px;
                padding: 8px 0;
                border-bottom: 1px solid rgba(255,255,255,0.05);
            }
            
            .flight-row:last-child {
                border-bottom: none;
            }
            
            .flight-label {
                font-size: 0.7rem;
                color: var(--text-muted);
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 600;
            }
            
            .flight-value {
                font-family: 'JetBrains Mono', monospace;
                font-weight: 700;
                font-size: 1rem;
                color: var(--text-primary);
            }
            
            /* ========== STATUS BADGES ========== */
            .status-badge {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 700;
                font-size: 0.8rem;
                padding: 4px 12px;
                border-radius: 4px;
                display: inline-block;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            
            .status-present {
                color: var(--success-green);
                border: 1px solid var(--success-green);
                background: rgba(76, 209, 55, 0.1);
                box-shadow: 0 0 10px rgba(76, 209, 55, 0.2);
            }
            
            .status-partial {
                color: var(--alert-amber);
                border: 1px solid var(--alert-amber);
                background: rgba(251, 197, 49, 0.1);
                box-shadow: 0 0 10px rgba(251, 197, 49, 0.2);
            }
            
            .status-absent {
                color: var(--alert-red);
                border: 1px solid var(--alert-red);
                background: rgba(232, 65, 24, 0.1);
                box-shadow: 0 0 10px rgba(232, 65, 24, 0.2);
            }
            
            .status-permit {
                color: #9c88ff;
                border: 1px solid #9c88ff;
                background: rgba(156, 136, 255, 0.1);
                box-shadow: 0 0 10px rgba(156, 136, 255, 0.2);
            }
            
            /* ========== ANOMALY BOXES ========== */
            .anomaly-box {
                padding: 12px 15px;
                margin-bottom: 10px;
                border-left: 4px solid;
                background: rgba(255, 255, 255, 0.03);
                font-family: 'JetBrains Mono', monospace;
                font-size: 0.85rem;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-radius: 4px;
                transition: all var(--transition-fast);
            }
            
            .anomaly-box:hover {
                background: rgba(255, 255, 255, 0.06);
                transform: translateX(4px);
            }
            
            .box-telat {
                border-color: var(--alert-red);
                background: linear-gradient(90deg, rgba(232, 65, 24, 0.15), transparent);
            }
            
            .box-izin {
                border-color: #9c88ff;
                background: linear-gradient(90deg, rgba(156, 136, 255, 0.15), transparent);
            }
            
            .box-alpha {
                border-color: var(--alert-amber);
                background: linear-gradient(90deg, rgba(251, 197, 49, 0.15), transparent);
            }
            
            /* ========== BUTTONS ========== */
            .stDownloadButton button, .stButton button {
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan)) !important;
                color: #000 !important;
                font-weight: 800 !important;
                border: none !important;
                border-radius: 6px !important;
                padding: 12px 24px !important;
                text-transform: uppercase !important;
                letter-spacing: 1px !important;
                transition: all var(--transition-normal) !important;
                box-shadow: var(--shadow-md) !important;
            }
            
            .stDownloadButton button:hover, .stButton button:hover {
                transform: translateY(-2px) !important;
                box-shadow: var(--shadow-lg), var(--shadow-glow) !important;
            }
            
            /* ========== INPUT FIELDS ========== */
            .stTextInput input, .stDateInput input, .stSelectbox select {
                background-color: var(--bg-secondary) !important;
                color: var(--text-primary) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 6px !important;
                padding: 10px !important;
                transition: all var(--transition-fast) !important;
            }
            
            .stTextInput input:focus, .stDateInput input:focus {
                border-color: var(--accent-blue) !important;
                box-shadow: 0 0 0 2px rgba(0, 168, 255, 0.2) !important;
            }
            
            /* ========== EXPANDER ========== */
            .streamlit-expanderHeader {
                background: rgba(255, 255, 255, 0.05) !important;
                color: var(--accent-blue) !important;
                font-family: 'Rajdhani', sans-serif !important;
                font-weight: 600 !important;
                border-radius: 6px !important;
                padding: 12px 16px !important;
                transition: all var(--transition-fast) !important;
            }
            
            .streamlit-expanderHeader:hover {
                background: rgba(255, 255, 255, 0.08) !important;
                border-color: var(--accent-cyan) !important;
            }
            
            /* ========== DATAFRAME STYLING ========== */
            div[data-testid="stDataFrame"] {
                background-color: var(--bg-secondary);
                border: 1px solid var(--border-color);
                border-radius: 8px;
                overflow: hidden;
            }
            
            /* ========== TABS ========== */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px;
                background-color: var(--bg-secondary);
                border-radius: 8px;
                padding: 8px;
            }
            
            .stTabs [data-baseweb="tab"] {
                height: 50px;
                background-color: transparent;
                border-radius: 6px;
                color: var(--text-secondary);
                font-family: 'Rajdhani', sans-serif;
                font-weight: 600;
                font-size: 1rem;
                transition: all var(--transition-fast);
            }
            
            .stTabs [data-baseweb="tab"]:hover {
                background-color: rgba(255, 255, 255, 0.05);
                color: var(--text-primary);
            }
            
            .stTabs [aria-selected="true"] {
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan)) !important;
                color: #000 !important;
            }
            
            /* ========== LOGIN SCREEN STYLING ========== */
            .login-container {
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                background-image: url('https://images.unsplash.com/photo-1464037866556-6812c9d1c72e?q=80&w=2070&auto=format&fit=crop');
                background-size: cover;
                background-position: center;
            }
            
            .login-card {
                background: rgba(255, 255, 255, 0.95);
                padding: 40px;
                border-radius: 12px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.5);
                text-align: center;
                color: #333;
                margin-top: 100px;
            }
            
            .login-card .stTextInput label {
                color: #333 !important;
                text-align: left;
                display: block;
                font-weight: 600;
            }
            
            .login-card .stTextInput input {
                background-color: #f8f9fa !important;
                color: #333 !important;
                border: 1px solid #ddd !important;
            }
            
            .login-card .stButton button {
                width: 100%;
                background: #0066ff !important;
                color: white !important;
                border-radius: 6px !important;
                padding: 12px !important;
                font-weight: 600 !important;
                border: none !important;
            }
            
            .login-card .stButton button:hover {
                background: #0052cc !important;
            }
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def get_avatar_url(name: str) -> str:
        """Generate avatar URL for employee."""
        encoded_name = name.replace(' ', '+')
        return f"https://ui-avatars.com/api/?name={encoded_name}&background=random&color=fff&size=128&bold=true&rounded=true"


# ================================================================================
# SECTION 6: UI COMPONENT LAYER
# ================================================================================

class ComponentRenderer:
    """
    Centralized component rendering with consistent styling.
    Implements Facade pattern for complex UI operations.
    """
    
    def __init__(self):
        self.time_service = TimeService()
    
    def render_employee_card(
        self, 
        employee_data: pd.Series, 
        status_dict: Dict[str, str]
    ) -> None:
        """
        Render individual employee attendance card with boarding pass design.
        """
        name = employee_data[AppConstants.COL_EMPLOYEE_NAME]
        morning = employee_data.get('Pagi', '')
        break_out = employee_data.get('Siang_1', '')
        break_in = employee_data.get('Siang_2', '')
        evening = employee_data.get('Sore', '')
        
        # Get division info
        division = DivisionRegistry.find_by_member(name)
        if division:
            div_name = division.name
            div_code = division.code
            div_color = division.color
            div_icon = division.icon
        else:
            div_name = "UNKNOWN"
            div_code = "UNK"
            div_color = "#666"
            div_icon = "❓"
        
        # Determine status
        manual_status = status_dict.get(name, "")
        times = [morning, break_out, break_in, evening]
        empty_count = sum(1 for t in times if t == '')
        
        if manual_status:
            status = AttendanceStatus.PERMIT
            status_text = f"PERMIT: {manual_status}"
        elif empty_count == 4:
            status = AttendanceStatus.ABSENT
            status_text = status.display_text
        elif empty_count > 0:
            status = AttendanceStatus.PARTIAL_DUTY
            status_text = status.display_text
        else:
            status = AttendanceStatus.FULL_DUTY
            status_text = status.display_text
        
        # Check if late
        is_late = morning and self.time_service.is_late(morning)
        late_indicator = "<span style='color:#e84118; font-weight:bold; margin-left:8px;'>⚠ DELAY</span>" if is_late else ""
        
        # Get avatar
        avatar_url = ThemeManager.get_avatar_url(name)
        
        # Render card
        st.markdown(f"""
        <div class="card" style="border-left: 4px solid {status.color};">
            <div class="card-header">
                <img src="{avatar_url}" 
                    style="width:50px; height:50px; border-radius:50%; margin-right:15px; 
                             border:3px solid {div_color}; box-shadow: 0 4px 8px rgba(0,0,0,0.3);">
                <div style="flex:1; min-width:0;">
                    <p class="card-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                        {div_icon} {name}
                    </p>
                    <span class="card-code" style="color:{div_color}; font-weight:bold;">
                        ✈ {div_code}
                    </span>
                    <span class="card-code" style="color:#888; margin-left:8px;">
                        // {div_name}
                    </span>
                </div>
            </div>
            <div class="card-body">
                <div class="flight-row">
                    <span class="flight-label">🛫 Jam Datang</span>
                    <span class="flight-value" style="font-size:1.2rem;">
                        {morning if morning else '━━:━━'} {late_indicator}
                    </span>
                </div>
                <div class="flight-row">
                    <span class="flight-label">🛬 Jam Pulang</span>
                    <span class="flight-value" style="font-size:1.2rem;">
                        {evening if evening else '━━:━━'}
                    </span>
                </div>
                <div style="margin-top:12px; text-align:right;">
                    <span class="status-badge {status.css_class}">{status_text}</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Detail popover
        with st.popover("📋 DETAILED FLIGHT LOG", use_container_width=True):
            self._render_detail_popover(name, morning, break_out, break_in, evening, 
                                        status_text, status.color, div_name, is_late)
    
    def _render_detail_popover(
        self, 
        name: str, 
        morning: str, 
        break_out: str, 
        break_in: str, 
        evening: str,
        status_text: str,
        status_color: str,
        division: str,
        is_late: bool
    ) -> None:
        """Render detailed attendance information in popover."""
        st.markdown(f"### ✈️ FLIGHT RECORD: {name}")
        st.markdown(f"**DIVISION:** {division}")
        st.markdown(f"**STATUS:** <span style='color:{status_color}; font-weight:bold'>{status_text}</span>", 
                    unsafe_allow_html=True)
        st.divider()
        
        # Time grid
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 🛫 JAM DATANG")
            
            check_in = morning if morning else "❌ NOT RECORDED"
            if is_late:
                check_in += " ⚠️ **LATE**"
            st.info(f"**Jam Datang:** {check_in}")
            
            break_out_display = break_out if break_out else "❌ NOT RECORDED"
            st.write(f"**Siang 1:** {break_out_display}")
        
        with col2:
            st.markdown("#### 🛬 JAM PULANG")
            
            break_in_display = break_in if break_in else "❌ NOT RECORDED"
            st.write(f"**Siang 2:** {break_in_display}")
            
            check_out = evening if evening else "❌ NOT RECORDED"
            st.success(f"**Jam Pulang:** {check_out}")
        
        # Calculate work duration
        if morning and evening:
            duration = self.time_service.calculate_duration(morning, evening)
            if duration:
                formatted_duration = self.time_service.format_duration(duration)
                st.divider()
                st.markdown(f"### ⏱️ TOTAL DUTY TIME")
                st.metric("Duration", formatted_duration)
                
                # Add overtime indicator
                if duration.seconds / 3600 > 9:
                    st.warning("⚠️ Extended duty hours detected")
    
    def render_metric_cards(self, metrics: Dict[str, Any]) -> None:
        """Render key metrics in card format."""
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "👥 TOTAL PERSONNEL",
                metrics['total'],
                help="Total registered employees"
            )
        
        with col2:
            attendance_pct = f"{metrics['attendance_rate']:.1f}%"
            st.metric(
                "✅ ON DUTY",
                metrics['present'],
                delta=attendance_pct,
                help="Employees present today"
            )
        
        with col3:
            st.metric(
                "📝 PERMITS",
                metrics['permit'],
                help="Approved leaves and permits"
            )
        
        with col4:
            absent_delta = f"-{metrics['absent']}" if metrics['absent'] > 0 else "0"
            st.metric(
                "❌ ABSENT",
                metrics['absent'],
                delta=absent_delta,
                delta_color="inverse",
                help="Unexplained absences"
            )
    
    def render_anomaly_section(self, metrics: Dict[str, Any]) -> None:
        """Render anomaly detection section."""
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("#### ⚠️ TERLAMBAT")
            late_list = metrics['late_list']
            
            if late_list:
                with st.container(height=180):
                    for name, time_val in late_list:
                        st.markdown(
                            f"<div class='anomaly-box box-telat'>"
                            f"<span>{name}</span>"
                            f"<span style='font-weight:bold'>{time_val}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.success("✓ ALL ON TIME")
        
        with col2:
            st.markdown("#### ℹ️ SAKIT/IZIN")
            permit_list = metrics['permit_list']
            
            if permit_list:
                with st.container(height=180):
                    for name, reason in permit_list:
                        st.markdown(
                            f"<div class='anomaly-box box-izin'>"
                            f"<span>{name}</span>"
                            f"<span style='font-weight:bold'>{reason}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.info("○ NO PERMITS")
        
        with col3:
            st.markdown("#### ❌ YANG TIDAK HADIR")
            absent_list = metrics['absent_list']
            
            if absent_list:
                with st.container(height=180):
                    for name in absent_list:
                        st.markdown(
                            f"<div class='anomaly-box box-alpha'>"
                            f"<span>{name}</span>"
                            f"<span style='color:#e84118'>ABSENT</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.success("✓ FULL ATTENDANCE")
    
    def render_division_tabs(
        self, 
        df: pd.DataFrame, 
        status_dict: Dict[str, str],
        search_query: str = ""
    ) -> None:
        """Render division-based tabs with employee cards."""
        divisions = sorted(
            DivisionRegistry.get_all().items(),
            key=lambda x: x[1].priority
        )
        
        tab_names = [f"{div[1].icon} {div[0]}" for div in divisions]
        tabs = st.tabs(tab_names)
        
        for tab, (div_name, div_config) in zip(tabs, divisions):
            with tab:
                # Division stats
                st.markdown(f"**{div_config.description}**")
                
                # Get members strictly from config order (FORCED SORT ORDER)
                ordered_members = div_config.members
                
                # Filter DF for this division
                df_division = df[df[AppConstants.COL_EMPLOYEE_NAME].isin(ordered_members)].copy()
                
                # Calculate stats
                present = len(df_division[df_division['Pagi'] != ''])
                total = len(df_division)
                rate = (present / total * 100) if total > 0 else 0
                
                st.progress(rate / 100, text=f"Attendance: {present}/{total} ({rate:.1f}%)")
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Apply search filter if exists
                if search_query:
                    ordered_members = [
                        m for m in ordered_members 
                        if search_query.lower() in m.lower()
                    ]
                
                # Create grid
                cols = st.columns(AppConstants.CARDS_PER_ROW)
                card_count = 0
                
                # Loop through ORDERED list
                for member_name in ordered_members:
                    # Find the data row for this member
                    member_row = df_division[df_division[AppConstants.COL_EMPLOYEE_NAME] == member_name]
                    
                    if not member_row.empty:
                        row_data = member_row.iloc[0]
                        with cols[card_count % AppConstants.CARDS_PER_ROW]:
                            self.render_employee_card(row_data, status_dict)
                        card_count += 1
                
                if card_count == 0:
                    st.info(f"No personnel found in {div_name}")

    def render_login_page(self, login_callback):
        """Render tampilan login Clean Card."""
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.markdown("""
            <div class='login-card'>
                <h1 style="color: #1a1f2e; font-family: 'Rajdhani'; margin:0;">Weda Bay Airport</h1>
                <p style="color: #666; font-size: 0.9em; margin-bottom: 20px;">Personnel Monitoring System</p>
                <hr style="margin: 20px 0; border-top: 1px solid #eee;">
            </div>
            """, unsafe_allow_html=True)
            
            # --- PERBAIKAN DI SINI ---
            # Kita gunakan nama KUNCI MATI (Static) agar tidak mereset saat diklik
            with st.form(key="login_form_fixed_secure_v1", clear_on_submit=False):
                username = st.text_input("Username", placeholder="Masukkan username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password")
                submitted = st.form_submit_button("SIGN IN", use_container_width=True)
                
                if submitted:
                    login_callback(username, password)
            
            st.caption("Contact Admin for access issues.")

    def render_admin_dashboard(self, user_repo):
        """Render Dashboard khusus Admin."""
        st.markdown("## 🛡️ Admin Access Management")
        st.info("Kelola akses pengguna sistem.")
        
        users = user_repo.get_all_users()
        
        data_list = []
        for u, d in users.items():
            data_list.append({"Username": u, "Role": d['role'], "Status": d['status']})
        
        df_users = pd.DataFrame(data_list)
        
        col_table, col_edit = st.columns([2, 1])
        
        with col_table:
            st.dataframe(df_users, use_container_width=True, hide_index=True)
            
        with col_edit:
            st.markdown("### Edit Status")
            target = st.selectbox("Pilih User", df_users['Username'].unique())
            action = st.selectbox("Set Status", ["Active", "Inactive"])
            
            if st.button("Update Access", use_container_width=True):
                if user_repo.update_user_status(target, action):
                    st.success(f"Status {target} diubah ke {action}")
                    st.rerun()

# ================================================================================
# SECTION 7: VISUALIZATION LAYER
# ================================================================================

class ChartBuilder:
    """
    Builder class for creating interactive charts and visualizations.
    Uses Plotly for rich, interactive data visualization.
    """
    
    @staticmethod
    def create_attendance_pie_chart(metrics: Dict[str, Any]) -> go.Figure:
        """Create pie chart for attendance distribution."""
        labels = ['Present', 'Permit', 'Absent']
        values = [metrics['present'], metrics['permit'], metrics['absent']]
        colors = ['#4cd137', '#9c88ff', '#e84118']
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.4,
            marker=dict(colors=colors, line=dict(color='#0b0e14', width=2)),
            textfont=dict(size=14, color='white', family='Rajdhani'),
            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title=dict(
                text='Attendance Distribution',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=True,
            legend=dict(
                font=dict(color='white', family='Inter'),
                bgcolor='rgba(21, 25, 34, 0.8)'
            ),
            height=350
        )
        
        return fig
    
    @staticmethod
    def create_division_bar_chart(division_stats: Dict[str, Dict]) -> go.Figure:
        """Create bar chart for division-wise attendance."""
        divisions = []
        present = []
        absent = []
        colors = []
        
        for div_name, stats in sorted(division_stats.items(), key=lambda x: x[1]['rate'], reverse=True):
            divisions.append(div_name)
            present.append(stats['present'])
            absent.append(stats['absent'])
            colors.append(stats['color'])
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='Present',
            x=divisions,
            y=present,
            marker_color='#4cd137',
            text=present,
            textposition='auto',
        ))
        
        fig.add_trace(go.Bar(
            name='Absent',
            x=divisions,
            y=absent,
            marker_color='#e84118',
            text=absent,
            textposition='auto',
        ))
        
        fig.update_layout(
            title=dict(
                text='Division-wise Attendance',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            barmode='stack',
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(21, 25, 34, 0.6)',
            font=dict(color='white', family='Inter'),
            xaxis=dict(
                title='Division',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            yaxis=dict(
                title='Count',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            legend=dict(
                bgcolor='rgba(21, 25, 34, 0.8)'
            ),
            height=400
        )
        
        return fig
    
    @staticmethod
    def create_time_distribution_chart(df: pd.DataFrame) -> go.Figure:
        """Create histogram of arrival times."""
        if df.empty or 'Jam' not in df.columns:
            return go.Figure()
        
        fig = go.Figure()
        
        fig.add_trace(go.Histogram(
            x=df['Jam'],
            nbinsx=24,
            marker=dict(
                color='#00a8ff',
                line=dict(color='#0b0e14', width=1)
            ),
            name='Arrivals'
        ))
        
        # Add late threshold line
        fig.add_vline(
            x=7.083,  # 7:05 AM
            line_dash="dash",
            line_color="#e84118",
            annotation_text="Late Threshold",
            annotation_position="top right"
        )
        
        fig.update_layout(
            title=dict(
                text='Arrival Time Distribution',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(21, 25, 34, 0.6)',
            font=dict(color='white', family='Inter'),
            xaxis=dict(
                title='Hour of Day',
                gridcolor='rgba(255,255,255,0.1)',
                range=[0, 24]
            ),
            yaxis=dict(
                title='Number of Arrivals',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            height=350
        )
        
        return fig


# ================================================================================
# SECTION 8: APPLICATION CONTROLLER LAYER
# ================================================================================

class AttendanceController:
    """
    Main controller implementing MVC pattern.
    Orchestrates all business logic and UI rendering.
    """
    
    def __init__(self):
        # Initialize repositories
        self.attendance_repo = AttendanceRepository(DataSourceConfig.ATTENDANCE_SHEET_URL)
        self.status_repo = StatusRepository(DataSourceConfig.STATUS_SHEET_URL)
        
        # Initialize services
        self.attendance_service = AttendanceService(self.attendance_repo, self.status_repo)
        self.analytics_service = AnalyticsService(self.attendance_repo)
        
        # Initialize Auth
        self.auth_service = AuthService()
        
        # Initialize UI components
        self.component_renderer = ComponentRenderer()
        self.chart_builder = ChartBuilder()
        self.excel_exporter = ExcelExporter()

    def handle_login(self, username, password):
        success, result = self.auth_service.login(username, password)
        if success:
            st.session_state['logged_in'] = True
            st.session_state['user_info'] = result
            st.success("Login Berhasil!")
            st.rerun()
        else:
            st.error(f"Login Gagal: {result}")

    def handle_logout(self):
        """Logic saat tombol Logout ditekan."""
        # 1. Reset semua session state
        st.session_state['logged_in'] = False
        st.session_state['user_info'] = None
        
        # 2. Opsional: Hapus key lain biar bersih total
        if 'selected_date' in st.session_state:
            del st.session_state['selected_date']
            
        # 3. PAKSA RERUN SEKARANG JUGA
        st.rerun()

    def run(self):
        """Entry point logic."""
        # 1. Cek Session State
        if 'logged_in' not in st.session_state:
            st.session_state['logged_in'] = False
            
        # 2. Routing Page
        if not st.session_state['logged_in']:
            self.component_renderer.render_login_page(self.handle_login)
        else:
            user_role = st.session_state['user_info']['role']
            user_name = st.session_state['user_info']['name']
            
            # Sidebar User
            with st.sidebar:
                st.info(f"👤 Logged in as: **{user_name}**")
                # Unique Key for Logout to prevent Duplicate Widget
                # Kita pakai on_click agar fungsi dijalankan SEBELUM halaman refresh
                st.button("🚪 Logout", key="logout_btn", on_click=self.handle_logout, use_container_width=True)
                    self.handle_logout()
                st.markdown("---")
            
            # Logic Admin
            if user_role == 'admin':
                menu_selection = st.sidebar.radio(
                    "Menu Admin", 
                    ["📊 Monitoring Dashboard", "🛡️ User Management", "⚙️ Settings"]
                )
                
                if menu_selection == "🛡️ User Management":
                    self.component_renderer.render_admin_dashboard(self.auth_service.repo)
                    return 
                elif menu_selection == "⚙️ Settings":
                    render_settings_page()
                    return
            
            # Render Dashboard Absensi
            self.run_dashboard_content()

    def run_dashboard_content(self) -> None:
        """Main dashboard view."""
        # 1. Header
        st.markdown('<div class="brand-title">WedaBayAirport</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">REAL-TIME PERSONNEL MONITORING SYSTEM</div>', 
                    unsafe_allow_html=True)
        
        # 2. Check data availability
        df_attendance = self.attendance_repo.fetch()
        if df_attendance is None:
            st.error("⚠️ SYSTEM OFFLINE - Unable to connect to attendance database")
            st.stop()
        
        # 3. Filters & Date Selection
        col1, col2, col3 = st.columns([2, 3, 2])
        
        with col1:
            unfiltered_dates = df_attendance['Tanggal'].unique()
            clean_dates = [d for d in unfiltered_dates if pd.notna(d)]
            available_dates = sorted(clean_dates, reverse=True)
            if not available_dates:
                st.error("No attendance data available")
                st.stop()
            selected_date = st.date_input("📅 OPERATION DATE", value=available_dates[0])
        
        with col2:
            search_query = st.text_input("🔍 PERSONNEL SEARCH", placeholder="Search by name...")
        
        with col3:
            view_mode = st.selectbox("👁️ VIEW MODE", ["Cards", "Table", "Analytics"])
        
        st.markdown("---")
        
        # 4. Build report variables
        with st.spinner("🔄 Loading flight data..."):
            df_final, status_dict = self.attendance_service.build_complete_report(selected_date)
            metrics = self.attendance_service.calculate_metrics(df_final, status_dict)
        
        # 5. Metrics section UI
        self.component_renderer.render_metric_cards(metrics)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 6. Anomalies section UI
        with st.container():
            self.component_renderer.render_anomaly_section(metrics)
        
        st.markdown("---")
        
        # 7. EXPORT & REPORTS SECTION
        st.markdown("### 📤 EXPORT REPORTS")
        
        export_tab1, export_tab2 = st.tabs(["📄 Daily Report", "📅 Range Report"])

        # TAB 1: DOWNLOAD PER HARI
        with export_tab1:
            col_ex1, col_ex2 = st.columns([1, 1])
            with col_ex1:
                st.info(f"Download report for selected date: **{selected_date.strftime('%d %B %Y')}**")
                
                excel_file = self.excel_exporter.create_attendance_report(
                    df_final, status_dict, selected_date, metrics
                )
                
                st.download_button(
                    "📥 DOWNLOAD DAILY EXCEL",
                    data=excel_file,
                    file_name=f"Attendance_{selected_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_ex2:
                 if st.button("📊 VIEW ANALYTICS", use_container_width=True):
                    st.session_state['show_analytics'] = True

        # TAB 2: DOWNLOAD RANGE TANGGAL
        with export_tab2:
            st.write("Select a date range to generate a multi-sheet Excel report.")
            col_rng1, col_rng2 = st.columns(2)
            with col_rng1:
                start_date_input = st.date_input("Start Date", value=selected_date - timedelta(days=7))
            with col_rng2:
                end_date_input = st.date_input("End Date", value=selected_date)

            if st.button("📦 GENERATE RANGE REPORT", use_container_width=True):
                if start_date_input > end_date_input:
                    st.error("Error: Start Date must be before End Date")
                else:
                    with st.spinner(f"Generating report from {start_date_input} to {end_date_input}..."):
                        range_data_map = {}
                        current_loop_date = start_date_input
                        
                        while current_loop_date <= end_date_input:
                            try:
                                day_df, day_status = self.attendance_service.build_complete_report(current_loop_date)
                                range_data_map[current_loop_date] = (day_df, day_status)
                            except Exception:
                                pass 
                            current_loop_date += timedelta(days=1)
                        
                        if range_data_map:
                            range_excel = self.excel_exporter.create_range_report(range_data_map)
                            st.success("✅ Report Generated Successfully!")
                            st.download_button(
                                label=f"📥 DOWNLOAD RANGE REPORT ({start_date_input} - {end_date_input})",
                                data=range_excel,
                                file_name=f"Attendance_Range_{start_date_input}_{end_date_input}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.warning("No data found for the selected range.")

        st.markdown("---")
        
        # 8. View modes Logic
        if view_mode == "Cards":
            st.markdown("### 📋 PERSONNEL ROSTER")
            self.component_renderer.render_division_tabs(df_final, status_dict, search_query)
        
        elif view_mode == "Table":
            self._render_table_view(df_final, status_dict)
        
        elif view_mode == "Analytics":
            self._render_analytics_view(df_final, status_dict, metrics, selected_date)
        
        # Additional analytics modal
        if st.session_state.get('show_analytics', False):
            with st.expander("📈 ADVANCED ANALYTICS", expanded=True):
                self._render_analytics_view(df_final, status_dict, metrics, selected_date)

    def _render_table_view(self, df: pd.DataFrame, status_dict: Dict[str, str]) -> None:
        """Render table view of attendance."""
        st.markdown("### 📊 DETAILED ATTENDANCE TABLE")
        
        df_display = df.copy()
        
        df_display['Division'] = df_display[AppConstants.COL_EMPLOYEE_NAME].apply(
            lambda x: DivisionRegistry.find_by_member(x).code 
            if DivisionRegistry.find_by_member(x) else "N/A"
        )
        
        df_display['Status'] = df_display[AppConstants.COL_EMPLOYEE_NAME].apply(
            lambda x: status_dict.get(x, "")
        )
        
        df_display = df_display.rename(columns={
            'Pagi': 'Jam Datang',
            'Siang_1': 'Siang 1',
            'Siang_2': 'Siang 2',
            'Sore': 'Jam Pulang'
        })
        
        display_columns = [
            AppConstants.COL_EMPLOYEE_NAME, 
            'Division', 
            'Jam Datang', 
            'Siang 1', 
            'Siang 2', 
            'Jam Pulang', 
            'Status'
        ]
        
        final_cols = [c for c in display_columns if c in df_display.columns]
        df_display = df_display[final_cols]
        
        def highlight_late(val):
            if isinstance(val, str) and ':' in val:
                if TimeService.is_late(val):
                    return 'color: #e84118; font-weight: bold'
            return ''
        
        try:
            st.dataframe(
                df_display.style.applymap(highlight_late, subset=['Jam Datang']),
                use_container_width=True,
                height=600,
                hide_index=True
            )
        except Exception:
            st.dataframe(df_display, use_container_width=True, height=600)
        
        csv = df_display.to_csv(index=False)
        st.download_button(
            "💾 DOWNLOAD CSV",
            data=csv,
            file_name=f"attendance_table_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    def _render_analytics_view(
        self, 
        df: pd.DataFrame, 
        status_dict: Dict[str, str],
        metrics: Dict[str, Any],
        selected_date: datetime.date
    ) -> None:
        """Render advanced analytics view."""
        st.markdown("### 📈 ANALYTICS DASHBOARD")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            pie_chart = self.chart_builder.create_attendance_pie_chart(metrics)
            st.plotly_chart(pie_chart, use_container_width=True)
        
        with chart_col2:
            division_stats = self.analytics_service.get_division_statistics(selected_date)
            bar_chart = self.chart_builder.create_division_bar_chart(division_stats)
            st.plotly_chart(bar_chart, use_container_width=True)
        
        df_attendance_day = self.attendance_service.get_attendance_for_date(selected_date)
        if df_attendance_day is not None and not df_attendance_day.empty:
            time_dist_chart = self.chart_builder.create_time_distribution_chart(df_attendance_day)
            st.plotly_chart(time_dist_chart, use_container_width=True)
        
        st.markdown("### 💡 KEY INSIGHTS")
        
        insight_col1, insight_col2, insight_col3 = st.columns(3)
        
        with insight_col1:
            st.metric(
                "Punctuality Rate",
                f"{metrics['punctuality_rate']:.1f}%",
                help="Percentage of on-time arrivals"
            )
        
        with insight_col2:
            best_division = max(
                division_stats.items(),
                key=lambda x: x[1]['rate']
            )
            st.metric(
                "Top Division",
                best_division[0],
                f"{best_division[1]['rate']:.1f}%"
            )
        
        with insight_col3:
            avg_late_time = "N/A"
            if metrics['late_list']:
                late_times = [
                    datetime.strptime(t, '%H:%M').time() 
                    for _, t in metrics['late_list']
                ]
                avg_minutes = sum(
                    t.hour * 60 + t.minute for t in late_times
                ) / len(late_times)
                avg_late_time = f"{int(avg_minutes // 60):02d}:{int(avg_minutes % 60):02d}"
            
            st.metric(
                "Avg Late Arrival",
                avg_late_time,
                help="Average time of late arrivals"
            )

    def run_report_form(self) -> None:
        """Report submission view."""
        st.markdown('<div class="brand-title">MANUAL REPORTING</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">SUBMIT PERMITS & LEAVE REQUESTS</div>', 
                    unsafe_allow_html=True)
        
        st.info("📝 Use the form below to submit permit requests, sick leaves, or other attendance modifications.")
        
        if "PASTE_LINK" in DataSourceConfig.REPORT_FORM_URL:
            st.warning("⚠️ Google Form URL not configured. Please contact system administrator.")
        else:
            components.iframe(
                DataSourceConfig.REPORT_FORM_URL,
                height=1200,
                scrolling=True
            )

# ================================================================================
# SECTION 9: ADDITIONAL FEATURES & UTILITIES
# ================================================================================

class ConfigurationManager:
    """
    Application configuration management.
    Handles user preferences and settings.
    """
    
    @staticmethod
    def initialize_session_state() -> None:
        """Initialize Streamlit session state variables."""
        if 'show_analytics' not in st.session_state:
            st.session_state['show_analytics'] = False
        
        if 'selected_date' not in st.session_state:
            st.session_state['selected_date'] = datetime.now().date()
        
        if 'view_mode' not in st.session_state:
            st.session_state['view_mode'] = 'Cards'
        
        if 'theme' not in st.session_state:
            st.session_state['theme'] = 'dark'
        
        if 'notifications_enabled' not in st.session_state:
            st.session_state['notifications_enabled'] = True


def configure_page() -> None:
    """Configure Streamlit page settings."""
    st.set_page_config(
        page_title=AppConstants.APP_TITLE,
        page_icon=AppConstants.APP_ICON,
        layout=AppConstants.LAYOUT_MODE,
        initial_sidebar_state=AppConstants.SIDEBAR_STATE,
        menu_items={
            'Get Help': 'https://wedabay.airport/support',
            'Report a bug': 'https://wedabay.airport/bug-report',
            'About': f"""
            # {AppConstants.APP_TITLE}
            Version {AppConstants.APP_VERSION}
            
            Enterprise Personnel Monitoring System
            
            © 2025 {AppConstants.COMPANY_NAME}
            """
        }
    )


def render_settings_page() -> None:
    """Render settings and configuration page."""
    st.markdown('<div class="brand-title">SETTINGS</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">SYSTEM CONFIGURATION</div>', 
                unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["⚙️ General", "👥 Division Management", "📊 Reports"])
    
    with tab1:
        st.markdown("### General Settings")
        st.checkbox("Enable Notifications", value=True)
        st.selectbox("Default View Mode", ["Cards", "Table", "Analytics"])
        st.slider("Cards Per Row", 2, 6, 4)
        st.number_input("Late Threshold (minutes)", 0, 60, 5)
    
    with tab2:
        st.markdown("### Division Management")
        divisions = DivisionRegistry.get_all()
        for div_name, div_config in divisions.items():
            with st.expander(f"{div_config.icon} {div_name} ({div_config.code})"):
                st.markdown(f"**Description:** {div_config.description}")
                st.markdown(f"**Color:** {div_config.color}")
                st.markdown(f"**Total Members:** {len(div_config.members)}")
                st.markdown("**Members:**")
                for member in div_config.members:
                    st.caption(f"• {member}")
    
    with tab3:
        st.markdown("### Report Configuration")
        st.selectbox("Default Export Format", ["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)", "PDF (.pdf)"])
        st.checkbox("Include Charts in Reports", value=True)


# ================================================================================
# SECTION 10: MAIN APPLICATION ENTRY POINT
# ================================================================================

def main() -> None:
    """
    Main application entry point.
    Orchestrates the entire application flow.
    """
    # 1. Konfigurasi Halaman & State
    configure_page()
    ConfigurationManager.initialize_session_state()
    initialize_divisions()
    
    # 2. Apply CSS Theme
    ThemeManager.apply_global_styles()
    
    # 3. Inisialisasi Controller Baru
    controller = AttendanceController()
    
    # 4. JALANKAN APLIKASI (Login -> Dashboard)
    controller.run()


if __name__ == "__main__":
    main()
