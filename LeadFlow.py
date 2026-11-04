# =============================================================================
#  LeadFlow — India Google Maps Lead Generation Tool
#  Production-Ready v4 | Auto-exports to Excel + Google Sheets
#
#  pip install selenium webdriver-manager pandas openpyxl gspread google-auth
#
#  GOOGLE SHEETS SETUP:
#    1. Go to https://console.cloud.google.com/ → Create project
#    2. Enable "Google Sheets API" and "Google Drive API"
#    3. Create a Service Account → Download JSON key → save as "credentials.json"
#       in the same folder as this script
#    4. Share your target Google Sheet with the service-account email (Editor role)
#    5. Enter the Sheet ID (from its URL) in the app before starting
#
#  FILTERS (strict enforcement):
#    - Min Rating / Max Rating sliders: e.g. parlour below 3 stars → Max=3
#    - Phone Only / Website Only checkboxes
#    - Location is STRICT: only the exact area you select is used in queries;
#      scraper validates address contains the selected city/area keywords.
#    - No lat/lng columns anywhere.
# =============================================================================

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import time
import os
import json
import random
import sys
import re
import queue
import shutil
import platform
import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple, Set

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

# =============================================================================
#  GOOGLE SHEETS HELPER (graceful fallback if gspread not installed)
# =============================================================================

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_gspread_client(creds_path: str):
    """Return an authorised gspread client, or None on failure."""
    if not GSPREAD_AVAILABLE:
        return None, "gspread / google-auth not installed. Run: pip install gspread google-auth"
    if not os.path.exists(creds_path):
        return None, f"credentials.json not found at: {creds_path}"
    try:
        creds = Credentials.from_service_account_file(creds_path, scopes=SHEETS_SCOPES)
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, str(e)


def push_to_sheets(records: list, sheet_id: str, creds_path: str,
                   worksheet_name: str, log_q=None):
    """
    Append records to the given Google Sheet.
    Creates the worksheet if it doesn't exist.
    Returns (rows_added, error_message).
    """
    if not records:
        return 0, None

    client, err = get_gspread_client(creds_path)
    if err:
        log(f"[Sheets] Error: {err}", log_q)
        return 0, err

    try:
        sh = client.open_by_key(sheet_id)
    except Exception as e:
        return 0, f"Cannot open sheet: {e}"

    # Open or create worksheet
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=5000, cols=len(COLUMNS))
        # Write header
        ws.append_row(COLUMNS)

    # Check if sheet is empty (needs header)
    try:
        existing_values = ws.get_all_values()
        if not existing_values:
            ws.append_row(COLUMNS)
            existing_values = [COLUMNS]
    except Exception:
        existing_values = [COLUMNS]

    # Build rows to append (order must match COLUMNS)
    rows_to_add = []
    for rec in records:
        row = [str(rec.get(col, "N/A") or "N/A") for col in COLUMNS]
        rows_to_add.append(row)

    if not rows_to_add:
        return 0, None

    try:
        ws.append_rows(rows_to_add, value_input_option="USER_ENTERED")
        log(f"[Sheets] ✅ Pushed {len(rows_to_add)} rows to Google Sheets.", log_q)
        return len(rows_to_add), None
    except Exception as e:
        return 0, str(e)


# =============================================================================
#  ALL-INDIA LOCATION DATABASE
# =============================================================================

INDIA_LOCATIONS = {
    "Maharashtra": {
        "Mumbai": [
            "Andheri West", "Andheri East", "Bandra West", "Bandra East",
            "Dadar", "Borivali", "Malad", "Kandivali", "Goregaon",
            "Santacruz", "Vile Parle", "Jogeshwari", "Lokhandwala",
            "Versova", "Juhu", "Powai", "Vikhroli", "Ghatkopar",
            "Kurla", "Chembur", "Mulund", "Thane City", "Kalyan",
            "Dombivli", "Navi Mumbai", "Worli", "Lower Parel", "Parel",
            "Wadala", "Sion", "Dharavi", "Colaba", "Fort", "Nariman Point",
            "Marine Lines", "Churchgate", "Malabar Hill", "Byculla",
        ],
        "Palghar": [
            "Virar", "Nalasopara", "Vasai", "Palghar", "Boisar",
            "Dahanu", "Mira Road", "Sapahle", "Naigaon",
        ],
        "Pune": [
            "Hinjewadi", "Wakad", "Baner", "Balewadi", "Kharadi",
            "Hadapsar", "Viman Nagar", "Kalyani Nagar", "Koregaon Park",
            "Aundh", "Kothrud", "Deccan", "Shivajinagar", "Camp",
            "Pimpri", "Chinchwad", "Nigdi", "Akurdi", "Swargate",
            "Kondhwa", "Bibwewadi", "Fatima Nagar", "NIBM", "Magarpatta",
            "Wagholi", "Bhosari", "Alandi Road", "Dhanori", "Lohegaon",
        ],
        "Nagpur": [
            "Dharampeth", "Sitabuldi", "Sadar", "Civil Lines", "Ramdaspeth",
            "Gandhibagh", "Nandanvan", "Manewada", "Wardha Road", "Amravati Road",
            "Hingna Road", "Kamptee Road", "MIDC Nagpur", "Pratap Nagar",
        ],
        "Nashik": [
            "College Road", "Gangapur Road", "Mumbai Naka", "Nashik Road",
            "Indira Nagar Nashik", "Panchavati", "Mahatma Nagar", "Cidco Nashik",
        ],
        "Thane": [
            "Ghodbunder Road", "Majiwada", "Manpada", "Kolshet",
            "Mira Road", "Bhayander", "Vasai", "Virar", "Nalasopara",
        ],
        "Aurangabad": [
            "CIDCO Aurangabad", "Garkheda", "Waluj", "Harsul", "Osmanpura",
            "Kranti Chowk", "TV Centre", "Begumpura",
        ],
        "Kolhapur": [
            "Rajarampuri", "Shahupuri", "Shivaji Park Kolhapur",
            "Kasba Bawada", "Tarabai Park",
        ],
        "Solapur": [
            "Hotgi Road", "Vijapur Road", "Akkalkot Road", "Shivaji Nagar Solapur",
        ],
        "Navi Mumbai": [
            "Vashi", "Nerul", "Kharghar", "Panvel", "Belapur", "Airoli",
            "Ghansoli", "Kopar Khairane", "Turbhe", "Sanpada",
        ],
        "Amravati": ["Rajkamal Chowk", "Jaistambh Square", "Vadar Ward"],
        "Latur": ["Ausa Road", "Udgir Road", "Nanded Road Latur"],
        "Dhule": ["Station Road Dhule", "Civil Lines Dhule"],
        "Jalgaon": ["MIDC Jalgaon", "Ring Road Jalgaon"],
        "Satara": ["Shivaji Chowk Satara", "Koregaon Satara"],
    },
    "Delhi": {
        "South Delhi": [
            "Saket", "Hauz Khas", "Green Park", "Greater Kailash",
            "Lajpat Nagar", "Defence Colony", "Vasant Kunj", "Vasant Vihar",
            "Malviya Nagar", "Panchsheel", "Safdarjung", "Nehru Place",
        ],
        "North Delhi": [
            "Rohini", "Pitampura", "Shalimar Bagh", "Model Town",
            "Civil Lines Delhi", "Kamla Nagar", "GTB Nagar", "Ashok Vihar",
        ],
        "East Delhi": [
            "Preet Vihar", "Laxmi Nagar", "Mayur Vihar", "Patparganj",
            "Shahdara", "Vivek Vihar", "Geeta Colony",
        ],
        "West Delhi": [
            "Rajouri Garden", "Janakpuri", "Tilak Nagar", "Vikaspuri",
            "Dwarka", "Palam", "Uttam Nagar",
        ],
        "Central Delhi": [
            "Connaught Place", "Paharganj", "Karol Bagh", "Patel Nagar",
            "Rajendra Nagar", "Daryaganj",
        ],
        "New Delhi": [
            "Gole Market", "RK Puram", "Lodhi Colony", "Khan Market", "Lodi Road",
        ],
        "Noida": [
            "Sector 18 Noida", "Sector 62", "Sector 63", "Sector 135",
            "Greater Noida", "Noida Extension",
        ],
        "Gurgaon": [
            "DLF Phase 1", "DLF Phase 2", "Cyber City", "MG Road Gurgaon",
            "Sohna Road", "Golf Course Road", "Sector 29 Gurgaon",
        ],
        "Faridabad": ["Sector 15 Faridabad", "NIT Faridabad", "Ballabhgarh"],
    },
    "Gujarat": {
        "Ahmedabad": [
            "Navrangpura", "Vastrapur", "Satellite", "Prahlad Nagar",
            "Bodakdev", "Ambawadi", "Paldi", "Maninagar", "Nikol",
            "Gota", "Chandkheda", "SG Road Ahmedabad", "CG Road",
            "Nehru Nagar", "Thaltej", "Science City Road",
        ],
        "Surat": [
            "Adajan", "Vesu", "Athwa", "Katargam", "Udhna",
            "Varachha", "Piplod", "Pal Surat", "Rander",
            "Dumas Road", "Althan", "Bhatar",
        ],
        "Vadodara": [
            "Alkapuri", "Gotri", "Waghodia Road", "Fatehgunj",
            "Karelibaug", "Manjalpur", "Harni Road", "Old Padra Road",
        ],
        "Rajkot": [
            "Kalawad Road", "Raiya Road", "150 Feet Ring Road Rajkot",
            "Gondal Road", "Mavdi", "Sadhuvasvani Road",
        ],
        "Gandhinagar": ["Sector 7 Gandhinagar", "Sector 11", "Sector 21"],
        "Anand": ["Vallabh Vidyanagar", "Anand City"],
        "Bharuch": ["Bharuch City"],
        "Mehsana": ["Mehsana City"],
        "Junagadh": ["Kalwa Chowk Junagadh"],
        "Jamnagar": ["Jamnagar City", "Lal Bungalow Jamnagar"],
    },
    "Karnataka": {
        "Bengaluru": [
            "Koramangala", "Indiranagar", "HSR Layout", "BTM Layout",
            "Whitefield", "Electronic City", "Marathahalli", "Bellandur",
            "Sarjapur Road", "Jayanagar", "JP Nagar", "Basavanagudi",
            "Rajajinagar", "Yeshwanthpur", "Malleshwaram", "Hebbal",
            "Yelahanka", "Bannerghatta Road", "Kanakapura Road",
            "Domlur", "Bommanahalli", "Banashankari", "Padmanabhanagar",
            "MG Road Bengaluru", "Brigade Road", "Commercial Street",
        ],
        "Mysuru": [
            "Vijayanagar Mysuru", "Kuvempunagar", "Jayalakshmipuram",
            "Devaraja Mohalla", "Hebbal Mysuru", "N R Mohalla",
        ],
        "Mangaluru": ["Hampankatta", "Kadri", "Bejai", "Kankanady", "Bondel"],
        "Hubballi": ["Vidyanagar Hubballi", "Keshwapur", "Gokul Road"],
        "Belagavi": ["Camp Belagavi", "Tilakwadi", "Shivaji Nagar Belagavi"],
        "Davangere": ["PJ Extension", "Basha Nagar Davangere"],
        "Shivamogga": ["Shivamogga City"],
        "Tumakuru": ["Tumakuru City"],
        "Udupi": ["Udupi City", "Manipal"],
        "Kolar": ["KGF", "Bangarpet"],
    },
    "Tamil Nadu": {
        "Chennai": [
            "Anna Nagar", "T Nagar", "Adyar", "Velachery", "Tambaram",
            "Porur", "Ambattur", "Perambur", "Villivakkam", "Kodambakkam",
            "Nungambakkam", "Egmore", "Mylapore", "Besant Nagar",
            "Thiruvanmiyur", "Sholinganallur", "OMR Chennai", "ECR Chennai",
            "Chromepet", "Pallavaram", "Guindy", "Saidapet",
        ],
        "Coimbatore": [
            "RS Puram", "Saibaba Colony", "Gandhipuram", "Peelamedu",
            "Singanallur", "Ramanathapuram Coimbatore", "Ganapathy",
        ],
        "Madurai": [
            "Anna Nagar Madurai", "KK Nagar Madurai", "Teppakulam",
            "Vilangudi", "Thirunagar",
        ],
        "Salem": ["Fairlands", "Race Course", "Karumanthai Salem"],
        "Tiruppur": ["Town Area Tiruppur", "Tiruppur City"],
        "Erode": ["Erode City"],
        "Trichy": ["Trichy City", "Cantonment Trichy"],
    },
    "Telangana": {
        "Hyderabad": [
            "Banjara Hills", "Jubilee Hills", "Kondapur", "Madhapur",
            "Secunderabad", "Ameerpet", "Hitech City", "HITEC Hyderabad",
            "Somajiguda", "Begumpet", "Mehdipatnam", "Kukatpally", "Miyapur",
        ],
        "Vikarabad": ["Vikarabad City"],
        "Medchal": ["Medchal City"],
    },
    "Rajasthan": {
        "Jaipur": [
            "C Scheme Jaipur", "Malviya Nagar Jaipur", "Bani Park",
            "Raja Park", "Gopalpura", "New Sanganer Road",
            "M I Road Jaipur", "Ashok Nagar Jaipur",
        ],
        "Jodhpur": ["High Court Jodhpur", "Jodhpur City"],
        "Udaipur": ["Udaipur City"],
    },
    "Uttar Pradesh": {
        "Lucknow": [
            "Hazratganj Lucknow", "Gomti Nagar", "Indira Nagar Lucknow",
            "Aliganj Lucknow", "Charbagh Lucknow",
        ],
        "Kanpur": ["Kanpur City"],
        "Varanasi": ["Varanasi City"],
        "Agra": ["Agra City"],
    },
    "West Bengal": {
        "Kolkata": [
            "Ballygunge", "Alipore", "Bhowanipore", "Jadavpur",
            "Salt Lake", "New Town Kolkata", "Bidhannagar",
            "Exchange Place Kolkata", "B B D Bagh",
        ],
        "Darjeeling": ["Darjeeling City"],
    },
    "Punjab": {
        "Chandigarh": ["Sector 17 Chandigarh", "Sector 35", "Sector 22"],
        "Mohali": ["Mohali City"],
        "Jalandhar": ["Jalandhar City"],
    },
    "Haryana": {
        "Faridabad": ["Sector 37 Faridabad", "Sector 80 Faridabad"],
    },
}

BUSINESS_TYPES = [
    "beauty parlour", "salon", "spa", "gym", "yoga center",
    "dance classes", "coaching classes", "tuition classes",
    "restaurant", "cafe", "bakery", "dhaba", "fast food",
    "builder", "real estate agent", "property dealer",
    "real estate developer", "construction company",
    "hospital", "clinic", "doctor", "dentist", "physiotherapy",
    "pharmacy", "medical store", "pathology lab",
    "hotel", "guest house", "resort", "lodge",
    "travel agency", "tour operator",
    "car dealer", "automobile", "car service center",
    "mobile shop", "electronics shop", "computer shop",
    "boutique", "clothing store", "tailor", "garment shop",
    "lawyer", "advocate", "legal services",
    "CA", "chartered accountant", "tax consultant",
    "interior designer", "architect", "civil contractor",
    "school", "college", "institute", "academy",
    "event management", "wedding planner", "caterer",
    "printing press", "packaging", "digital printing",
    "pest control", "cleaning services", "home services",
    "insurance agent", "financial advisor", "stock broker",
    "grocery store", "supermarket", "departmental store",
    "hardware store", "plumbing", "electrical shop",
    "jewelry shop", "watch shop", "gift shop",
    "photography", "videography", "studio",
    "gym equipment", "sports shop", "fitness center",
    "driving school", "courier service", "logistics",
    "food delivery", "tiffin service", "meal prep",
    "nursery", "plants", "garden center",
    "pet shop", "veterinary", "pet care",
    "bank", "ATM", "finance company",
]

# =============================================================================
#  CONSTANTS
# =============================================================================

PROGRESS_FILE  = "lead_gen_progress.json"
LOG_QUEUE      = queue.Queue()
MAX_SCROLL     = 30
SCROLL_PAUSE   = 2.2
CARD_TIMEOUT   = 15
DETAIL_TIMEOUT = 12
CARD_RETRIES   = 3
MAX_DRIVER_RESTARTS = 5

DARK   = "#13141f"
PANEL  = "#1a1c2b"
CARD   = "#222538"
ACCENT = "#2e334d"
WHITE  = "#f1f5f9"
MUTED  = "#94a3b8"
GREEN  = "#10b981"
RED    = "#ef4444"
BLUE   = "#3b82f6"
YELLOW = "#f59e0b"
ORANGE = "#f97316"

# Output columns — NO latitude/longitude
COLUMNS = [
    "Business Name", "Category", "Phone Number", "WhatsApp",
    "Website", "Email", "Full Address", "City", "State",
    "Area", "Pincode", "Rating", "Reviews Count",
    "Google Maps URL",
]

# =============================================================================
#  CROSS-PLATFORM CHROME DETECTION
# =============================================================================

def find_chrome_binary() -> Optional[str]:
    system = platform.system()
    candidates = []
    if system == "Windows":
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
    elif system == "Darwin":
        candidates = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ]
    else:
        candidates = [
            "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser", "/usr/bin/chromium",
            shutil.which("google-chrome") or "",
            shutil.which("chromium-browser") or "",
        ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


# =============================================================================
#  HARDENED ANTI-DETECTION CHROME LAUNCH
# =============================================================================

def start_driver(log_q=None) -> webdriver.Chrome:
    opts = Options()
    chrome_bin = find_chrome_binary()
    if chrome_bin:
        opts.binary_location = chrome_bin
        log(f"Chrome binary: {chrome_bin}", log_q)
    else:
        log("Chrome binary not found at known paths; using system default.", log_q)

    opts.add_argument("--start-maximized")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--lang=en-US,en;q=0.9")
    opts.add_argument("--window-size=1366,768")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-infobars")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--allow-running-insecure-content")

    ua_versions = ["122.0.0.0", "123.0.0.0", "124.0.0.0", "125.0.0.0"]
    ua = (
        f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        f"AppleWebKit/537.36 (KHTML, like Gecko) "
        f"Chrome/{random.choice(ua_versions)} Safari/537.36"
    )
    opts.add_argument(f"user-agent={ua}")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    })

    svc = Service(ChromeDriverManager().install())
    svc.log_path = os.devnull if platform.system() != "Windows" else "NUL"

    try:
        driver = webdriver.Chrome(service=svc, options=opts)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
                window.chrome = { runtime: {} };
            """
        })
        log("Chrome launched successfully.", log_q)
        return driver
    except Exception as e:
        log(f"Chrome startup failed: {e}", log_q)
        raise


# =============================================================================
#  UTILITY HELPERS
# =============================================================================

def human_sleep(mn: float = 1.5, mx: float = 3.5):
    time.sleep(random.uniform(mn, mx))


def log(msg: str, log_q: queue.Queue = None):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    if log_q:
        log_q.put(line)


def smart_wait(driver, css: str, timeout: int = DETAIL_TIMEOUT):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css))
        )
        return True
    except TimeoutException:
        return False


# =============================================================================
#  PROGRESS HELPERS
# =============================================================================

def load_progress() -> dict:
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"done": [], "total_saved": 0}


def save_progress(prog: dict):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(prog, f, indent=2, ensure_ascii=False)


def load_existing_keys(filepath: str) -> set:
    keys = set()
    if os.path.exists(filepath):
        try:
            df = pd.read_excel(filepath)
            for _, row in df.iterrows():
                n = str(row.get("Business Name", "")).strip().lower()
                p = str(row.get("Phone Number", "")).strip()
                keys.add((n, p))
        except Exception:
            pass
    return keys


# =============================================================================
#  STRICT QUERY GENERATOR
#  Queries include ONLY the user-selected area; neighbouring areas excluded.
# =============================================================================

def generate_queries(state: str, city: str, areas: list, biz_type: str) -> list:
    """
    Strict location queries: each query is pinned to the exact area(s) the
    user selected. We do NOT add generic city-wide queries when a specific
    area is chosen, to prevent results from neighbouring localities.
    """
    queries = []
    for area in areas:
        if area:
            queries.append(f"{biz_type} in {area} {city} {state}")
            queries.append(f"{biz_type} near {area} {city}")
            queries.append(f"best {biz_type} in {area} {city}")
        else:
            # No specific area → city-level only
            queries.append(f"{biz_type} in {city} {state}")
            queries.append(f"best {biz_type} in {city} {state}")

    seen, unique = set(), []
    for q in queries:
        if q not in seen:
            seen.add(q)
            unique.append(q)
    return unique


# =============================================================================
#  CAPTCHA DETECTION
# =============================================================================

def detect_captcha(driver) -> bool:
    src = driver.page_source.lower()
    return any(k in src for k in ["captcha", "recaptcha", "unusual traffic", "not a robot"])


def handle_captcha(driver, log_q=None):
    log("⚠ CAPTCHA detected! Please solve it in the browser window.", log_q)
    log("  Waiting up to 120 seconds for manual resolution...", log_q)
    for _ in range(120):
        time.sleep(1)
        if not detect_captcha(driver):
            log("✅ Captcha resolved. Continuing...", log_q)
            return True
    log("⏳ Captcha wait timed out. Skipping query.", log_q)
    return False


# =============================================================================
#  ROBUST SCROLL — ActionChains + stall detection
# =============================================================================

def scroll_results(driver, needed_leads: int, log_q=None):
    from selenium.webdriver.common.keys import Keys
    log("  Scrolling results panel...", log_q)

    # Aim for enough candidate cards so filters have data to evaluate
    min_cards_target = max(needed_leads * 3, 25)

    FEED_SELECTORS = [
        "div[role='feed']",
        "div.m6QErb.DxyBCb",
        "div.m6QErb.WNBkOb",
        "div.m6QErb",
    ]
    scrollable = None
    for sel in FEED_SELECTORS:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                scrollable = el
                break
        except NoSuchElementException:
            continue

    if not scrollable:
        log("  ⚠ Feed container not found; attempting window scroll.", log_q)

    last_count = 0
    stall_streak = 0
    CARD_SELECTORS = ["a.hfpxzc", "div.Nv2PK", "div[data-result-index]"]

    for attempt in range(MAX_SCROLL):
        cards = []
        for cs in CARD_SELECTORS:
            cards = driver.find_elements(By.CSS_SELECTOR, cs)
            if cards:
                break
        count = len(cards)
        log(f"  Scroll {attempt+1}/{MAX_SCROLL} | Cards visible: {count}", log_q)

        if count >= min_cards_target:
            log(f"  Candidate card pool reached ({count} cards).", log_q)
            break

        try:
            driver.find_element(
                By.XPATH,
                "//*[contains(text(),\"You've reached the end\") or "
                "contains(text(),'end of results') or "
                "contains(text(),\"You've reached\")]"
            )
            log("  End of results reached.", log_q)
            break
        except NoSuchElementException:
            pass

        scrolled = False
        if scrollable:
            try:
                driver.execute_script(
                    "arguments[0].scrollTop = arguments[0].scrollHeight", scrollable
                )
                scrolled = True
            except Exception:
                pass

        if not scrolled:
            try:
                ActionChains(driver).move_to_element(
                    scrollable or driver.find_element(By.TAG_NAME, "body")
                ).send_keys(Keys.PAGE_DOWN).perform()
                scrolled = True
            except Exception:
                pass

        if not scrolled:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        time.sleep(SCROLL_PAUSE + random.uniform(-0.3, 0.5))

        if count == last_count:
            stall_streak += 1
            if stall_streak >= 3:
                log("  No new cards after 3 scroll cycles.", log_q)
                break
        else:
            stall_streak = 0
        last_count = count


# =============================================================================
#  9-METHOD PHONE EXTRACTION WITH VALIDATION
# =============================================================================

_PHONE_RE  = re.compile(r"(?:\+91[\s\-]?)?[6-9]\d{9}|0\d{2,4}[\s\-]\d{6,8}")
_MOBILE_RE = re.compile(r"[6-9]\d{9}")


def _clean_phone(raw: str) -> str:
    raw = raw.strip().replace("\u200b", "").replace("\xa0", " ")
    return raw


def extract_phone_enhanced(driver, log_q=None) -> str:
    for _pass in range(3):
        try:
            for el in driver.find_elements(
                By.XPATH, "//button[@data-item-id and contains(@data-item-id,'phone')]"
            ):
                t = _clean_phone(el.text)
                if _PHONE_RE.search(t):
                    return t
        except Exception:
            pass

        try:
            for link in driver.find_elements(By.XPATH, "//a[contains(@href,'tel:')]"):
                href = (link.get_attribute("href") or "").replace("tel:", "").strip()
                if href and re.search(r"\d{7,}", href):
                    return _clean_phone(href)
        except Exception:
            pass

        try:
            for el in driver.find_elements(
                By.XPATH, "//button[contains(@aria-label,'Phone') or contains(@aria-label,'phone')]"
            ):
                lbl = el.get_attribute("aria-label") or el.text
                lbl = _clean_phone(lbl)
                if _PHONE_RE.search(lbl):
                    return lbl
        except Exception:
            pass

        try:
            for el in driver.find_elements(
                By.XPATH, "//button[contains(@aria-label,'Call') or contains(@aria-label,'call')]"
            ):
                lbl = el.get_attribute("aria-label") or el.text
                lbl = _clean_phone(lbl)
                if _PHONE_RE.search(lbl):
                    return lbl
        except Exception:
            pass

        try:
            for el in driver.find_elements(By.XPATH, "//*[@data-tooltip]"):
                tip = el.get_attribute("data-tooltip") or ""
                if _PHONE_RE.search(tip):
                    return _clean_phone(tip)
        except Exception:
            pass

        try:
            for sel in [
                "span.UsdlK", "div.rogA2c", "span.LrzXr",
                "div[data-item-id] span", "div.Io6YTe"
            ]:
                for el in driver.find_elements(By.CSS_SELECTOR, sel):
                    t = _clean_phone(el.text)
                    if _PHONE_RE.search(t):
                        return t
        except Exception:
            pass

        try:
            els = driver.find_elements(
                By.XPATH,
                "//*[contains(@class,'fontBodyMedium') or contains(@class,'rogA2c')]"
            )
            for el in els:
                t = _clean_phone(el.text)
                if _PHONE_RE.search(t):
                    return t
        except Exception:
            pass

        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            hits = _PHONE_RE.findall(body_text)
            if hits:
                return _clean_phone(hits[0])
        except Exception:
            pass

        try:
            src = driver.page_source
            hits = _MOBILE_RE.findall(src)
            if hits:
                return _clean_phone(hits[0])
        except Exception:
            pass

        if _pass < 2:
            time.sleep(0.6)

    return "N/A"


# =============================================================================
#  MULTI-SELECTOR FIELD HELPERS
# =============================================================================

def _text_first(driver, selectors: list, default: str = "N/A") -> str:
    for sel in selectors:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            t = el.text.strip()
            if t:
                return t
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return default


def _attr_first(driver, selectors: list, attr: str, default: str = "N/A") -> str:
    for sel in selectors:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            v = (el.get_attribute(attr) or "").strip()
            if v:
                return v
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return default


def extract_website(driver) -> str:
    try:
        for btn in driver.find_elements(
            By.XPATH,
            "//a[contains(@data-item-id,'authority')] | "
            "//a[@aria-label and contains(@aria-label,'website')] | "
            "//a[@aria-label and contains(@aria-label,'Website')]"
        ):
            href = btn.get_attribute("href") or ""
            if href.startswith("http") and "google.com" not in href:
                return href
    except Exception:
        pass

    try:
        for a in driver.find_elements(By.XPATH, "//a[contains(@href,'http')]"):
            href = a.get_attribute("href") or ""
            if (href.startswith("http")
                    and "google.com" not in href
                    and "maps.app" not in href
                    and "goo.gl" not in href):
                return href
    except Exception:
        pass

    try:
        for el in driver.find_elements(By.XPATH, "//span[contains(text(),'www.')]"):
            t = el.text.strip()
            if t:
                return "http://" + t if not t.startswith("http") else t
    except Exception:
        pass

    return "N/A"


def extract_email_from_website(driver, website: str) -> str:
    if website == "N/A":
        return "N/A"
    try:
        driver.execute_script(f"window.open('{website}','_blank')")
        time.sleep(2.5)
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        body = driver.find_element(By.TAG_NAME, "body").text
        emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", body)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        return emails[0] if emails else "N/A"
    except Exception:
        try:
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
        return "N/A"


# =============================================================================
#  EXTRACT DATA — No lat/lng
# =============================================================================

def extract_data(driver, config: dict, log_q=None) -> dict:
    # NOTE: No Latitude / Longitude fields — removed by design
    rec = {k: "N/A" for k in COLUMNS}

    HEADING_SELECTORS = [
        "h1.DUwDvf", "h1.fontHeadlineLarge",
        "h1[class*='DUwDvf']", "div.tAiQdd h1", "div[role='main'] h1",
    ]
    found_heading = False
    for sel in HEADING_SELECTORS:
        if smart_wait(driver, sel, timeout=DETAIL_TIMEOUT):
            found_heading = True
            break

    if not found_heading:
        log("  Detail panel heading not found; skipping card.", log_q)
        return rec

    human_sleep(0.5, 1.0)

    # Business Name
    rec["Business Name"] = _text_first(driver, [
        "h1.DUwDvf", "h1.fontHeadlineLarge",
        "h1[class*='DUwDvf']", "div.tAiQdd h1",
    ])

    # Category
    rec["Category"] = _text_first(driver, [
        "button.DkEaL", "span.mgr77e", "div.skqShb",
        "button[jsaction*='category']", "span[aria-label*='type']",
    ])

    # Rating
    rec["Rating"] = _text_first(driver, [
        "div.F7nice span[aria-hidden='true']",
        "span.MW4etd",
        "div.F7nice > span > span[aria-hidden]",
        "span[class*='ceNzKf']",
    ])

    # Reviews Count
    rec["Reviews Count"] = _text_first(driver, [
        "div.F7nice span[aria-label*='review']",
        "span.UY7F9",
        "div.F7nice span[aria-label]",
        "button[jsaction*='review'] span",
    ])
    if rec["Reviews Count"] != "N/A":
        m = re.search(r"[\d,]+", rec["Reviews Count"])
        if m:
            rec["Reviews Count"] = m.group(0).replace(",", "")

    # Phone Number
    rec["Phone Number"] = extract_phone_enhanced(driver, log_q)

    # Website
    rec["Website"] = extract_website(driver)

    # Google Maps URL
    rec["Google Maps URL"] = driver.current_url

    # Full Address
    try:
        addr_found = False
        for btn in driver.find_elements(
            By.XPATH,
            "//button[@data-item-id='address'] | "
            "//button[contains(@aria-label,'Address')] | "
            "//button[contains(@data-tooltip,'Address')]"
        ):
            t = btn.text.strip()
            if t and len(t) > 5:
                rec["Full Address"] = t
                addr_found = True
                break

        if not addr_found:
            t = _text_first(driver, [
                "div[data-item-id='address']",
                "span.LrzXr[data-dtype='d3adr']",
                "div.Io6YTe.fontBodyMedium",
            ])
            if t != "N/A" and len(t) > 5:
                rec["Full Address"] = t

        if rec["Full Address"] != "N/A":
            pin = re.search(r"\b[1-9][0-9]{5}\b", rec["Full Address"])
            if pin:
                rec["Pincode"] = pin.group(0)

    except Exception:
        pass

    # WhatsApp
    phone = rec["Phone Number"]
    clean = phone.replace(" ", "").replace("+91", "").replace("-", "")
    if _MOBILE_RE.match(clean):
        rec["WhatsApp"] = f"https://wa.me/91{clean}"

    # Meta from config
    rec["City"]  = config.get("city", "N/A")
    rec["State"] = config.get("state", "N/A")
    rec["Area"]  = config.get("area", "N/A")

    # Optional email
    if config.get("fetch_email") and rec["Website"] != "N/A":
        rec["Email"] = extract_email_from_website(driver, rec["Website"])

    return rec


# =============================================================================
#  STRICT LOCATION VALIDATOR
#  Ensures scraped record address belongs to target area and not neighboring areas.
#  e.g. If user selects Virar, records mentioning Vasai/Nalasopara are rejected.
# =============================================================================

def validate_location(record: dict, config: dict) -> tuple[bool, str]:
    target_area = config.get("area", "").strip().lower()
    target_city = config.get("city", "").strip()
    state = config.get("state", "").strip()

    if not target_area or target_area == "(all areas)":
        return True, ""

    full_addr = str(record.get("Full Address", "")).lower()
    biz_name = str(record.get("Business Name", "")).lower()
    check_text = f"{biz_name} {full_addr}"

    # Get sibling neighbor areas from the same city/district
    all_areas = [a.lower() for a in INDIA_LOCATIONS.get(state, {}).get(target_city, [])]
    forbidden_neighbors = [a for a in all_areas if a != target_area and len(a) > 2]

    # If the listing mentions a different neighbor town and NOT the target area
    for neighbor in forbidden_neighbors:
        if neighbor in check_text and target_area not in check_text:
            return False, f"Located in {neighbor.title()}, outside {target_area.title()}"

    return True, ""


# =============================================================================
#  FILTER ENGINE (rating range, phone/website required, location strict)
# =============================================================================

def apply_filters(record: dict, filters: dict, config: dict) -> tuple[bool, str]:
    """
    Returns (is_valid: bool, reject_reason: str).
    """
    # Phone required
    phone = str(record.get("Phone Number", "")).strip()
    if filters.get("phone_only") and (not phone or phone == "N/A"):
        return False, "No phone number found (Phone Only is active)"

    # Website required
    website = str(record.get("Website", "")).strip()
    if filters.get("website_only") and (not website or website == "N/A"):
        return False, "No website found (Website Only is active)"

    # Rating filters
    rating_str = record.get("Rating", "N/A")
    try:
        rating = float(rating_str)
    except (ValueError, TypeError):
        rating = None

    min_r = float(filters.get("min_rating", 0) or 0)
    max_r = float(filters.get("max_rating", 0) or 0)

    if min_r > 0:
        if rating is None:
            if not filters.get("include_unrated", True):
                return False, "Unrated business (Min Rating active)"
        elif rating < min_r:
            return False, f"Rating {rating}⭐ < Min {min_r}⭐"

    if max_r > 0:
        if rating is None:
            if not filters.get("include_unrated", True):
                return False, "Unrated business (Max Rating active)"
        elif rating > max_r:
            return False, f"Rating {rating}⭐ > Max {max_r}⭐"

    # Strict location validation
    is_loc_ok, loc_reason = validate_location(record, config)
    if not is_loc_ok:
        return False, loc_reason

    return True, ""


# =============================================================================
#  ATOMIC EXCEL / CSV APPEND (no Latitude / Longitude columns)
# =============================================================================

def save_data(records: list, filepath: str, existing_keys: set,
              fmt: str = "xlsx") -> int:
    if not records:
        return 0

    # Remove duplicates
    new_rows = []
    for rec in records:
        key = (
            str(rec.get("Business Name", "")).strip().lower(),
            str(rec.get("Phone Number", "")).strip(),
        )
        if key[0] in ("", "n/a", "none"):
            continue
        if key not in existing_keys:
            existing_keys.add(key)
            new_rows.append(rec)

    if not new_rows:
        return 0

    df = pd.DataFrame(new_rows)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = "N/A"
    df = df[COLUMNS]
    df.drop_duplicates(subset=["Business Name", "Phone Number"], inplace=True)

    if df.empty:
        return 0

    if fmt == "csv":
        csv_path = filepath.replace(".xlsx", ".csv")
        df.to_csv(
            csv_path, mode="a",
            header=not os.path.exists(csv_path),
            index=False, encoding="utf-8-sig"
        )
    else:
        from openpyxl import load_workbook
        base, ext = os.path.splitext(filepath)
        tmp_path = f"{base}_temp{ext}"
        try:
            if os.path.exists(filepath):
                shutil.copy2(filepath, tmp_path)
                wb = load_workbook(tmp_path)
                ws = wb.active
                start_row = ws.max_row
                with pd.ExcelWriter(
                    tmp_path, engine="openpyxl", mode="a", if_sheet_exists="overlay"
                ) as writer:
                    df.to_excel(writer, index=False, header=False, startrow=start_row)
            else:
                with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False)

            if os.path.exists(filepath):
                os.replace(tmp_path, filepath)
            else:
                os.rename(tmp_path, filepath)

        except Exception as e:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            raise e

    return len(df)


# =============================================================================
#  SAFE CARD CLICK
# =============================================================================

def safe_click_card(driver, card, log_q=None):
    for attempt in range(3):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", card)
            time.sleep(random.uniform(0.5, 1.2))
            try:
                card.click()
            except Exception:
                driver.execute_script("arguments[0].click();", card)
            return True
        except StaleElementReferenceException:
            log(f"Stale element retry {attempt+1}", log_q)
            time.sleep(1)
        except ElementClickInterceptedException:
            try:
                driver.execute_script("arguments[0].click();", card)
                return True
            except Exception:
                pass
        except Exception as e:
            log(f"Click retry {attempt+1}: {e}", log_q)
            time.sleep(1)
    return False


# =============================================================================
#  SCRAPE WORKER (background thread)
# =============================================================================

def scrape_worker(config: dict, pause_event: threading.Event,
                  stop_event: threading.Event, log_q: queue.Queue,
                  count_callback, done_callback):
    state     = config["state"]
    city      = config["city"]
    areas     = config.get("areas") or [""]
    biz_type  = config["biz_type"]
    max_leads = int(config.get("max_leads", 100))
    filepath  = config["filepath"]
    fmt       = config.get("export_fmt", "xlsx")
    filters   = config.get("filters", {})
    fast_mode = config.get("fast_mode", False)
    sheet_id  = config.get("sheet_id", "").strip()
    creds_path = config.get("creds_path", "credentials.json")
    ws_name   = config.get("ws_name", "LeadFlow Leads")

    queries      = generate_queries(state, city, areas, biz_type)
    progress     = load_progress()
    done_set     = set(progress.get("done", []))
    total_saved  = progress.get("total_saved", 0)
    existing     = load_existing_keys(filepath)
    driver_fails = 0
    driver       = None

    log(f"🚀 LeadFlow started | {biz_type} in {', '.join(areas)} {city}, {state}", log_q)
    log(f"Generated {len(queries)} search queries.", log_q)
    log(f"Already done: {len(done_set)} | Target: {max_leads}", log_q)
    if sheet_id:
        log(f"[Sheets] Will sync to sheet ID: {sheet_id[:20]}...", log_q)

    try:
        driver = start_driver(log_q)

        qi = 0
        while qi < len(queries):
            if stop_event.is_set():
                break
            if total_saved >= max_leads:
                log(f"✅ Target of {max_leads} leads reached.", log_q)
                break

            query = queries[qi]
            if query in done_set:
                log(f"[Skip] {query}", log_q)
                qi += 1
                continue

            while pause_event.is_set() and not stop_event.is_set():
                time.sleep(0.5)
            if stop_event.is_set():
                break

            log(f"\n[{qi+1}/{len(queries)}] {query}", log_q)

            try:
                url = "https://www.google.com/maps/search/" + query.replace(" ", "+")
                driver.get(url)
                human_sleep(3.0, 5.5)

                if detect_captcha(driver):
                    if not handle_captcha(driver, log_q):
                        done_set.add(query)
                        save_progress({"done": list(done_set), "total_saved": total_saved})
                        qi += 1
                        continue

                feed_found = False
                FEED_WAIT_SELECTORS = [
                    "div[role='feed']", "div.m6QErb.DxyBCb",
                    "div.m6QErb", "a.hfpxzc",
                ]
                for fsel in FEED_WAIT_SELECTORS:
                    try:
                        WebDriverWait(driver, CARD_TIMEOUT).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, fsel))
                        )
                        feed_found = True
                        log(f"  Feed detected via '{fsel}'.", log_q)
                        break
                    except TimeoutException:
                        continue

                if not feed_found:
                    log(f"  No results feed for: {query}", log_q)
                    done_set.add(query)
                    save_progress({"done": list(done_set), "total_saved": total_saved})
                    qi += 1
                    continue

                scroll_results(driver, max_leads - total_saved, log_q)

                cards = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")
                if not cards:
                    cards = driver.find_elements(By.CSS_SELECTOR, "div.Nv2PK a")
                log(f"  Found {len(cards)} cards.", log_q)

                card_idx = 0

                while card_idx < len(cards):
                    if stop_event.is_set() or total_saved >= max_leads:
                        break
                    while pause_event.is_set() and not stop_event.is_set():
                        time.sleep(0.5)

                    try:
                        cards = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")
                        if not cards:
                            cards = driver.find_elements(By.CSS_SELECTOR, "div.Nv2PK a")

                        if card_idx >= len(cards):
                            # Try scrolling down more if target not reached yet
                            if total_saved < max_leads:
                                scroll_results(driver, max_leads - total_saved, log_q)
                                time.sleep(2.0)
                                cards = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")
                                if not cards:
                                    cards = driver.find_elements(By.CSS_SELECTOR, "div.Nv2PK a")
                            if card_idx >= len(cards):
                                break

                        card = cards[card_idx]
                    except Exception:
                        break

                    try:
                        name_lbl = card.get_attribute("aria-label") or f"#{card_idx+1}"
                        log(f"  [{card_idx+1}/{len(cards)}] {name_lbl[:60]}", log_q)

                        click_ok = safe_click_card(driver, card, log_q)
                        if not click_ok:
                            card_idx += 1
                            continue

                        human_sleep(1.5 if fast_mode else 2.2, 2.5 if fast_mode else 4.0)

                        rec = extract_data(driver, config, log_q)

                        if rec["Business Name"] == "N/A":
                            log("  Extraction returned N/A name; skipping.", log_q)
                            card_idx += 1
                            continue

                        is_valid, reject_reason = apply_filters(rec, filters, config)
                        if is_valid:
                            # Instant save to Excel immediately — no leads lost on stop!
                            saved = save_data([rec], filepath, existing, fmt)
                            if saved > 0:
                                total_saved += saved
                                count_callback(total_saved)
                                log(
                                    f"  💾 Saved [{total_saved}/{max_leads}] {rec['Business Name']} | "
                                    f"📞 {rec['Phone Number']} | "
                                    f"⭐ {rec['Rating']}",
                                    log_q
                                )
                                if sheet_id:
                                    push_to_sheets([rec], sheet_id, creds_path, ws_name, log_q)

                                if total_saved >= max_leads:
                                    log(f"🎯 Target of {max_leads} leads reached! Stopping.", log_q)
                                    break
                        else:
                            log(
                                f"  ✗ FILTERED OUT: {rec['Business Name']} | "
                                f"Reason: {reject_reason}",
                                log_q
                            )

                    except StaleElementReferenceException:
                        log(f"  Stale element at card {card_idx+1}; re-fetching.", log_q)
                        time.sleep(0.5)
                        continue
                    except Exception as e:
                        log(f"  [Err] Card {card_idx+1}: {e}", log_q)

                    card_idx += 1

                if total_saved >= max_leads:
                    break

                if total_saved == 0:
                    log(f"  ⚠ 0 matching leads in this query. (Tip: check your Rating/Phone filters)", log_q)

                cool = random.uniform(2.0, 4.0) if fast_mode else random.uniform(5.0, 9.0)
                log(f"  Cooling down {cool:.1f}s...", log_q)
                time.sleep(cool)

            except WebDriverException as e:
                driver_fails += 1
                log(f"  [WebDriver crash #{driver_fails}] {e}", log_q)
                try:
                    driver.quit()
                except Exception:
                    pass
                if driver_fails > MAX_DRIVER_RESTARTS:
                    log("  Max driver restarts exceeded. Stopping.", log_q)
                    break
                log(f"  Restarting driver (attempt {driver_fails}/{MAX_DRIVER_RESTARTS})...", log_q)
                time.sleep(5 + driver_fails * 2)
                try:
                    driver = start_driver(log_q)
                except Exception as re_e:
                    log(f"  Driver restart failed: {re_e}. Stopping.", log_q)
                    break
                continue

            except Exception as e:
                log(f"  [Error] Query skipped due to: {e}", log_q)
                if saved > 0:
                    done_set.add(query)
                    save_progress({"done": list(done_set), "total_saved": total_saved})

            qi += 1

    except Exception as e:
        log(f"[FATAL] {e}", log_q)
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass
        log(f"\n🏁 Done. Total leads saved: {total_saved}", log_q)
        count_callback(total_saved)
        done_callback()


# =============================================================================
#  GUI APPLICATION — LeadFlow
# =============================================================================

class LeadFlowApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("⚡ LeadFlow — India Lead Generation Tool")
        self.geometry("1160x760")
        self.minsize(1020, 680)
        self.configure(bg=DARK)

        self._pause_event  = threading.Event()
        self._stop_event   = threading.Event()
        self._thread       = None
        self._log_q        = queue.Queue()

        # Location vars
        self._state_var  = tk.StringVar()
        self._city_var   = tk.StringVar()
        self._area_var   = tk.StringVar()

        # Input vars
        self._biz_var    = tk.StringVar()
        self._leads_var  = tk.StringVar(value="100")
        self._fmt_var    = tk.StringVar(value="xlsx")
        self._fast_var   = tk.BooleanVar(value=False)
        self._email_var  = tk.BooleanVar(value=False)

        # Filter vars
        self._phone_var  = tk.BooleanVar(value=False)
        self._web_var    = tk.BooleanVar(value=False)
        self._min_rating_var = tk.DoubleVar(value=0)
        self._max_rating_var = tk.DoubleVar(value=0)
        self._unrated_var    = tk.BooleanVar(value=True)

        # Google Sheets vars
        self._sheet_id_var  = tk.StringVar(value="")
        self._creds_var     = tk.StringVar(value=os.path.join(os.getcwd(), "credentials.json"))
        self._ws_name_var   = tk.StringVar(value="LeadFlow Leads")
        self._sheets_on_var = tk.BooleanVar(value=False)

        # Status
        self._count_var  = tk.StringVar(value="0")

        self._build_ui()
        self._populate_states()
        self._poll_logs()

    # -------------------------------------------------------------------------
    def _build_ui(self):
        # Clean Top Header
        hdr = tk.Frame(self, bg="#0b0c13", height=48, padx=16, pady=8)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="⚡ LeadFlow",
            bg="#0b0c13", fg=GREEN,
            font=("Segoe UI", 16, "bold")
        ).pack(side="left")

        main = tk.Frame(self, bg=DARK, padx=8, pady=8)
        main.pack(fill="both", expand=True)

        # ---- Left Control Panel ----
        left = tk.Frame(main, bg=PANEL, width=420, padx=12, pady=10)
        left.pack(side="left", fill="y", padx=(0, 8))
        left.pack_propagate(False)

        def sec_title(parent, text):
            f = tk.Frame(parent, bg=PANEL)
            f.pack(fill="x", pady=(8, 4))
            tk.Label(f, text=text, bg=PANEL, fg=GREEN, font=("Segoe UI", 9, "bold")).pack(side="left")

        def field_lbl(parent, text):
            return tk.Label(parent, text=text, bg=PANEL, fg=MUTED, font=("Segoe UI", 8, "bold"))

        # --- Section 1: Target ---
        sec_title(left, "Target")

        loc_grid = tk.Frame(left, bg=PANEL)
        loc_grid.pack(fill="x", pady=(0, 4))
        loc_grid.columnconfigure(0, weight=1)
        loc_grid.columnconfigure(1, weight=1)

        field_lbl(loc_grid, "State").grid(row=0, column=0, sticky="w", padx=(0, 4))
        field_lbl(loc_grid, "City").grid(row=0, column=1, sticky="w", padx=(4, 0))

        self._state_cb = ttk.Combobox(loc_grid, textvariable=self._state_var, state="readonly")
        self._state_cb.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=(2, 4))
        self._state_cb.bind("<<ComboboxSelected>>", self._on_state)

        self._city_cb = ttk.Combobox(loc_grid, textvariable=self._city_var, state="readonly")
        self._city_cb.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=(2, 4))
        self._city_cb.bind("<<ComboboxSelected>>", self._on_city)

        field_lbl(left, "Area").pack(anchor="w")
        self._area_cb = ttk.Combobox(left, textvariable=self._area_var, state="readonly")
        self._area_cb.pack(fill="x", pady=(2, 6))

        field_lbl(left, "Business Type").pack(anchor="w")
        biz_cb = ttk.Combobox(left, textvariable=self._biz_var)
        biz_cb.config(values=BUSINESS_TYPES)
        biz_cb.pack(fill="x", pady=(2, 6))

        sub_grid = tk.Frame(left, bg=PANEL)
        sub_grid.pack(fill="x", pady=(0, 6))
        sub_grid.columnconfigure(0, weight=1)
        sub_grid.columnconfigure(1, weight=1)

        field_lbl(sub_grid, "Target Leads").grid(row=0, column=0, sticky="w", padx=(0, 4))
        field_lbl(sub_grid, "Format").grid(row=0, column=1, sticky="w", padx=(4, 0))

        tk.Entry(sub_grid, textvariable=self._leads_var, bg=CARD, fg=WHITE,
                 font=("Segoe UI", 9), relief="flat", insertbackground=WHITE
                 ).grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=(2, 0), ipady=3)

        fmt_f = tk.Frame(sub_grid, bg=PANEL)
        fmt_f.grid(row=1, column=1, sticky="w", padx=(4, 0), pady=(2, 0))
        tk.Radiobutton(fmt_f, text="Excel", variable=self._fmt_var, value="xlsx",
                       bg=PANEL, fg=WHITE, selectcolor=CARD, font=("Segoe UI", 8)).pack(side="left")
        tk.Radiobutton(fmt_f, text="CSV", variable=self._fmt_var, value="csv",
                       bg=PANEL, fg=WHITE, selectcolor=CARD, font=("Segoe UI", 8)).pack(side="left", padx=(10, 0))

        # --- Section 2: Filters ---
        sec_title(left, "Filters")

        rate_grid = tk.Frame(left, bg=PANEL)
        rate_grid.pack(fill="x", pady=(0, 4))
        rate_grid.columnconfigure(0, weight=1)
        rate_grid.columnconfigure(1, weight=1)

        self._min_lbl = tk.Label(rate_grid, text="Min: Off", bg=PANEL, fg=MUTED, font=("Segoe UI", 8, "bold"))
        self._min_lbl.grid(row=0, column=0, sticky="w", padx=(0, 4))
        self._max_lbl = tk.Label(rate_grid, text="Max: Off", bg=PANEL, fg=MUTED, font=("Segoe UI", 8, "bold"))
        self._max_lbl.grid(row=0, column=1, sticky="w", padx=(4, 0))

        def on_min_slide(val):
            v = float(val)
            self._min_lbl.config(text=f"Min: {v}⭐" if v > 0 else "Min: Off")
            self._update_filter_label()

        def on_max_slide(val):
            v = float(val)
            self._max_lbl.config(text=f"Max: {v}⭐" if v > 0 else "Max: Off")
            self._update_filter_label()

        self._min_r_scale = tk.Scale(
            rate_grid, from_=0, to=5, resolution=0.5, orient="h",
            variable=self._min_rating_var, bg=CARD, fg=WHITE,
            highlightthickness=0, relief="flat", length=140,
            command=on_min_slide
        )
        self._min_r_scale.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=2)

        self._max_r_scale = tk.Scale(
            rate_grid, from_=0, to=5, resolution=0.5, orient="h",
            variable=self._max_rating_var, bg=CARD, fg=WHITE,
            highlightthickness=0, relief="flat", length=140,
            command=on_max_slide
        )
        self._max_r_scale.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=2)

        chk_grid = tk.Frame(left, bg=PANEL)
        chk_grid.pack(fill="x", pady=(2, 6))
        chk_grid.columnconfigure(0, weight=1)
        chk_grid.columnconfigure(1, weight=1)

        def chk(parent, text, var):
            return tk.Checkbutton(parent, text=text, variable=var, bg=PANEL, fg=WHITE,
                                  selectcolor=CARD, activebackground=PANEL, activeforeground=WHITE,
                                  font=("Segoe UI", 8), command=self._update_filter_label)

        chk(chk_grid, "Phone only", self._phone_var).grid(row=0, column=0, sticky="w")
        chk(chk_grid, "Website only", self._web_var).grid(row=0, column=1, sticky="w")
        chk(chk_grid, "Include unrated", self._unrated_var).grid(row=1, column=0, sticky="w")
        chk(chk_grid, "Fast mode", self._fast_var).grid(row=1, column=1, sticky="w")
        chk(chk_grid, "Fetch email", self._email_var).grid(row=2, column=0, sticky="w")

        # --- Section 3: Google Sheets ---
        sec_title(left, "Google Sheets")

        sheets_top = tk.Frame(left, bg=PANEL)
        sheets_top.pack(fill="x", pady=(0, 2))
        chk(sheets_top, "Enable Sheets Export", self._sheets_on_var).pack(side="left")
        tk.Button(sheets_top, text="Test", command=self._test_sheets,
                  bg=CARD, fg=WHITE, font=("Segoe UI", 8), relief="flat", cursor="hand2", padx=6
                  ).pack(side="right")

        field_lbl(left, "Sheet ID").pack(anchor="w")
        tk.Entry(left, textvariable=self._sheet_id_var, bg=CARD, fg=WHITE,
                 font=("Segoe UI", 8), relief="flat", insertbackground=WHITE
                 ).pack(fill="x", pady=(2, 4), ipady=2)

        # --- Section 4: Controls ---
        sec_title(left, "Actions")

        btn_bar = tk.Frame(left, bg=PANEL)
        btn_bar.pack(fill="x", pady=(4, 4))
        btn_bar.columnconfigure((0, 1, 2, 3), weight=1)

        def act_btn(parent, text, cmd, col, bg_col):
            b = tk.Button(parent, text=text, command=cmd, bg=bg_col, fg=WHITE,
                          font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
                          activebackground=bg_col, activeforeground=WHITE, pady=4)
            b.grid(row=0, column=col, sticky="ew", padx=2)
            return b

        self._start_btn  = act_btn(btn_bar, "▶ Start", self._start, 0, GREEN)
        self._pause_btn  = act_btn(btn_bar, "⏸ Pause", self._pause, 1, YELLOW)
        self._pause_btn.config(state="disabled")
        self._resume_btn = act_btn(btn_bar, "▶ Resume", self._resume, 2, BLUE)
        self._resume_btn.config(state="disabled")
        self._stop_btn   = act_btn(btn_bar, "⏹ Stop", self._stop, 3, RED)
        self._stop_btn.config(state="disabled")

        tk.Button(left, text="📂 Open Export File", command=self._open_file,
                  bg=CARD, fg=WHITE, font=("Segoe UI", 9), relief="flat", cursor="hand2", pady=4
                  ).pack(fill="x", pady=(4, 0))

        # ---- Right Panel (Live Monitor) ----
        right = tk.Frame(main, bg=PANEL, padx=12, pady=10)
        right.pack(side="left", fill="both", expand=True)

        top_stat = tk.Frame(right, bg=PANEL)
        top_stat.pack(fill="x", pady=(0, 6))

        tk.Label(top_stat, text="Leads Saved:", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 11)).pack(side="left")
        tk.Label(top_stat, textvariable=self._count_var, bg=PANEL, fg=GREEN,
                 font=("Segoe UI", 20, "bold")).pack(side="left", padx=(6, 16))

        self._filter_lbl = tk.Label(top_stat, text="Filters: None", bg=PANEL, fg=YELLOW,
                                    font=("Segoe UI", 9))
        self._filter_lbl.pack(side="left")

        tk.Button(top_stat, text="Clear Log", command=self._clear_log,
                  bg=CARD, fg=WHITE, font=("Segoe UI", 8), relief="flat", cursor="hand2", padx=8
                  ).pack(side="right")

        self._log_box = scrolledtext.ScrolledText(
            right, bg="#0c0d14", fg="#10b981",
            font=("Consolas", 9), relief="flat", wrap="word",
            insertbackground="#10b981", state="disabled"
        )
        self._log_box.pack(fill="both", expand=True, pady=(0, 4))

        self._status = tk.Label(
            self, text="Ready", bg=DARK, fg=MUTED, font=("Segoe UI", 8), anchor="w", padx=12, pady=3
        )
        self._status.pack(fill="x")

    # -------------------------------------------------------------------------
    def _populate_states(self):
        states = sorted(INDIA_LOCATIONS.keys())
        self._state_cb.config(values=states)
        self._state_var.set(states[0])
        self._on_state(None)

    def _on_state(self, _):
        state = self._state_var.get()
        cities = sorted(INDIA_LOCATIONS.get(state, {}).keys())
        self._city_cb.config(values=cities)
        if cities:
            self._city_var.set(cities[0])
        self._on_city(None)

    def _on_city(self, _):
        state = self._state_var.get()
        city  = self._city_var.get()
        areas = INDIA_LOCATIONS.get(state, {}).get(city, [])
        self._area_cb.config(values=["(All areas)"] + sorted(areas))
        self._area_var.set("(All areas)")

    def _browse_creds(self):
        f = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
            title="Select credentials.json"
        )
        if f:
            self._creds_var.set(f)

    def _test_sheets(self):
        sheet_id  = self._sheet_id_var.get().strip()
        creds_path = self._creds_var.get().strip()
        if not sheet_id:
            messagebox.showwarning("Missing", "Please enter a Google Sheet ID first.")
            return
        if not GSPREAD_AVAILABLE:
            messagebox.showerror(
                "Package Missing",
                "gspread is not installed.\nRun: pip install gspread google-auth"
            )
            return
        client, err = get_gspread_client(creds_path)
        if err:
            messagebox.showerror("Connection Failed", err)
            return
        try:
            sh = client.open_by_key(sheet_id)
            messagebox.showinfo("Connected ✅",
                                f"Successfully connected to:\n{sh.title}\n\n"
                                f"Worksheets: {[w.title for w in sh.worksheets()]}")
        except Exception as e:
            messagebox.showerror("Connection Failed", str(e))

    def _update_filter_label(self):
        parts = []
        if self._phone_var.get():
            parts.append("Phone required")
        if self._web_var.get():
            parts.append("Website required")
        min_r = self._min_rating_var.get()
        max_r = self._max_rating_var.get()
        if min_r > 0:
            parts.append(f"Rating >= {min_r} stars")
        if max_r > 0:
            parts.append(f"Rating <= {max_r} stars")
        area = self._area_var.get()
        if area and area != "(All areas)":
            parts.append(f"Area = {area} ONLY")
        self._filter_lbl.config(
            text="Active Filters: " + ("  |  ".join(parts) if parts else "none")
        )

    # -------------------------------------------------------------------------
    def _start(self):
        self._update_filter_label()

        state    = self._state_var.get().strip()
        city     = self._city_var.get().strip()
        area_sel = self._area_var.get().strip()
        biz_type = self._biz_var.get().strip()
        max_lead = self._leads_var.get().strip()
        fmt      = self._fmt_var.get()

        if not all([state, city, biz_type, max_lead]):
            messagebox.showerror("Missing Input", "Please fill all required fields.")
            return
        try:
            int(max_lead)
        except ValueError:
            messagebox.showerror("Invalid", "Target leads must be a number.")
            return

        # Strict area selection
        if area_sel and area_sel != "(All areas)":
            areas = [area_sel]
            area_for_config = area_sel
        else:
            areas = INDIA_LOCATIONS.get(state, {}).get(city, [""])
            area_for_config = ""

        biz_clean  = re.sub(r"[^a-zA-Z0-9]", "_", biz_type)
        city_clean = re.sub(r"[^a-zA-Z0-9]", "_", city)
        area_clean = re.sub(r"[^a-zA-Z0-9]", "_", area_for_config) if area_for_config else ""
        name_parts = [state, city_clean]
        if area_clean:
            name_parts.append(area_clean)
        name_parts.append(biz_clean)
        dynamic_name = "_".join(name_parts).replace(" ", "_")
        script_dir = os.path.dirname(os.path.abspath(__file__))
        filepath = os.path.join(script_dir, f"{dynamic_name}.{fmt}")
        self._last_filepath = filepath

        # Google Sheets config
        sheets_on = self._sheets_on_var.get()
        sheet_id  = self._sheet_id_var.get().strip() if sheets_on else ""
        creds_path = self._creds_var.get().strip()
        ws_name   = self._ws_name_var.get().strip() or "LeadFlow Leads"

        if sheets_on and not sheet_id:
            messagebox.showwarning(
                "Sheet ID Missing",
                "Google Sheets export is enabled but no Sheet ID was provided.\n"
                "Please enter the Sheet ID or disable Sheets export."
            )
            return

        config = {
            "state": state, "city": city, "areas": areas,
            "area": area_for_config,
            "biz_type": biz_type, "max_leads": int(max_lead),
            "filepath": filepath, "export_fmt": fmt,
            "fast_mode": self._fast_var.get(),
            "fetch_email": self._email_var.get(),
            "sheet_id": sheet_id,
            "creds_path": creds_path,
            "ws_name": ws_name,
            "filters": {
                "phone_only": self._phone_var.get(),
                "website_only": self._web_var.get(),
                "min_rating": self._min_rating_var.get(),
                "max_rating": self._max_rating_var.get(),
                "include_unrated": self._unrated_var.get(),
            },
        }

        self._pause_event.clear()
        self._stop_event.clear()
        self._count_var.set("0")
        loc_str = f"{area_sel} → " if area_sel != "(All areas)" else ""
        self._set_status(f"🚀 Scraping: {biz_type} in {loc_str}{city}, {state}")

        self._start_btn.config(state="disabled")
        self._pause_btn.config(state="normal")
        self._stop_btn.config(state="normal")

        self._thread = threading.Thread(
            target=scrape_worker,
            args=(
                config,
                self._pause_event,
                self._stop_event,
                self._log_q,
                self._update_count,
                self._on_done,
            ),
            daemon=True,
        )
        self._thread.start()

    def _pause(self):
        self._pause_event.set()
        self._pause_btn.config(state="disabled")
        self._resume_btn.config(state="normal")
        self._set_status("⏸ Paused")

    def _resume(self):
        self._pause_event.clear()
        self._resume_btn.config(state="disabled")
        self._pause_btn.config(state="normal")
        self._set_status("▶ Resumed")

    def _stop(self):
        self._stop_event.set()
        self._pause_event.clear()
        self._set_status("⏹ Stopping — leads saved to Excel...")

    def _on_done(self):
        self.after(0, self._reset_buttons)

    def _reset_buttons(self):
        self._start_btn.config(state="normal")
        self._pause_btn.config(state="disabled")
        self._resume_btn.config(state="disabled")
        self._stop_btn.config(state="disabled")
        self._set_status("✅ Done. Leads saved in Excel.")

    def _update_count(self, n: int):
        self.after(0, lambda: self._count_var.set(str(n)))

    def _set_status(self, msg: str):
        self.after(0, lambda: self._status.config(text=msg))

    def _open_file(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        target = getattr(self, "_last_filepath", None)
        if not target or not os.path.exists(target):
            target = filedialog.askopenfilename(
                initialdir=script_dir,
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All", "*.*")]
            )
        if target and os.path.exists(target):
            if sys.platform == "win32":
                os.startfile(target)
            elif sys.platform == "darwin":
                os.system(f"open '{target}'")
            else:
                os.system(f"xdg-open '{target}'")

    def _clear_log(self):
        self._log_box.config(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.config(state="disabled")

    def _poll_logs(self):
        while not self._log_q.empty():
            msg = self._log_q.get_nowait()
            self._log_box.config(state="normal")
            self._log_box.insert("end", msg + "\n")
            self._log_box.see("end")
            self._log_box.config(state="disabled")
        self.after(150, self._poll_logs)


# =============================================================================
#  ENTRY POINT
# =============================================================================

def main():
    app = LeadFlowApp()
    app.mainloop()


if __name__ == "__main__":
    main()

