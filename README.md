# ⚡ LeadFlow

**LeadFlow** is an automated Google Maps lead generation and extraction desktop application tailored for Indian businesses and digital marketing agencies. It features strict geographic and rating filters, real-time lead extraction, instant atomic saving, and automatic synchronization with both **Excel (.xlsx)** and **Google Sheets**.

---

## ✨ Features

- **📍 Strict Geographic Filtering**: Pinpoint searches to exact localities (e.g., *Virar*) while automatically filtering out results from neighboring towns (*Vasai, Nalasopara*).
- **⭐ Rating Range Filter**: Filter leads by Minimum and Maximum star ratings (e.g., find businesses with rating $\le$ 3.5★).
- **💾 Instant Lead Saving**: Leads are written to Excel and pushed to Google Sheets in real-time as each lead is discovered. No data is lost if stopped mid-run.
- **📊 Google Sheets Auto-Sync**: Automatically stream generated leads directly into a shared Google Sheet via the Google Sheets API.
- **🧹 Clean Data Output**: Extracts Business Name, Category, Phone Number, WhatsApp deep-link, Website, Email, Full Address, City, State, Area, Pincode, Rating, and Review Count. (No latitude/longitude clutter).
- **🛡️ Anti-Detection & Auto-Recovery**: Built-in human-like scrolling, randomized pauses, CDP-level navigator patching, and automatic WebDriver restart on stalls.
- **🖥️ Modern Desktop GUI**: Built using Tkinter with a clean dark theme, live activity terminal, and real-time lead counter.

---

## 🚀 Quick Start

### 1. Prerequisites
- **Python 3.10+**
- **Google Chrome** installed on your system

### 2. Installation
Clone the repository and install required dependencies:
```bash
git clone https://github.com/RishiVartak6/LeadFlow.git
cd LeadFlow
pip install -r requirements.txt
```

### 3. Launching the App
Run the helper batch script or launch directly with Python:
- **Windows**: Double-click `run.bat` or run:
  ```bash
  python LeadFlow.py
  ```

---

## 📊 Google Sheets Integration (Optional)

To automatically export leads to Google Sheets:

1. **Google Cloud Console**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/) and create a project.
   - Enable both **Google Sheets API** and **Google Drive API**.
2. **Create Service Account**:
   - Navigate to **APIs & Services > Credentials > Create Credentials > Service Account**.
   - Create a JSON key for the service account and save it as `credentials.json` in the project root directory.
3. **Share Your Google Sheet**:
   - Open your Google Sheet, click **Share**, and grant **Editor** access to the service account email (found in `credentials.json`).
4. **Configure in LeadFlow**:
   - Check **Enable Sheets Export**, paste your **Sheet ID** (from the sheet URL), and click **Test** to verify connection.

---

## 📁 Output Data Format

Exported Excel (`.xlsx`) and Google Sheets include the following columns:

| Column | Description |
|---|---|
| **Business Name** | Name of the business listing |
| **Category** | Primary business category |
| **Phone Number** | Formatted Indian contact number |
| **WhatsApp** | Direct WhatsApp chat link |
| **Website** | Official business website URL |
| **Email** | Scraped business email (if enabled) |
| **Full Address** | Complete address string |
| **City** | Selected city/district |
| **State** | Selected state |
| **Area** | Selected locality |
| **Pincode** | 6-digit postal code |
| **Rating** | Star rating (0.0 – 5.0) |
| **Reviews Count** | Total review count |
| **Google Maps URL** | Direct link to Google Maps listing |

---

## 📜 License
MIT License. Free for personal and commercial use.
