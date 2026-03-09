# RAPA Analyzer: Chrome Extension & Data Analysis Tool 🎪📉

Professional suite for analyzing competition data from the Russian Aerial Acrobatics Federation (RAPA).

## 🚀 Features

### 🧩 Chrome Extension (RAPA Analyzer)
- **Direct Data Extraction:** One-click participant loading from `referee.f-rapa.ru`.
- **Beautiful Dashboard:** Interactive results view in a separate tab with high-performance filtering.
- **Dynamic Cascading Filters:** Filter by City, School, Discipline, and Category with real-time updates.
- **Advanced Metrics:** Automated age calculation, unique participant counts, and school-based statistics.
- **Google Sheets Integration:** Seamless export/import from the cloud using OAuth2.
- **Reliable Excel Export:** Structured XLSX files with auto-filters and formatting.
- **Offline Mode:** Load and analyze previously exported Excel files without an internet connection.

### 🖥 Python Analysis Service
- **Streamlit Web App:** Lightweight web interface for local Excel file analysis (`app.py`).
- **CLI Tool:** Fast console-based statistics generator (`parse_excel.py`).
- **School Normalization:** Intelligent mapping of various school name variations.

## 📦 Installation

### Chrome Extension
1. Download or clone this repository.
2. Open Chrome and navigate to `chrome://extensions/`.
3. Enable **"Developer mode"** (top right).
4. Click **"Load unpacked"** and select the `Chrome_Extension` folder.
5. (Optional) Follow `ИНСТРУКЦИЯ_GOOGLE.md` in the extension folder to enable Google Sheets integration.

### Python Tools
1. Ensure you have Python 3.10+ installed.
2. Install dependencies:
   ```bash
   pip install streamlit pandas openpyxl
   ```
3. Run the web app:
   ```bash
   streamlit run app.py
   ```

## 🛠 Project Structure
- `Chrome_Extension/`: All files for the browser extension.
- `app.py`: Main Streamlit application.
- `parse_excel.py`: Core logic for Excel parsing and normalization.
- `run_app.sh`: Helper script for quick startup on macOS.

---
*Created for the Russian Aerial Acrobatics Federation (RAPA) analytics.*
