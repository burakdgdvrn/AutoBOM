# SmartBOM Extractor

An AI-powered automation system for extracting and comparing Bill of Materials (BOM) from technical PDF drawings.

## 🚀 Features
- **AI OCR Extraction:** Automatically reads isometric technical drawings (PDFs) and extracts Fabrication Materials, Erection Materials, and Cut Pipe Lengths using EasyOCR and OpenCV.
- **Smart Heuristic Comparison:** Compares extracted AI data with ground-truth Master Excel files using a cell-based heuristic scoring algorithm to minimize false positives.
- **Interactive UI:** Provides a dark/light themed, responsive dashboard to review mismatches, manually correct OCR errors via auto-complete, and visualize cell-level discrepancies side-by-side.
- **Automated QA Reporting:** Exports a comprehensive QA Excel report containing perfectly matched rows, AI-missed rows, and cell discrepancies.

## 🏗️ Architecture
- **Backend:** Python (Flask, Pandas, EasyOCR, OpenCV)
- **Frontend:** Vanilla JavaScript, HTML, CSS, Handsontable, SweetAlert2
- **Data:** Stores raw PDFs and Master Excel files dynamically and securely.

## ⚙️ Installation
1. Clone the repository.
2. Setup a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Usage
1. Place your target PDFs in the `data/input_pdfs/` directory.
2. Place your master Excel file in the `data/master_excel/` directory.
3. Start the backend API server:
   ```bash
   python backend/app.py
   ```
4. Open `frontend/ERP Sistemi.html` in your web browser or via Live Server.
