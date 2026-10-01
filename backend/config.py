# -*- coding: utf-8 -*-
"""
Yapılandırma sabitleri ve tablo tanımları.
Tüm bölge, sütun ve kategori konfigürasyonları bu modülde toplanmıştır.
"""

import os

# ============================================================================
# YAPILANDIRMA
# ============================================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF_FOLDER = os.path.join(BASE_DIR, "data", "input_pdfs", "PDFler")
OUTPUT_FILE = os.path.join(BASE_DIR, "data", "outputs", "sonuc.xlsx")
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
ZOOM = 3  # 3x zoom = ~450 DPI

# --- Aşama 1: Konfigürasyonlar ---
# Item Code başındaki 'I', '1' veya '10' gibi öneklerin ERP için tamamen atılıp atılmayacağını belirler.
# False ise sadece OCR düzeltmesi yapılır (örn: 1239 -> I239). True ise önek silinir.
ERP_STRIP_PREFIX = True

# Bölge tanımları (normalize oranlar, tam sayfa üzerinde)
REGIONS = {
    "FABRICATION MATERIALS": {"x1": 0.66, "y1": 0.01, "x2": 0.99, "y2": 0.30},
    "ERECTION MATERIALS":    {"x1": 0.66, "y1": 0.27, "x2": 0.99, "y2": 0.62},
    "CUT PIPE LENGTH":       {"x1": 0.63, "y1": 0.55, "x2": 0.99, "y2": 0.79},
}

# Sütun sınırları (kırpılmış bölge içinde, normalize x)
# OCR debug analizi ile kalibre edilmiş
FAB_EREC_COLS = [
    # (x_start, x_end, name)
    (0.00, 0.185, "P/L NO"),
    (0.185, 0.515, "COMPONENT DESCRIPTION"),
    (0.515, 0.570, "N.S. (INS)"),
    (0.570, 0.665, "ITEM CODE"),
    (0.665, 0.755, "CLIENT"),
    (0.755, 0.845, "QTY"),
    (0.845, 1.000, "WEIGHT"),
]

CUT_PIPE_COLS = [
    (0.00, 0.170, "PIECE NO"),
    (0.170, 0.230, "PT NO"),
    (0.230, 0.350, "LENGTH (MM)"),
    (0.350, 0.450, "N.S. (INS)"),
    (0.450, 0.555, "ITEM NO"),
    (0.555, 0.830, "SPOOL NO"),
    (0.830, 1.000, "REMARKS"),
]

TABLE_CONFIGS = {
    "FABRICATION MATERIALS": {
        "headers": [c[2] for c in FAB_EREC_COLS],
        "col_ranges": FAB_EREC_COLS,
        "title_keyword": "FABRICATION MATERIALS",
    },
    "ERECTION MATERIALS": {
        "headers": [c[2] for c in FAB_EREC_COLS],
        "col_ranges": FAB_EREC_COLS,
        "title_keyword": "ERECTION MATERIALS",
    },
    "CUT PIPE LENGTH": {
        "headers": [c[2] for c in CUT_PIPE_COLS],
        "col_ranges": CUT_PIPE_COLS,
        "title_keyword": "CUT PIPE LENGTH",
    },
}

# Kategori başlıkları (tam eşleşme veya çok kısa olmalı)
CATEGORY_KEYWORDS_EXACT = {
    "PIPE", "FITTINGS", "FLANGES", "BOLTS", "GASKETS",
    "SUPPORTS", "SPECIAL ITEMS", "INSTRUMENTS",
}
# Prefix eşleşme (sadece bu prefix ile başlayıp sonrası kısa olmalı)
CATEGORY_PREFIXES = [
    "VALVES", "IN-LINE ITEMS", "VALVES / IN-LINE ITEMS",
    "VALVES/IN-LINE ITEMS",
]
