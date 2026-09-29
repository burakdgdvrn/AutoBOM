# -*- coding: utf-8 -*-
"""
PDF'den Fabrication Materials, Erection Materials ve Cut Pipe Length tablolarını
OCR ile çıkarıp Excel'e aktaran ana script. v3

Strateji: Her tablo bölgesini kırp -> OCR ile tüm metni al ->
          Y koordinatına göre satırlara grupla ->
          X koordinatına göre sütunlara ata ->
          Kategori / veri ayrımı yap ->
          Devam satırlarını birleştir -> Excel'e yaz

Kullanım:
  1. PDFler/ klasörüne PDF dosyalarını koyun
  2. Bu scripti çalıştırın
  3. Çıktı: sonuc.xlsx
"""

import pymupdf
import easyocr
import os
import sys
import re
from io import BytesIO
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================================
# YAPILANDIRMA
# ============================================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF_FOLDER = os.path.join(BASE_DIR, "data", "input_pdfs", "PDFler")
OUTPUT_FILE = os.path.join(BASE_DIR, "data", "outputs", "sonuc.xlsx")
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
ZOOM = 3  # 3x zoom = ~450 DPI

# YOLO Model (gecikmeli yukleme icin None, ilk kullanımda yuklenir)
_yolo_model = None
def get_yolo_model():
    """YOLO modelini gecikmeli olarak yukle (lazy loading)."""
    global _yolo_model
    if _yolo_model is None:
        try:
            from ultralytics import YOLO
            model_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "models", "pdf_read_yolov26s.pt")
            if os.path.exists(model_path):
                _yolo_model = YOLO(model_path)
                print(f"  YOLO modeli yuklendi: {model_path}")
            else:
                print(f"  UYARI: YOLO modeli bulunamadi: {model_path}")
        except ImportError:
            print("  UYARI: ultralytics paketi bulunamadi. YOLO devre disi.")
    return _yolo_model

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
    "SUPPORTS", "SPECIAL ITEMS",
}
# Prefix eşleşme (sadece bu prefix ile başlayıp sonrası kısa olmalı)
CATEGORY_PREFIXES = [
    "VALVES", "IN-LINE ITEMS", "VALVES / IN-LINE ITEMS",
    "VALVES/IN-LINE ITEMS",
]

# ============================================================================
# OCR MOTORU
# ============================================================================

print("OCR motoru baslatiliyor...")
reader = easyocr.Reader(['en'], gpu=False, verbose=False)
print("OCR motoru hazir.\n")

# ============================================================================
# YARDIMCI FONKSİYONLAR
# ============================================================================

def canonicalize_spool(raw_spool, base_drawing, yolo_candidates=None):
    """
    Validates and canonicalizes a raw OCR spool string strictly.
    Returns: {raw_spool, canonical_spool, normalization_reason, candidate_spools, selected_candidate, confidence, status}
    """
    if yolo_candidates is None:
        yolo_candidates = []
        
    result = {
        "raw_spool": raw_spool,
        "canonical_spool": None,
        "normalization_reason": None,
        "candidate_spools": [c["text"] for c in yolo_candidates],
        "selected_candidate": None,
        "confidence": 0.0,
        "status": "UNRESOLVED"
    }
    
    if not raw_spool:
        result["normalization_reason"] = "Empty raw spool"
        return result
        
    # Isolate spool part if there's noise prepended (e.g. "1 15722 IG-502450...")
    spool_part = raw_spool
    if "502" in raw_spool:
        parts = raw_spool.split()
        for p in parts:
            if "502" in p:
                spool_part = p
                break
                
    raw_spool_upper = spool_part.upper().strip()
    
    # Strictly extract expected suffix (e.g., SP followed by digits or O->0)
    # E.g., SPO1, SP02, SP3
    match = re.search(r'SP[O0\d]+$', raw_spool_upper.replace(' ', '-'))
    
    if match:
        raw_suffix = match.group(0)
        canonical_suffix = raw_suffix.replace('O', '0')
        if not canonical_suffix.startswith('SP0') and len(canonical_suffix) == 3:
            # SP1 -> SP01
            canonical_suffix = canonical_suffix.replace('SP', 'SP0')
            
        canonical = f"{base_drawing}-{canonical_suffix}"
        
        # Check YOLO candidates that match this suffix
        matching_candidates = []
        for y_spool in yolo_candidates:
            y_text = y_spool['text'].upper()
            if y_text.endswith(canonical_suffix):
                matching_candidates.append(y_spool)
                
        if len(matching_candidates) == 1:
            y_spool = matching_candidates[0]
            result["canonical_spool"] = canonical
            result["selected_candidate"] = y_spool["text"]
            result["normalization_reason"] = f"Prefix from base_drawing, unique suffix match YOLO ({canonical_suffix})"
            result["confidence"] = y_spool.get('conf', 0.9)
            result["status"] = "VALID"
            return result
            
        elif len(matching_candidates) > 1:
            # Sort by confidence
            matching_candidates.sort(key=lambda x: x.get('conf', 0), reverse=True)
            conf_diff = matching_candidates[0].get('conf', 0) - matching_candidates[1].get('conf', 0)
            if conf_diff > 0.10:  # Arbitrary threshold for high confidence lead
                y_spool = matching_candidates[0]
                result["canonical_spool"] = canonical
                result["selected_candidate"] = y_spool["text"]
                result["normalization_reason"] = f"Prefix from base_drawing, suffix matched YOLO by highest conf diff ({canonical_suffix})"
                result["confidence"] = y_spool.get('conf', 0.9)
                result["status"] = "VALID"
            else:
                result["normalization_reason"] = f"Ambiguous candidates for suffix {canonical_suffix}"
                result["status"] = "AMBIGUOUS"
            return result
                
        # If no YOLO candidate matches, but we parsed a valid suffix
        result["canonical_spool"] = canonical
        result["normalization_reason"] = f"Prefix from base_drawing, parsed suffix {canonical_suffix} (no YOLO match)"
        result["confidence"] = 0.8
        result["status"] = "VALID"
        return result
        
    result["normalization_reason"] = "Could not confidently extract suffix"
    result["status"] = "UNRESOLVED"
    return result

def validate_piece_no(raw_piece):
    """Validates and normalizes Piece No."""
    result = {
        "raw": raw_piece,
        "normalized": None,
        "status": "UNRESOLVED"
    }
    if not raw_piece:
        result["status"] = "EMPTY"
        return result
        
    # Remove noise characters like <, >
    cleaned = re.sub(r'[<>]', '', raw_piece.strip())
    
    # If the remaining is just digits
    if cleaned.isdigit():
        if len(cleaned) > 1 and cleaned != raw_piece.strip():
             result["normalized"] = None
             result["status"] = "AMBIGUOUS"
             result["reason"] = "Multiple digits extracted with noise, requires bbox proof"
        else:
             result["normalized"] = cleaned
             result["status"] = "VALID"
    else:
        result["status"] = "AMBIGUOUS"
        
    return result

def pdf_to_image(pdf_path):
    """PDF sayfasını yüksek çözünürlüklü görüntüye çevir."""
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    mat = pymupdf.Matrix(ZOOM, ZOOM)
    pix = page.get_pixmap(matrix=mat)
    img = Image.open(BytesIO(pix.tobytes("png")))
    doc.close()
    return img


def detect_yolo_labels(img):
    """
    YOLO modeli ile cizim uzerindeki spool_label ve piece_no etiketlerini bul.
    """
    model = get_yolo_model()
    if model is None:
        return [], []
    
    img_np = np.array(img)
    results = model(img_np, conf=0.30, verbose=False)
    
    spool_labels = []
    piece_labels = []
    for box in results[0].boxes:
        cls_name = model.names[int(box.cls[0])]
        conf_val = float(box.conf[0])
        if cls_name in ('spool_label', 'piece_no'):
            x1, y1, x2, y2 = [int(v) for v in box.xyxy[0]]
            pad = 5 if cls_name == 'spool_label' else 2
            x1 = max(0, x1 - pad)
            y1 = max(0, y1 - pad)
            x2 = min(img_np.shape[1], x2 + pad)
            y2 = min(img_np.shape[0], y2 + pad)
            cropped = img_np[y1:y2, x1:x2]
            ocr_results = reader.readtext(cropped, detail=0, paragraph=False)
            raw_text = " ".join(ocr_results).strip()
            
            if cls_name == 'spool_label':
                cleaned = _clean_spool_ocr(raw_text)
                if cleaned:
                    spool_labels.append({"text": cleaned, "conf": conf_val, "box": (x1, y1, x2, y2)})
            elif cls_name == 'piece_no':
                import re
                cleaned_piece = re.sub(r'[^0-9A-Z\-]', '', raw_text)
                if cleaned_piece:
                    piece_labels.append({"text": cleaned_piece, "conf": conf_val, "box": (x1, y1, x2, y2)})
    
    return spool_labels, piece_labels


def _clean_spool_ocr(text):
    """
    OCR'dan gelen spool label metnini temizle.
    Ornek hatalar: 'SPO1' -> 'SP01', 'Po6o2' -> 'P0602', 'COOO1' -> 'C0001'
    """
    if not text:
        return ""
    # Bosluk, tire ve alt cizgi disindaki karakterleri birlestir
    text = text.replace(" ", "-")
    # Cift tireleri tekle
    text = re.sub(r'-+', '-', text)
    # OCR O (harf) <-> 0 (sifir) karisikligi: 
    # Harf+O+rakam kalibinda O'lari 0 yap (P0602 yerine Po6o2 gibi)
    text = re.sub(r'([A-Za-z])([Oo]+)(\d)', lambda m: m.group(1) + '0'*len(m.group(2)) + m.group(3), text)
    # SPO1 -> SP01, SPO2 -> SP02
    text = re.sub(r'SPO(\d)', r'SP0\1', text, flags=re.IGNORECASE)
    # Rakam+O/o+rakam kalibinda O'lari 0 yap (502433 yerine 5O2433 gibi)
    # Birden fazla kez uygula (6O2 -> 602, ama ayni zamanda 5O24O3 -> 502403)
    for _ in range(3):
        text = re.sub(r'(\d)[Oo](\d)', r'\g<1>0\g<2>', text)
    # Buyuk harfe cevir
    text = text.upper()
    return text

def extract_drawing_names(pdf_name):
    """
    Dosya adından Teknik Çizim adını ve Spool Prefix'i için Base Çizim adını çıkarır.
    Örn: '10-FG-502433-2001-P1401_1_1_Rev-0B' ->
    tech_drawing = '10-FG-502433-2001-P1401_1_1'
    base_drawing = '10-FG-502433-2001-P1401'
    """
    # 1. Revizyon bilgisini (_Rev-0B vb.) silerek Technical Drawing adını bul
    tech_drawing = re.sub(r'_Rev.*$', '', pdf_name, flags=re.IGNORECASE)
    
    # 2. Spool'lara eklenecek Base Drawing adını bul (Sadece ana isim kısmı)
    # _1_2 gibi sondaki sayisal uzantilari kaldir
    base_drawing = re.sub(r'_\d+(?:_\d+)*$', '', tech_drawing)
    
    return tech_drawing, base_drawing

def crop_region(img, region):
    """Görüntüden belirli bölgeyi kırp."""
    iw, ih = img.size
    return img.crop((
        int(iw * region["x1"]), int(ih * region["y1"]),
        int(iw * region["x2"]), int(ih * region["y2"]),
    ))


def fix_ocr_text(text):
    """Yaygın OCR hatalarını düzelt."""
    fixes = {
        "A1O5N": "A105N", "A1OSN": "A105N", "AIOSN": "A105N",
        "EITTINGS": "FITTINGS", "HTTINGS": "FITTINGS", "ETTINGS": "FITTINGS",
        "ELTTINGS": "FITTINGS", "ETTNGS": "FITTINGS", "FTTINGS": "FITTINGS",
        "BQLTS": "BOLTS", "BQLIS": "BOLTS",
        "SUPPQRTS": "SUPPORTS",
        "GASKEIS": "GASKETS",
        "ELANGES": "FLANGES",
        "MISC_": "MISC.", "MISC_": "MISC.",
    }
    for wrong, right in fixes.items():
        text = text.replace(wrong, right)
    return text


def is_category_text(text):
    """
    Kategori başlığı mı? Çok katı kontrol:
    - Metin kısa olmalı (< 35 karakter)
    - Rakam içermemeli (item code gibi)
    - Tam eşleşme veya bilinen prefix olmalı
    """
    clean = fix_ocr_text(text.strip()).upper()
    
    # Çok uzun metin kategori olamaz
    if len(clean) > 35:
        return False
    
    # Rakam içeren metin kategori olamaz (ama "IN-LINE" gibi tire OK)
    if any(c.isdigit() for c in clean):
        return False
    
    # Tam eşleşme veya içeriyorsa
    if clean in CATEGORY_KEYWORDS_EXACT:
        return True
    
    # Fuzzy matching / substring
    import re
    from difflib import SequenceMatcher
    clean_alpha = re.sub(r'[^A-Z]', '', clean)
    for kw in CATEGORY_KEYWORDS_EXACT:
        if kw in clean_alpha or clean_alpha in kw and len(clean_alpha) > 3:
            return True
        if len(clean_alpha) > 4:
            ratio = SequenceMatcher(None, clean_alpha, kw).ratio()
            if ratio > 0.8:
                return True
    
    # Prefix eşleşme (VALVES... varyasyonları)
    for prefix in CATEGORY_PREFIXES:
        if clean.startswith(prefix) or prefix.startswith(clean):
            return True
    
    # VALVES ile başlayan ve LINE/ITEM içeren varyasyonlar
    if "VALVE" in clean and ("LINE" in clean or "ITEM" in clean):
        return True
    
    return False


def is_table_title(text, title_keyword):
    """Tablo başlık satırı mı? (FABRICATION MATERIALS, ERECTION MATERIALS, CUT PIPE LENGTH)"""
    return title_keyword in text.upper()


def is_column_header(text):
    """Sütun başlığı satırı mı?"""
    upper = text.upper()
    header_kw = ["COMPONENT DESCRIPTION", "ITEM CODE", "PIECE NO", "SPOOL NO", "LENGTH (MM)"]
    return any(kw in upper for kw in header_kw)


def is_garbled_header(text):
    """
    OCR'ın yanlış okuduğu sütun başlığı satırı mı?
    Cut Pipe Length başlık satırları genellikle garbled olarak okunur.
    """
    upper = text.upper()
    garbled_kw = [
        "PIESE", "PNZE", "PVSE", "PNEE", "PIE",
        "YENGTH", "4ENGTH", "HENGTH", "LEXGTH", "ENGTH",
        "SPQOL", "SPQQL", "SPOOL NO",
        "(NS;", "(NS)", "'TB", "'TE", "'T3M",
        "REMARKS",
    ]
    match_count = sum(1 for kw in garbled_kw if kw in upper)
    # En az 2 garbled keyword varsa, bu bir başlık satırıdır
    return match_count >= 2


def is_noise_row(cols, table_name):
    """
    Gürültü satırı mı? Çizim boyut ek açıklamaları, kenar metinleri vs.
    """
    filled = [i for i, c in enumerate(cols) if c.strip()]
    if not filled:
        return True
    
    # Tüm sütunlardaki metni birleştir
    all_text = " ".join(c.strip() for c in cols if c.strip())
    all_upper = all_text.upper()
    
    # Sadece 1 sütunda veri var
    if len(filled) == 1:
        val = cols[filled[0]].strip()
        # Tek kelime header kırıntıları: REMARKS, Ro, (NS; gibi
        noise_words = {"REMARKS", "RO", "(NS;", "(NS)", "NO", "PNEE", "PIESE", 
                       "SPQOL", "SPQQL", "PVSE", "PNZE"}
        if val.upper() in noise_words:
            return True
        # Kısa sayısal değerler (2.0, 4.0, 8.0, 18.C, 26.C vs.) -> boyut annotation
        if len(val) <= 5 and any(c.isdigit() for c in val):
            if not val.replace('.', '').replace(',', '').isdigit():
                return True
            if '.' in val or ',' in val:
                return True
    
    # 2 sütunda veri, ama ikisi de çok kısa ve garbled -> gürültü
    if len(filled) <= 2:
        vals = [cols[i].strip() for i in filled]
        total_len = sum(len(v) for v in vals)
        if total_len <= 6:
            # Çok kısa metinler: "Ro" + "(NS;", "ZL)" tek başına, vs.
            has_meaningful = any(len(v) > 3 and any(c.isdigit() for c in v) for v in vals)
            if not has_meaningful:
                return True
    
    # El yazısı notasyonu kalıntıları (Spo, Spol, ZL, 1Js gibi)
    noise_patterns = {"SPO", "SPOL", "ZL)", "1JS"}
    if len(filled) <= 3 and any(all_upper.startswith(p) or p in all_upper for p in noise_patterns):
        # Ama sadece gürültü ise (spool numarası değilse)
        if not any("502" in cols[i] for i in filled):  # Spool numaraları "502" içerir
            if all(len(cols[i].strip()) <= 5 for i in filled):
                return True
        
    # Footer/Title filter
    if table_name == "CUT PIPE LENGTH":
        if any(w in all_upper for w in ["BRICATION", "TOTAL", "JTAL", "ECTION", "ERECTION", "FABRICATION"]):
            # If the row has a valid Spool Number or Length, it's not just noise, it's a valid row with an overlapping stamp.
            if len(cols) > 5 and ("502" in cols[5] or "SP" in cols[5]):
                return False
            if len(cols) > 2 and cols[2].strip().replace('.', '').isdigit():
                return False
            return True
            
    return False


def get_col_idx(x_norm, col_ranges):
    """X pozisyonuna göre sütun indeksi."""
    for i, (x_min, x_max) in enumerate(col_ranges):
        if x_min <= x_norm < x_max:
            return i
    return len(col_ranges) - 1

def adjust_col_ranges_dynamically(header_items, default_ranges):
    """
    OCR ile bulunan başlık satırındaki kelimelerin X merkezlerine bakarak
    sütun sınırlarını dinamik olarak esnetir/düzeltir.
    Üçüncü göz (Third-party) kod incelemesi sonrası refactor edildi:
    - Sadece NS ve ITEM olan tablolar değil, NS olmayan ERECTION tabloları da desteklenir.
    - Sabit indeksler (idx 2, 3) yerine dinamik arama yapılır.
    """
    ranges = list(default_ranges) # kopya
    
    x_ns = None
    x_item = None
    
    for it in header_items:
        text = it["text"].upper()
        x = it["x_norm"]
        if "N.S" in text or "NS:" in text or text == "NS":
            x_ns = x
        elif "ITEM" in text or "CODE" in text:
            x_item = x
            
    # default_ranges içinde "ITEM CODE" veya "ITEM NO" sütun indeksini bul
    item_idx = -1
    for i, r in enumerate(ranges):
        # r = (start, end, name) VEYA (start, end) olabilir (eski listelerde name yoktu)
        # Ama FAB_EREC_COLS'ta name var. Eger 3 elemanli degilse tuple genislemez.
        if len(r) >= 3 and r[2] in ["ITEM CODE", "ITEM NO"]:
            item_idx = i
            break
            
    if x_item and item_idx != -1:
        # ITEM CODE'un tahmini baslangic/bitis sinirlari
        new_start = x_item - 0.055
        new_end = x_item + 0.055
        
        # Eger N.S. de varsa ikisinin ortasi cok guvenilir bir sinirdir
        if x_ns and x_item > x_ns:
            new_start = (x_ns + x_item) / 2
            
        # ITEM CODE sinirlarini guncelle
        r_item = ranges[item_idx]
        ranges[item_idx] = (new_start, new_end, r_item[2] if len(r_item)>2 else "ITEM CODE")
        
        # Bir onceki sutunun bitisini ITEM CODE'un baslangicina daya
        if item_idx > 0:
            prev = ranges[item_idx-1]
            ranges[item_idx-1] = (prev[0], new_start, prev[2] if len(prev)>2 else "")
            
        # Bir sonraki sutunun baslangicini ITEM CODE'un bitisine daya
        if item_idx < len(ranges) - 1:
            next_col = ranges[item_idx+1]
            ranges[item_idx+1] = (new_end, max(next_col[1], new_end + 0.04), next_col[2] if len(next_col)>2 else "")
                
    return ranges


def group_by_lines(items, y_tolerance):
    """OCR öğelerini Y koordinatına göre satırlara grupla."""
    if not items:
        return []
    
    sorted_items = sorted(items, key=lambda x: x["y"])
    lines = []
    current = [sorted_items[0]]
    
    for item in sorted_items[1:]:
        if abs(item["y"] - current[0]["y"]) <= y_tolerance:
            current.append(item)
        else:
            lines.append(sorted(current, key=lambda x: x["x"]))
            current = [item]
    lines.append(sorted(current, key=lambda x: x["x"]))
    
    return lines


def parse_table_region(img, table_name):
    """
    Bir tablo bölgesini OCR ile oku ve yapısal verilere çevir.
    Faz 1 Mimarisi: General OCR + Cell-Level OCR + OpenCV Grid
    """
    region = REGIONS[table_name]
    config = TABLE_CONFIGS[table_name]
    default_ranges = config["col_ranges"]
    title_kw = config["title_keyword"]
    num_cols = len(default_ranges)
    headers = config["headers"]
    
    # Bölgeyi kırp
    cropped = crop_region(img, region)
    cw, ch = cropped.size
    img_np = np.array(cropped)
    
    # 1. GENERAL OCR
    ocr_results = reader.readtext(img_np, detail=1, paragraph=False)
    
    if not ocr_results:
        return {"headers": headers, "rows": []}
    
    header_bboxes = {"PL_NO": None, "ITEM_CODE": None, "QTY": None, "DESC": None, "LENGTH": None, "NS": None, "SPOOL": None}
    title_y = None
    
    items = []
    for bbox, text, conf in ocr_results:
        text_strip = text.strip()
        if not text_strip or conf < 0.10: continue
        x_center = (bbox[0][0] + bbox[2][0]) / 2
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        items.append({"text": fix_ocr_text(text_strip), "bbox": bbox, "xc": x_center, "yc": y_center, "conf": conf, "raw": text_strip})
        
        text_u = text_strip.upper()
        if title_kw in text_u:
            title_y = y_center
        
        # Header detection
        if "COMPONENT" in text_u and ("RS" in text_u or "P/L" in text_u):
            idx = text_u.find("COMPONENT")
            char_w = (bbox[1][0] - bbox[0][0]) / max(len(text_strip), 1)
            if idx > 0: header_bboxes["PL_NO"] = [bbox[0][0], bbox[0][0] + idx * char_w]
            else: header_bboxes["PL_NO"] = [bbox[0][0], bbox[1][0]]
        elif "P/L NO" in text_u or "PT NO" in text_u:
            header_bboxes["PL_NO"] = [bbox[0][0], bbox[1][0]]
        elif text_u == "ITEM CODE" or text_u == "ITEM NO" or "ITEM CODE" in text_u:
            if not header_bboxes["ITEM_CODE"]: header_bboxes["ITEM_CODE"] = [bbox[0][0], bbox[1][0]]
            else: header_bboxes["ITEM_CODE"] = [min(header_bboxes["ITEM_CODE"][0], bbox[0][0]), max(header_bboxes["ITEM_CODE"][1], bbox[1][0])]
        elif text_u == "QTY" or "QTY" in text_u:
            header_bboxes["QTY"] = [bbox[0][0], bbox[1][0]]
        elif "DESCRIPTION" in text_u:
            header_bboxes["DESC"] = [bbox[0][0], bbox[1][0]]
        elif "LENGTH" in text_u:
            header_bboxes["LENGTH"] = [bbox[0][0], bbox[1][0]]
        elif "N.S" in text_u or "NS" in text_u:
            header_bboxes["NS"] = [bbox[0][0], bbox[1][0]]
        elif "SPOOL" in text_u:
            header_bboxes["SPOOL"] = [bbox[0][0], bbox[1][0]]

    # 3. OpenCV Vertical Lines (Secondary)
    import cv2
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, ch // 20))
    vert_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vert_kernel, iterations=2)
    v_cnts, _ = cv2.findContours(vert_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    v_xs = [cv2.boundingRect(c)[0] + cv2.boundingRect(c)[2]//2 for c in v_cnts if cv2.boundingRect(c)[3] > ch // 10]
    
    def find_nearest_vline(target_x, search_range=20):
        cands = [x for x in v_xs if abs(x - target_x) < search_range]
        return cands[0] if cands else target_x

    # 4. Construct Column Boundaries
    col_bounds = {}
    col_names = []
    
    # Init from default ranges
    for i, rng in enumerate(default_ranges):
        col_name = rng[2] if len(rng) > 2 else f"COL_{i}"
        col_names.append(col_name)
        col_bounds[col_name] = [int(rng[0] * cw), int(rng[1] * cw)]

    # Refine ITEM CODE
    if header_bboxes["ITEM_CODE"] and "ITEM CODE" in col_bounds:
        ic_left, ic_right = header_bboxes["ITEM_CODE"]
        ic_bound_left = find_nearest_vline(ic_left - 10)
        ic_bound_right = find_nearest_vline(ic_right + 10)
        col_bounds["ITEM CODE"] = [ic_bound_left, ic_bound_right]
        # Adjust previous column
        for name in ["N.S", "DESCRIPTION"]:
            if name in col_bounds and col_bounds[name][1] > ic_bound_left:
                col_bounds[name][1] = ic_bound_left
    
    # Refine P/L NO / PT NO
    pt_no_key = "PT NO" if "PT NO" in col_bounds else ("P/L NO" if "P/L NO" in col_bounds else None)
    if header_bboxes["PL_NO"] and pt_no_key:
        pl_left, pl_right = header_bboxes["PL_NO"]
        pl_right_bound = find_nearest_vline(pl_right + 10)
        col_bounds[pt_no_key] = [max(0, pl_left - 10), pl_right_bound]
        if "DESCRIPTION" in col_bounds:
            col_bounds["DESCRIPTION"][0] = pl_right_bound
        if "LENGTH (MM)" in col_bounds:
            col_bounds["LENGTH (MM)"][0] = pl_right_bound
            
    # Refine CUT PIPE LENGTH columns
    if table_name == "CUT PIPE LENGTH":
        # Based on exact data token X-coordinates from PDF:
        # PIECE NO (~201), PT NO (~236), LENGTH (~297), N.S. (~361), ITEM NO (~413), SPOOL (~507)
        # We snap to the nearest vertical lines between these columns:
        
        def map_reference_x(ref_x):
            # Reference width (cw) was 857 at ZOOM=2
            return int(ref_x * (cw / 857.0))
            
        pt_start = find_nearest_vline(map_reference_x(218))
        len_start = find_nearest_vline(map_reference_x(293))
        ns_start = find_nearest_vline(map_reference_x(344))
        item_start = find_nearest_vline(map_reference_x(400))
        spool_start = find_nearest_vline(map_reference_x(490))
        
        # Enforce Contiguous Boundaries
        if "PIECE NO" in col_bounds:
            col_bounds["PIECE NO"][1] = pt_start
        if "PT NO" in col_bounds:
            col_bounds["PT NO"][0] = pt_start
            col_bounds["PT NO"][1] = len_start
        if "LENGTH (MM)" in col_bounds:
            col_bounds["LENGTH (MM)"][0] = len_start
            col_bounds["LENGTH (MM)"][1] = ns_start
        if "N.S. (INS)" in col_bounds:
            col_bounds["N.S. (INS)"][0] = ns_start
            col_bounds["N.S. (INS)"][1] = item_start
        if "ITEM NO" in col_bounds:
            col_bounds["ITEM NO"][0] = item_start
            col_bounds["ITEM NO"][1] = spool_start
        if "SPOOL NO" in col_bounds:
            col_bounds["SPOOL NO"][0] = spool_start

    # Refine QTY
    if header_bboxes["QTY"] and "QTY" in col_bounds:
        q_left, q_right = header_bboxes["QTY"]
        q_left_bound = find_nearest_vline(q_left - 10)
        q_right_bound = find_nearest_vline(q_right + 10)
        col_bounds["QTY"] = [q_left_bound, q_right_bound]
        if "ITEM CODE" in col_bounds:
            col_bounds["ITEM CODE"][1] = min(col_bounds["ITEM CODE"][1], q_left_bound)

    # Group rows
    if title_y is None: return {"headers": headers, "rows": []}
    data_items = [it for it in items if it["yc"] > title_y + 10]
    
    lines = []
    data_items.sort(key=lambda x: x["yc"])
    if not data_items: return {"headers": headers, "rows": []}
    
    current = [data_items[0]]
    for it in data_items[1:]:
        if abs(it["yc"] - current[0]["yc"]) <= 10:
            current.append(it)
        else:
            lines.append(sorted(current, key=lambda x: x["xc"]))
            current = [it]
    lines.append(sorted(current, key=lambda x: x["xc"]))
    
    result_rows = []
    
    def cell_ocr(x_min, x_max, row_y_min, row_y_max, allow=''):
        crop_x1, crop_x2 = max(0, int(x_min)), min(cw, int(x_max))
        crop_y1, crop_y2 = max(0, int(row_y_min)-4), min(ch, int(row_y_max)+4)
        
        # Add padding so easyocr can read text touching vertical lines
        pad_x = 8
        cx1 = max(0, int(crop_x1) - pad_x)
        cx2 = min(img_np.shape[1], int(crop_x2) + pad_x)
        
        if crop_x2 <= crop_x1 or crop_y2 <= crop_y1: return ""
        cell = img_np[crop_y1:crop_y2, cx1:cx2]
        if cell.size == 0: return ""
        scaled = cv2.resize(cell, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
        gray_cell = cv2.cvtColor(scaled, cv2.COLOR_RGB2GRAY)
        res = reader.readtext(gray_cell, allowlist=allow, detail=1)
        if not res: return ""
        valid_res = []
        cell_h = (crop_y2 - crop_y1) * 3
        target_y = cell_h / 2
        
        for b, t, c in res:
            if c < 0.2: continue
            t_strip = t.strip()
            if not t_strip: continue
            if allow == '0123456789' and not any(ch.isdigit() for ch in t_strip):
                continue
            
            # Reject tokens that are bleeding from the adjacent column on the right
            token_x_center = (b[0][0] + b[1][0]) / 2
            if token_x_center > cell.shape[1] * 3 * 0.85:
                continue
            
            token_y_center = (b[0][1] + b[2][1]) / 2
            dist = abs(token_y_center - target_y)
            valid_res.append((b, t_strip, c, dist))
            
        if not valid_res: return ""
        # Sort by distance to center (ascending), then confidence (descending)
        valid_res.sort(key=lambda x: (x[3], -x[2]))
        
        best_val = valid_res[0][1]
        print(f"DEBUG cell_ocr({x_min}, {x_max}, {row_y_min}, {row_y_max}) -> {best_val}")
        return best_val

    print(f"DEBUG TABLE {table_name} BOUNDS: {col_bounds}")

    for line_items in lines:
        if len(line_items) < 1: continue
        
        line_text = " ".join([it["text"] for it in line_items])
        if is_column_header(line_text) or is_garbled_header(line_text) or is_table_title(line_text, title_kw):
            continue
            
        avg_conf = sum(it["conf"] for it in line_items) / len(line_items)
        if avg_conf < 0.15: continue
        
        # Kategori kontrolü
        if len(line_items) <= 2:
            if is_category_text(line_text):
                result_rows.append(("category", line_text.strip().upper()))
                continue
                
        row_y_min = min(it["bbox"][0][1] for it in line_items)
        row_y_max = max(it["bbox"][2][1] for it in line_items)
        
        cols = [""] * num_cols
        for i, col_name in enumerate(col_names):
            bound = col_bounds[col_name]
            
            # Cell-Level OCR overrides for critical columns
            if col_name == pt_no_key:
                val = cell_ocr(bound[0], bound[1], row_y_min, row_y_max, '0123456789')
                if val: cols[i] = val
                continue
            if col_name == "LENGTH (MM)":
                val = cell_ocr(bound[0], bound[1], row_y_min, row_y_max, '0123456789.')
                if val: cols[i] = val
                continue
            if col_name == "QTY":
                val = cell_ocr(bound[0], bound[1], row_y_min, row_y_max, '0123456789.')
                if val: cols[i] = val
                continue
                
            # General OCR matching for others
            col_text = ""
            for it in line_items:
                if bound[0] <= it["xc"] < bound[1]:
                    raw = it["text"]
                    if col_name == "ITEM CODE":
                        pass
                    col_text += raw + " "
            cols[i] = col_text.strip()
            
            # Fallback Cell-Level OCR for Item Code if empty but shouldn't be
            if col_name == "ITEM CODE" and not cols[i]:
                val = cell_ocr(bound[0], bound[1], row_y_min, row_y_max, '0123456789')
                if val: cols[i] = val

        if is_noise_row(cols, table_name):
            if any(c for c in cols):
                result_rows.append(("NOISE", cols))
            continue
            
        # Row validation for CUT PIPE LENGTH
        if table_name == "CUT PIPE LENGTH":
            has_item = bool(cols[4].strip())  # ITEM NO is at index 4
            has_spool = bool(cols[5].strip()) # SPOOL NO is at index 5
            has_pt = bool(cols[1].strip())    # PT NO is at index 1
            
            if not has_item and not has_spool and not has_pt:
                if any(c for c in cols):
                    result_rows.append(("NOISE", cols))
                continue
            elif not (has_item and has_spool):
                result_rows.append(("INCOMPLETE", cols))
                continue
            else:
                result_rows.append(("data", cols))
                continue
            
        if any(c for c in cols):
            result_rows.append(("data", cols))
    
    # Devam satırlarını birleştir
    merged = _merge_continuations(result_rows)
    
    return {"headers": headers, "rows": merged}


def _merge_continuations(rows):
    """
    Devam satırlarını üst satırla birleştir.
    Kural: Eğer bir satırda:
    - İlk sütun (numara) boş
    - Sadece 1-2 sütunda veri var  
    - Veri sadece description sütununda (idx 1)
    O zaman üst veri satırıyla birleştir (çünkü uzun açıklamanın devamı).
    """
    if not rows:
        return rows
    
    merged = []
    
    for rtype, rdata in rows:
        if rtype == "category":
            merged.append((rtype, rdata))
            continue
        
        # İlk sütun boş mu?
        first_col = rdata[0].strip()
        
        # Hangi sütunlarda veri var?
        filled = [(i, rdata[i]) for i in range(len(rdata)) if rdata[i].strip()]
        
        # Devam satırı mı? Koşullar:
        # 1. İlk sütun boş (numara yok)
        # 2. Sadece 1 sütunda veri var
        # 3. O sütun description sütunu (idx 1)
        # 4. Üstte bir veri satırı var
        is_continuation = (
            not first_col
            and len(filled) == 1
            and filled[0][0] == 1
            and merged
        )
        
        if is_continuation:
            # Üst veri satırını bul
            for idx in range(len(merged) - 1, -1, -1):
                if merged[idx][0] == "data":
                    prev_data = list(merged[idx][1])
                    # Description'a ekle
                    if prev_data[1]:
                        prev_data[1] += " " + rdata[1]
                    else:
                        prev_data[1] = rdata[1]
                    merged[idx] = ("data", prev_data)
                    break
            else:
                merged.append((rtype, rdata))
        else:
            merged.append((rtype, rdata))
    
    return merged


# ============================================================================
# EXCEL YAZICI
# ============================================================================

FONT_PDF = Font(name='Calibri', bold=True, size=13, color="FFFFFF")
FILL_PDF = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")

FONT_TABLE = Font(name='Calibri', bold=True, size=11, color="FFFFFF")
FILL_TABLE = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

FONT_COL = Font(name='Calibri', bold=True, size=10)
FILL_COL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

FONT_CAT = Font(name='Calibri', bold=True, size=10, color="1F4E79")
FILL_CAT = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

FONT_DATA = Font(name='Calibri', size=10)
FILL_ODD = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FILL_EVEN = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

BORDER = Border(
    left=Side(style='thin', color='B4C6E7'),
    right=Side(style='thin', color='B4C6E7'),
    top=Side(style='thin', color='B4C6E7'),
    bottom=Side(style='thin', color='B4C6E7'),
)

AL_C = Alignment(horizontal='center', vertical='center')
AL_L = Alignment(horizontal='left', vertical='center', wrap_text=True)


def write_excel(all_data):
    """Tüm verileri tek Excel, tek sheet'e yaz."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tablo Verileri"
    
    row = 1
    
    for pdf_data in all_data:
        # PDF BAŞLIĞI
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        c = ws.cell(row=row, column=1, value=pdf_data["pdf_name"])
        c.font = FONT_PDF
        c.fill = FILL_PDF
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[row].height = 30
        row += 1
        
        has_any_table = False
        
        for tname in ["FABRICATION MATERIALS", "ERECTION MATERIALS", "CUT PIPE LENGTH"]:
            td = pdf_data["tables"].get(tname)
            if not td or not td["rows"]:
                continue
            
            has_any_table = True
            
            # TABLO BAŞLIĞI
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
            c = ws.cell(row=row, column=1, value=f"  {tname}")
            c.font = FONT_TABLE
            c.fill = FILL_TABLE
            c.alignment = Alignment(horizontal='left', vertical='center')
            ws.row_dimensions[row].height = 24
            row += 1
            
            # SÜTUN BAŞLIKLARI
            for ci, h in enumerate(td["headers"], 1):
                c = ws.cell(row=row, column=ci, value=h)
                c.font = FONT_COL
                c.fill = FILL_COL
                c.alignment = AL_C
                c.border = BORDER
            ws.row_dimensions[row].height = 22
            row += 1
            
            # VERİ
            di = 0
            for rtype, rdata in td["rows"]:
                if rtype == "category":
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
                    c = ws.cell(row=row, column=1, value=rdata)
                    c.font = FONT_CAT
                    c.fill = FILL_CAT
                    c.alignment = AL_L
                    c.border = BORDER
                else:
                    fill = FILL_EVEN if di % 2 == 0 else FILL_ODD
                    for ci, val in enumerate(rdata[:7], 1):
                        c = ws.cell(row=row, column=ci, value=val)
                        c.font = FONT_DATA
                        c.fill = fill
                        c.border = BORDER
                        c.alignment = AL_C if ci in [1, 3, 6, 7] else AL_L
                    di += 1
                row += 1
            
            row += 1  # Tablolar arası boşluk
        
        if not has_any_table:
            row -= 1  # Boş PDF başlığını geri al
        
        row += 1  # PDF grupları arası boşluk
    
    # SÜTUN GENİŞLİKLERİ
    for i, w in enumerate([12, 48, 12, 18, 22, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    ws.sheet_properties.tabColor = "2F5496"
    wb.save(OUTPUT_FILE)
    print(f"\nExcel kaydedildi: {OUTPUT_FILE}")


# ============================================================================
# ANA İŞLEM
# ============================================================================

def process_pdf(pdf_path, pdf_name):
    """Tek bir PDF dosyasını işle."""
    print(f"\n  Isleniyor: {pdf_name}")
    
    img = pdf_to_image(pdf_path)
    
    # Dosya adından teknik çizim ve base çizim bilgilerini al
    tech_drawing, base_drawing = extract_drawing_names(pdf_name)
    print(f"    Teknik Cizim No: {tech_drawing}")
    
    tables = {}
    
    # YOLO ile spool_label ve piece_no tespiti
    yolo_spools, yolo_pieces = detect_yolo_labels(img)
    if yolo_spools:
        print(f"    YOLO Spool Tespiti: {len(yolo_spools)} adet bulundu")
        for sp in yolo_spools:
            print(f"      -> {sp['text']} (conf={sp['conf']:.2f})")
    else:
        print(f"    YOLO Spool Tespiti: Bulunamadi (eski yontem kullanilacak)")
        
    if yolo_pieces:
        print(f"    YOLO Piece Tespiti: {len(yolo_pieces)} adet bulundu")
    
    for tname in ["FABRICATION MATERIALS", "ERECTION MATERIALS", "CUT PIPE LENGTH"]:
        print(f"    {tname}...", end=" ", flush=True)
        
        td = parse_table_region(img, tname)
        
        # APPLY PHASE 2A: SPOOL CANONICALIZATION
        if tname == "CUT PIPE LENGTH":
            for i, row in enumerate(td["rows"]):
                if row[0] in ["data", "INCOMPLETE"]:
                    cols = row[1]
                    raw_spool = cols[5]
                    canon = canonicalize_spool(raw_spool, base_drawing, yolo_spools)
                    
                    # Store piece no validation as well for Phase 2B prep
                    piece_no_val = validate_piece_no(cols[0])
                    
                    # Convert to rich objects
                    cols[0] = piece_no_val
                    cols[5] = canon
                    
        dc = len([r for r in td["rows"] if r[0] == "data"])
        cc = len([r for r in td["rows"] if r[0] == "category"])
        print(f"{dc} veri, {cc} kategori")
        
        tables[tname] = td
    
    return {
        "pdf_name": pdf_name,
        "technical_drawing": tech_drawing,
        "base_drawing": base_drawing,
        "tables": tables,
        "yolo_spools": yolo_spools,
        "yolo_pieces": yolo_pieces,
    }


def extract_bom_data_to_json(pdf_path, pdf_name):
    """PDF işler ve Excel şemasına uygun düz (flat) bir JSON listesi döner."""
    result = process_pdf(pdf_path, pdf_name)
    tables = result["tables"]
    tech_drawing = result.get("technical_drawing", pdf_name)
    base_drawing = result.get("base_drawing", pdf_name)
    yolo_spools = result.get("yolo_spools", [])
    yolo_pieces = result.get("yolo_pieces", [])
    
    # 1. Mock Master DB for Domain Normalization
    MASTER_CODES = {"2467784", "535745", "2467617", "2467614", "15722", "11196253"}

    def normalize_item_code(raw_code):
        raw_code = raw_code.strip()
        candidate = raw_code
        status = "RAW"
        rule = None
        evidence = []
        
        if len(raw_code) == 8 and raw_code.startswith("1"):
            hyp_candidate = raw_code[1:]
            if hyp_candidate in MASTER_CODES:
                candidate = hyp_candidate
                rule = "FABRICATION_PREFIX_HYPOTHESIS"
                evidence.append("Length is 8 and starts with 1")
                evidence.append(f"Candidate {hyp_candidate} exact match in MASTER_CODES")
                status = "VERIFIED_CANDIDATE"
            else:
                candidate = raw_code
                rule = "FABRICATION_PREFIX_HYPOTHESIS_FAILED"
                evidence.append("Length is 8 and starts with 1")
                evidence.append(f"Hypothesis {hyp_candidate} NOT found in MASTER_CODES. Reverting to raw.")
                status = "RAW_FALLBACK"
                
        return {
            "raw": raw_code,
            "candidate": candidate,
            "rule": rule,
            "evidence": evidence,
            "status": status
        }

    def mock_spatial_association(pl_no):
        # We integrate the real YOLO spatial association here
        import math
        norm_pl_no = re.sub(r'[^0-9A-Z\-]', '', str(pl_no))
        if norm_pl_no and yolo_spools and yolo_pieces:
            matching_pieces = [p for p in yolo_pieces if p["text"] == norm_pl_no]
            if matching_pieces:
                # Use the first matched piece marker
                p = matching_pieces[0]
                px, py = (p["box"][0]+p["box"][2])/2, (p["box"][1]+p["box"][3])/2
                best_dist = float('inf')
                best_spool = ""
                for s in yolo_spools:
                    sx, sy = (s["box"][0]+s["box"][2])/2, (s["box"][1]+s["box"][3])/2
                    dist = math.hypot(px - sx, py - sy)
                    if dist < best_dist:
                        best_dist = dist
                        best_spool = s["text"]
                
                if best_spool and best_dist < 600:
                    # Format spool
                    f_sp = best_spool
                    if base_drawing not in f_sp:
                        if "SP" in f_sp:
                            sp_match = re.search(r'(SP\d+)', f_sp)
                            if sp_match:
                                f_sp = f"{base_drawing}-{sp_match.group(1)}"
                            else:
                                f_sp = f"{base_drawing}-{f_sp}"
                    
                    return {
                        "spool": f_sp,
                        "method": "SPATIAL_PROXIMITY_AND_MARKER",
                        "score": 0.85,
                        "status": "SPATIAL_HIGH_CONFIDENCE",
                        "reason": f"Piece marker {norm_pl_no} localized, nearest spool {f_sp} within margin"
                    }
        
        return {
            "spool": None,
            "method": "SPATIAL_WEAK",
            "score": 0.0,
            "status": "UNRESOLVED",
            "reason": "Weak spatial candidate, no localized piece marker found"
        }

    # 2. Build CUT PIPE Lookup
    cut_pipes = []
    cut_dict = {}
    if "CUT PIPE LENGTH" in tables and tables["CUT PIPE LENGTH"]["rows"]:
        for rtype, rdata in tables["CUT PIPE LENGTH"]["rows"]:
            if rtype == "data" and len(rdata) >= 6:
                p_dict = rdata[0] if isinstance(rdata[0], dict) else {}
                piece_no = p_dict.get("normalized", p_dict.get("raw", rdata[0]))
                if isinstance(piece_no, dict): piece_no = piece_no.get("normalized", piece_no.get("raw", ""))
                piece_no = str(piece_no).strip()
                
                pt_no = str(rdata[1]).strip()
                pt_match = re.search(r'<([A-Za-z0-9-]+)>', pt_no)
                if pt_match: pt_no = pt_match.group(1)
                
                raw_length_str = re.sub(r'<[^>]*>', '', str(rdata[2])).strip()
                length = raw_length_str
                
                item_no = str(rdata[4]).strip()
                
                sp_dict = rdata[5] if isinstance(rdata[5], dict) else {}
                spool_no = sp_dict.get("canonical_spool", sp_dict.get("raw_spool", str(rdata[5])))
                if not spool_no or isinstance(spool_no, dict): spool_no = str(rdata[5]).strip()
                
                c_pipe = {
                    "piece_no": piece_no,
                    "pt_no": pt_no,
                    "item_no": item_no,
                    "spool_no": spool_no,
                    "length": length
                }
                cut_pipes.append(c_pipe)
                if item_no not in cut_dict: cut_dict[item_no] = []
                cut_dict[item_no].append(c_pipe)

    rows = []
    
    # 3. Association Engine for Fabrication Materials
    for tname in ["FABRICATION MATERIALS", "ERECTION MATERIALS"]:
        if tname not in tables or not tables[tname]["rows"]:
            continue
            
        current_category = "UNKNOWN"
        for rtype, rdata in tables[tname]["rows"]:
            if rtype == "category":
                current_category = rdata
                continue
                
            if rtype == "data" and len(rdata) >= 7:
                pl_raw = rdata[0]
                if isinstance(pl_raw, dict): pl_raw = pl_raw.get("normalized", pl_raw.get("raw", ""))
                pl_no = str(pl_raw).strip()
                
                desc = str(rdata[1]).strip()
                raw_item = str(rdata[3]).strip()
                qty = str(rdata[5]).strip()
                weight = str(rdata[6]).strip()
                
                # SUPPORTS filtresi
                if "SUPPORT" in current_category.upper() or raw_item.upper().startswith("X"):
                    continue
                    
                # A. Domain Normalization
                norm_data = normalize_item_code(raw_item)
                lookup_code = norm_data["candidate"]
                
                trace = {
                    "piece_no": pl_no,
                    "item_code_raw": norm_data["raw"],
                    "item_code_normalized": lookup_code,
                    "normalization_status": norm_data["status"],
                    "category": current_category,
                    "spool": None,
                    "method": None,
                    "score": 0.0,
                    "status": "UNRESOLVED",
                    "reason": ""
                }
                
                # B. Category Routing & Association
                if "PIPE" in current_category.upper():
                    candidates = cut_dict.get(lookup_code, [])
                    if len(candidates) == 1:
                        trace["spool"] = candidates[0]["spool_no"]
                        trace["method"] = "DETERMINISTIC_CROSS_REF"
                        trace["score"] = 1.0
                        trace["status"] = "STRONG_DETERMINISTIC"
                        trace["reason"] = f"Exact item match '{lookup_code}' in CUT PIPE with unique spool candidate"
                        
                        # Add Pipe sub-pieces as well
                        row_dict = {
                            "id": "",
                            "technical_drawing": tech_drawing,
                            "assembly": candidates[0]["spool_no"],
                            "assembly_description": "",
                            "assembly_weight": "",
                            "assembly_qty": "1",
                            "sub_assembly": candidates[0]["pt_no"],
                            "sub_assembly_defination": desc,
                            "item_code": norm_data["raw"],
                            "qty": "1.0",
                            "unit_weight": weight,
                            "pose_no": "",
                            "giris_yapan": "",
                            "giris_tarihi": "",
                            "_debug": trace
                        }
                        rows.append(row_dict)
                        continue # Pipe parçalarını alt parçalar olarak ekledik, ana Pipe satırını atla
                        
                    elif len(candidates) > 1:
                        trace["spool"] = "BİLİNMEYEN (MULTIPLE CUT PIPE)"
                        trace["method"] = "DETERMINISTIC_CROSS_REF"
                        trace["score"] = 0.5
                        trace["status"] = "AMBIGUOUS"
                        trace["reason"] = f"Multiple CUT PIPE candidates ({len(candidates)}) found for item '{lookup_code}'"
                    else:
                        spatial_res = mock_spatial_association(pl_no)
                        trace["spool"] = spatial_res["spool"] if spatial_res["spool"] else "BİLİNMEYEN (SPATIAL FAIL)"
                        trace["method"] = f"SPATIAL_FALLBACK ({spatial_res['method']})"
                        trace["score"] = spatial_res["score"]
                        trace["status"] = spatial_res["status"]
                        trace["reason"] = f"No CUT PIPE match for '{lookup_code}'. Fallback: {spatial_res['reason']}"
                        
                elif any(c in current_category.upper() for c in ["FITTNGS", "FITTINGS", "FLANGES"]):
                    spatial_res = mock_spatial_association(pl_no)
                    trace["spool"] = spatial_res["spool"] if spatial_res["spool"] else "BİLİNMEYEN (SPATIAL FAIL)"
                    trace["method"] = spatial_res["method"]
                    trace["score"] = spatial_res["score"]
                    trace["status"] = spatial_res["status"]
                    trace["reason"] = f"Category '{current_category}' bypasses CUT PIPE. {spatial_res['reason']}"
                else:
                    trace["spool"] = "BİLİNMEYEN (UNKNOWN CAT)"
                    trace["status"] = "UNRESOLVED"
                    trace["reason"] = f"Unknown category '{current_category}'"
                
                # Emit row
                try:
                    total_qty = int(float(qty))
                except:
                    total_qty = 1
                    
                row_dict = {
                    "id": "",
                    "technical_drawing": tech_drawing,
                    "assembly": trace["spool"],
                    "assembly_description": "",
                    "assembly_weight": "",
                    "assembly_qty": "1",
                    "sub_assembly": re.sub(r'[^0-9A-Z\-]', '', str(pl_no)),
                    "sub_assembly_defination": desc,
                    "item_code": norm_data["raw"],
                    "qty": str(total_qty) + ".0",
                    "unit_weight": weight,
                    "pose_no": "",
                    "giris_yapan": "",
                    "giris_tarihi": "",
                    "_debug": trace
                }
                rows.append(row_dict)

    return rows
def main():
    print("=" * 60)
    print("  PDF'den Tablo Cikarma Sistemi v3")
    print("=" * 60)
    
    if not os.path.exists(PDF_FOLDER):
        print(f"\nHATA: '{PDF_FOLDER}' bulunamadi!")
        sys.exit(1)
    
    pdf_files = sorted([f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")])
    if not pdf_files:
        print(f"\nHATA: PDF bulunamadi!")
        sys.exit(1)
    
    print(f"\n{len(pdf_files)} PDF bulundu:")
    for f in pdf_files:
        print(f"  - {f}")
    
    all_data = []
    for fn in pdf_files:
        try:
            result = process_pdf(os.path.join(PDF_FOLDER, fn), fn.replace('.pdf', ''))
            all_data.append(result)
        except Exception as e:
            print(f"\n  HATA: {fn}: {e}")
            import traceback
            traceback.print_exc()
    
    if all_data:
        write_excel(all_data)
        print(f"\n{'='*60}")
        print(f"  TAMAMLANDI! {len(all_data)} PDF -> {OUTPUT_FILE}")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
