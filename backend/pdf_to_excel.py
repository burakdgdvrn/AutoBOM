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
        "col_ranges": [(c[0], c[1]) for c in FAB_EREC_COLS],
        "title_keyword": "FABRICATION MATERIALS",
    },
    "ERECTION MATERIALS": {
        "headers": [c[2] for c in FAB_EREC_COLS],
        "col_ranges": [(c[0], c[1]) for c in FAB_EREC_COLS],
        "title_keyword": "ERECTION MATERIALS",
    },
    "CUT PIPE LENGTH": {
        "headers": [c[2] for c in CUT_PIPE_COLS],
        "col_ranges": [(c[0], c[1]) for c in CUT_PIPE_COLS],
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
    
    # Tam eşleşme
    if clean in CATEGORY_KEYWORDS_EXACT:
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
    """
    region = REGIONS[table_name]
    config = TABLE_CONFIGS[table_name]
    col_ranges = config["col_ranges"]
    title_kw = config["title_keyword"]
    num_cols = len(col_ranges)
    headers = config["headers"]
    
    # Bölgeyi kırp
    cropped = crop_region(img, region)
    cw, ch = cropped.size
    
    # OCR
    ocr_results = reader.readtext(np.array(cropped), detail=1, paragraph=False)
    
    if not ocr_results:
        return {"headers": headers, "rows": []}
    
    pl_no_header_bounds = None
    if table_name == "FABRICATION MATERIALS":
        for bbox, text, conf in ocr_results:
            text_upper = text.upper()
            if "COMPONENT" in text_upper and ("RS" in text_upper or "P/L" in text_upper):
                total_len = len(text)
                idx = text_upper.find("COMPONENT")
                x1, x2 = bbox[0][0], bbox[1][0]
                char_w = (x2 - x1) / max(total_len, 1)
                pl_no_header_bounds = (int(x1), int(x1 + (idx * char_w)))
                break
            elif "P/L NO" in text_upper:
                x1, x2 = bbox[0][0], bbox[1][0]
                pl_no_header_bounds = (int(x1), int(x2))
                break
                
    # OCR sonuçlarını işle
    items = []
    for bbox, text, conf in ocr_results:
        text = text.strip()
        if not text or conf < 0.10:
            continue
        x_center = (bbox[0][0] + bbox[2][0]) / 2
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        items.append({
            "x": x_center,
            "y": y_center,
            "x_norm": x_center / cw,
            "text": fix_ocr_text(text),
            "conf": conf,
            "bbox": bbox
        })
    
    if not items:
        return {"headers": headers, "rows": []}
    
    # Tablo başlığı satırını bul ve altındaki içeriği al
    title_y = None
    for item in items:
        if title_kw in item["text"].upper():
            title_y = item["y"]
            break
    
    # Başlık bulunamazsa, tablo boş olabilir
    if title_y is None:
        return {"headers": headers, "rows": []}
    
    # Sadece başlık satırının ALTINDAKI öğeleri al
    data_items = [it for it in items if it["y"] > title_y + 10]
    
    if not data_items:
        return {"headers": headers, "rows": []}
    
    # Y toleransı: satır yüksekliğinin yarısı (~8-10 px)
    y_tol = ch * 0.012
    if y_tol < 8:
        y_tol = 8
    
    # Satırlara grupla
    lines = group_by_lines(data_items, y_tol)
    
    # İlk satır sütun başlığı ise atla
    if lines:
        first_line_text = " ".join([it["text"] for it in lines[0]])
        if is_column_header(first_line_text):
            col_ranges = adjust_col_ranges_dynamically(lines[0], col_ranges)
            lines = lines[1:]
    
    # Her satırı işle
    result_rows = []
    
    for line_items in lines:
        # Satır metnini oluştur
        line_text = " ".join([it["text"] for it in line_items])
        
        # Sütun başlığı satırlarını atla
        if is_column_header(line_text):
            continue
        
        # Garbled (bozuk OCR) başlık satırlarını atla
        if is_garbled_header(line_text):
            continue
        
        # Tablo başlığı satırlarını atla
        if is_table_title(line_text, title_kw):
            continue
        
        # Çok düşük güvenilirlikli satırları atla
        avg_conf = sum(it["conf"] for it in line_items) / len(line_items)
        if avg_conf < 0.15:
            continue
        
        # Tablo dışı metinleri filtrele (sol kenardan gelen çizim metinleri)
        table_items = [it for it in line_items if it["x_norm"] >= 0.02]
        if not table_items:
            continue
        
        # Kategori kontrolü
        if len(table_items) <= 2:
            combined = " ".join([it["text"] for it in table_items])
            if is_category_text(combined):
                result_rows.append(("category", combined.strip().upper()))
                continue
        
        if len(table_items) == 1:
            if is_category_text(table_items[0]["text"]):
                result_rows.append(("category", table_items[0]["text"].strip().upper()))
                continue
        
        # Sütunlara ata
        cols = [""] * num_cols
        col_texts = {i: [] for i in range(num_cols)}
        
        for item in sorted(table_items, key=lambda x: x["x"]):
            ci = get_col_idx(item["x_norm"], col_ranges)
            col_texts[ci].append(item["text"])
        
        for i in range(num_cols):
            cols[i] = " ".join(col_texts[i]).strip()
            
        # P/L NO Fallback OCR
        if table_name == "FABRICATION MATERIALS" and not cols[0] and pl_no_header_bounds:
            row_y_min = min((it["bbox"][0][1] for it in table_items), default=-1)
            row_y_max = max((it["bbox"][2][1] for it in table_items), default=-1)
            if row_y_min >= 0:
                import cv2
                crop_x1 = max(0, pl_no_header_bounds[0] - 8)
                crop_x2 = min(cw, pl_no_header_bounds[1] + 8)
                crop_y1 = max(0, int(row_y_min) - 6)
                crop_y2 = min(ch, int(row_y_max) + 6)
                
                if crop_x2 > crop_x1 and crop_y2 > crop_y1:
                    sub_crop = np.array(cropped)[crop_y1:crop_y2, crop_x1:crop_x2]
                    if sub_crop.size > 0:
                        upscaled = cv2.resize(sub_crop, (0, 0), fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
                        gray = cv2.cvtColor(upscaled, cv2.COLOR_RGB2GRAY)
                        
                        fallback_res = reader.readtext(gray, allowlist='0123456789', detail=1)
                        if fallback_res:
                            best_fallback = max(fallback_res, key=lambda x: x[2])
                            cols[0] = best_fallback[1]
                            
                            debug_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scratch", "debug", "pl_no")
                            os.makedirs(debug_dir, exist_ok=True)
                            cv2.imwrite(os.path.join(debug_dir, f"crop_row_{crop_y1}_res_{cols[0]}.png"), gray)
        
        # Gürültü satırlarını atla
        if is_noise_row(cols, table_name):
            continue
        
        # En az bir sütunda anlamlı veri varsa ekle
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
    
    rows = []
    
    def clean_item_code(code):
        code = str(code).strip()
        if not code: return code
        
        # OCR Hata Düzeltmeleri
        code = re.sub(r'([A-Z])[Il1][Oo]([A-Z\-])', r'\g<1>10\g<2>', code) # P1O, PIO -> P10
        code = re.sub(r'([A-Z])[Il](0)([A-Z\-])', r'\g<1>10\g<3>', code)   # PI0 -> P10
        
        # Dinamik sütun ayarlaması sayesinde artık Item Code'ların başındaki "I" 
        # (ki OCR bunu genellikle 1, l veya I olarak okur) kesilmeden tam geliyor.
        # Örneğin: PDF'de I1209668 yazıyor, OCR 11209668 okuyor.
        # ERP_STRIP_PREFIX açıksa bu baştaki karakteri atıyoruz.
        if ERP_STRIP_PREFIX:
            # Baştaki 1, I veya l harfini (ve varsa yanındaki 0'ı) sil
            code = re.sub(r'^[1Il]0?', '', code)
            return code
            
        # Eğer şirket öneki tutmak istiyorsa, OCR'ın okuduğu 1'i I yap
        code = re.sub(r'^[1l](\d{3,})', r'I\1', code)
            
        return code

    # 1. CUT PIPE LENGTH verilerini parse edip Pipe'lar için Assembly ve SubAssembly bulalım
    cut_pipes = []
    last_piece_no = "" # Forward-fill için
    if "CUT PIPE LENGTH" in tables and tables["CUT PIPE LENGTH"]["rows"]:
        for rtype, rdata in tables["CUT PIPE LENGTH"]["rows"]:
            if rtype == "data" and len(rdata) >= 6:
                piece_no, pt_no, length, ns, item_no, spool_no = rdata[:6]
                
                # Sütunlar bazen birbirine karışabiliyor, PT NO'yu (<1-1> veya <5> gibi) bulmak için hepsine bakalım
                combined_text = f"{piece_no} {pt_no} {length} {ns}".replace("'", "").replace('"', '')
                pt_match = re.search(r'<([A-Za-z0-9-]+)>', combined_text)
                if pt_match:
                    pt_no = pt_match.group(1)
                else:
                    pt_no = pt_no.strip()

                # LENGTH'i bulalım (sadece sayılardan oluşan kısmı al)
                raw_length_str = re.sub(r'<[^>]*>', '', length).strip()
                if not raw_length_str and ns:
                    # length boşaldıysa ns içinden length'i alalım
                    nums = re.findall(r'\b\d{2,5}\b', ns)
                    if nums:
                        length = nums[0]
                    else:
                        length = ns.strip()
                elif raw_length_str:
                    # length içinde rakam kaldıysa (örn: "1183" veya "M44812000")
                    nums = re.findall(r'\b\d{2,5}\b', raw_length_str)
                    if nums:
                        length = nums[0]
                    else:
                        match = re.search(r'\d{2,5}', raw_length_str)
                        length = match.group(0) if match else raw_length_str

                # PIECE NO'dan <1-1> gibi kısımları temizle
                piece_no = re.sub(r'<[^>]*>', '', piece_no).strip()

                # PIECE NO Forward-fill mantığı (alt satırlarda boşsa üsttekinden alır)
                if piece_no:
                    last_piece_no = piece_no
                else:
                    piece_no = last_piece_no
                
                # Item No temizliği
                item_no = clean_item_code(item_no)
                
                # Spool no'yu düzenle (örn: 10-FG...-SPO1 -> 10-FG...-SP01)
                spool_no = spool_no.replace('PO', 'P0').replace('SPO', 'SP0')
                # OCR O (harf) ile 0 (sıfır) karıştırması: C0001 -> COOO1 gibi
                # İçinde harf+O+ olan yerleri bulup O'ları 0'a çeviriyoruz (COO01 -> C0001)
                spool_no = re.sub(r'([A-Z])([O]+)(\d*)', lambda m: m.group(1) + '0'*len(m.group(2)) + m.group(3), spool_no)
                
                # Akıllı Spool Ön Ek Eklemesi
                if spool_no and base_drawing not in spool_no:
                    # Tabloda sadece SP01 gibi kısa bir kod varsa
                    if "SP" in spool_no:
                        sp_match = re.search(r'(SP\d+)', spool_no)
                        if sp_match:
                            spool_no = f"{base_drawing}-{sp_match.group(1)}"
                        else:
                            spool_no = f"{base_drawing}-{spool_no}"
                
                cut_pipes.append({
                    "piece_no": piece_no,
                    "pt_no": pt_no,
                    "item_no": item_no,
                    "spool_no": spool_no,
                    "length": length
                })

    # 2. FABRICATION & ERECTION MATERIALS
    for tname in ["FABRICATION MATERIALS", "ERECTION MATERIALS"]:
        if tname not in tables or not tables[tname]["rows"]:
            continue
            
        current_category = ""
        for rtype, rdata in tables[tname]["rows"]:
            if rtype == "category":
                current_category = rdata
                continue
                
            if rtype == "data" and len(rdata) >= 7:
                pl_no = rdata[0].strip()
                desc = rdata[1].strip()
                item_code = clean_item_code(rdata[3].strip())
                alt_code = clean_item_code(rdata[4].strip())
                if (not item_code or "X" in item_code.upper()) and alt_code:
                    item_code = alt_code
                
                qty = rdata[5].strip()
                weight = rdata[6].strip()
                
                # Eğer bu bir boru (PIPE) ise ve CUT PIPE LENGTH tablosunda detayları varsa, detayları satır olarak ekle
                if "PIPE" in current_category and cut_pipes:
                    # piece_no (1) ile Fabrication'daki pl_no (1) eşleşmeli VEYA item_code eşleşmeli
                    pipes_found = [p for p in cut_pipes if (p["piece_no"] == pl_no and pl_no != "") or (p["item_no"] == item_code and item_code != "")]
                    if pipes_found:
                        for p in pipes_found:
                            row_dict = {
                                "id": "",
                                "technical_drawing": tech_drawing,
                                "assembly": p["spool_no"],
                                "assembly_description": "",
                                "assembly_weight": "",
                                "assembly_qty": "1",
                                "sub_assembly": p["pt_no"],
                                "sub_assembly_defination": desc,
                                "item_code": item_code, # Ana borudan (Fabrication tablosundan) miras alınır
                                "qty": p["length"], # Borular için adet yerine uzunluk yazılabilir
                                "unit_weight": weight,
                                "pose_no": "",
                                "giris_yapan": "",
                                "giris_tarihi": ""
                            }
                            rows.append(row_dict)
                        continue # Pipe parçalarını alt parçalar olarak ekledik, ana Pipe satırını atla
                
                # Spool (Assembly) Tahmini (Association Architecture)
                import math
                norm_pl_no = re.sub(r'[^0-9A-Z\-]', '', pl_no)
                auto_spool = ""
                assoc_method = "UNRESOLVED"
                assoc_score = 0.0
                assoc_reason = "No matching piece_no or context found"

                # 1. CUT PIPE LENGTH CROSS-REFERENCE (IF PIPE)
                if "PIPE" in current_category and cut_pipes:
                    pipe_candidates = [p for p in cut_pipes if p.get("spool_no")]
                    exact_pipes = [p for p in pipe_candidates if p["item_no"] == item_code and p["pt_no"].strip("<>") == norm_pl_no]
                    if exact_pipes:
                        auto_spool = exact_pipes[0]["spool_no"]
                        assoc_method = "CUT_PIPE_CROSS_REFERENCE"
                        assoc_score = 1.0
                        assoc_reason = f"Matched pipe PT NO {norm_pl_no} and Item Code {item_code}"

                # 2. PIECE NO BBOX PROXIMITY
                if not auto_spool and norm_pl_no and yolo_spools and yolo_pieces:
                    # Normalize YOLO pieces
                    matching_pieces = [p for p in yolo_pieces if p["text"] == norm_pl_no]
                    
                    if matching_pieces:
                        best_dist = float('inf')
                        second_best_dist = float('inf')
                        best_spool = None
                        
                        for p in matching_pieces:
                            px, py = (p["box"][0]+p["box"][2])/2, (p["box"][1]+p["box"][3])/2
                            for s in yolo_spools:
                                sx, sy = (s["box"][0]+s["box"][2])/2, (s["box"][1]+s["box"][3])/2
                                dist = math.hypot(px - sx, py - sy)
                                if dist < best_dist:
                                    second_best_dist = best_dist
                                    best_dist = dist
                                    best_spool = s["text"]
                                elif dist < second_best_dist:
                                    second_best_dist = dist
                        
                        if best_spool:
                            # Dynamic threshold checking
                            margin_ratio = second_best_dist / max(best_dist, 1)
                            if best_dist < 600 and margin_ratio > 1.2:
                                auto_spool = best_spool
                                assoc_method = "PIECE_TO_SPOOL_PROXIMITY"
                                assoc_score = round(1.0 / (1.0 + best_dist/100), 2)
                                assoc_reason = f"Nearest spool to piece {norm_pl_no} (dist={int(best_dist)}px, margin={round(margin_ratio,1)}x)"
                            else:
                                assoc_reason = f"Distance={int(best_dist)}px or low margin={round(margin_ratio,1)}x -> Unreliable"
                    else:
                        assoc_reason = f"Piece {norm_pl_no} not found by YOLO"

                if auto_spool and base_drawing not in auto_spool:
                    if "SP" in auto_spool:
                        sp_match = re.search(r'(SP\d+)', auto_spool)
                        if sp_match:
                            auto_spool = f"{base_drawing}-{sp_match.group(1)}"
                        else:
                            auto_spool = f"{base_drawing}-{auto_spool}"

                # False positive filter
                desc_upper = desc.upper()
                if item_code == "" and qty in ["", "QTY"] and desc_upper in ["CLIENT", "WEIGHT", "DESCRIPTION", "SUPPORIS", ""]:
                    continue # Garbage row

                # Diğer materyaller (Fittings, Flanges, vb.)
                row_dict = {
                    "id": "",
                    "technical_drawing": tech_drawing,
                    "assembly": auto_spool,
                    "assembly_description": "",
                    "assembly_weight": "",
                    "assembly_qty": "1",
                    "sub_assembly": norm_pl_no if norm_pl_no else "",
                    "sub_assembly_defination": desc,
                    "item_code": item_code,
                    "qty": qty if qty else "1",
                    "unit_weight": weight,
                    "pose_no": "",
                    "giris_yapan": "",
                    "giris_tarihi": "",
                    "_debug": {
                        "pl_no": pl_no,
                        "item_code": item_code,
                        "assoc_method": assoc_method,
                        "assoc_score": assoc_score,
                        "assoc_reason": assoc_reason
                    }
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
