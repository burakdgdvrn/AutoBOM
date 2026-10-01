# -*- coding: utf-8 -*-
"""
OCR motoru başlatma, YOLO model yükleme ve yardımcı fonksiyonlar.
Metin temizleme, doğrulama ve düzeltme işlevleri burada toplanmıştır.
"""

import os
import re
import numpy as np
from io import BytesIO
from PIL import Image
import pymupdf
import easyocr
from difflib import SequenceMatcher

from config import (
    BASE_DIR, ZOOM,
    CATEGORY_KEYWORDS_EXACT, CATEGORY_PREFIXES,
)

# ============================================================================
# YOLO Model (gecikmeli yukleme icin None, ilk kullanımda yuklenir)
# ============================================================================
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
    
    # Strictly extract expected suffix (e.g., SP followed by digits or O-> 0)
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



