# -*- coding: utf-8 -*-
"""
Cutting Table Reader — El yazısı cutting table'lardan SPOOL, CUT NO, CUT LENGTH okuma.

Strateji:
  1. YOLO ile cutting_table bounding box'ını al
  2. HSV mavi filtre ile grid çizgilerini tespit et
  3. Sütun/satır sınırlarını belirle
  4. Her hücreyi TrOCR-handwritten ile oku
  5. Domain-specific post-processing ile sonucu düzelt

Sütun yapısı (sabit):
  HVT CODE | SPOOL | CUT NO | CUT LENGHT | DESCRIPTION | DIA/INCH
"""

import os
import re
import cv2
import numpy as np
from PIL import Image

# ============================================================================
# TrOCR Model (lazy loading — ilk çağrıda indirilir)
# ============================================================================
_trocr_processor = None
_trocr_model = None

def _get_trocr():
    """TrOCR modelini lazy load et."""
    global _trocr_processor, _trocr_model
    if _trocr_processor is None:
        from transformers import TrOCRProcessor, VisionEncoderDecoderModel
        model_name = "microsoft/trocr-base-handwritten"
        print(f"  TrOCR modeli yukleniyor: {model_name}...")
        _trocr_processor = TrOCRProcessor.from_pretrained(model_name)
        _trocr_model = VisionEncoderDecoderModel.from_pretrained(model_name)
        _trocr_model.eval()
        print(f"  TrOCR hazir.")
    return _trocr_processor, _trocr_model


def trocr_read_cell(cell_img):
    """
    Tek bir hücre görüntüsünü TrOCR ile oku.
    cell_img: PIL Image veya numpy array (RGB)
    """
    processor, model = _get_trocr()

    if isinstance(cell_img, np.ndarray):
        cell_img = Image.fromarray(cell_img)

    # Gri tonlamada ise RGB'ye çevir
    if cell_img.mode != "RGB":
        cell_img = cell_img.convert("RGB")

    pixel_values = processor(images=cell_img, return_tensors="pt").pixel_values

    import torch
    with torch.no_grad():
        generated_ids = model.generate(pixel_values, max_new_tokens=20)

    text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
    return text.strip()


# ============================================================================
# BLUE GRID DETECTION (HSV filtreleme ile mavi çizgileri bul)
# ============================================================================

def _detect_blue_grid(img_np):
    """
    Mavi grid çizgilerini HSV ile tespit et.
    Returns: (horizontal_ys, vertical_xs) — sıralı çizgi koordinatları
    """
    hsv = cv2.cvtColor(img_np, cv2.COLOR_RGB2HSV)
    h, w = img_np.shape[:2]

    # Mavi renk aralığı (koyu mavi - açık mavi)
    lower_blue = np.array([90, 30, 50])
    upper_blue = np.array([135, 255, 255])
    blue_mask = cv2.inRange(hsv, lower_blue, upper_blue)

    # Gürültüyü temizle
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    blue_mask = cv2.morphologyEx(blue_mask, cv2.MORPH_CLOSE, kernel)

    # ----- Yatay çizgiler -----
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (w // 3, 1))
    h_lines = cv2.morphologyEx(blue_mask, cv2.MORPH_OPEN, h_kernel, iterations=1)
    h_cnts, _ = cv2.findContours(h_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    h_ys = sorted(set([
        cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2
        for c in h_cnts if cv2.boundingRect(c)[2] > w // 4
    ]))

    # ----- Dikey çizgiler -----
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, h // 3))
    v_lines = cv2.morphologyEx(blue_mask, cv2.MORPH_OPEN, v_kernel, iterations=1)
    v_cnts, _ = cv2.findContours(v_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    v_xs = sorted(set([
        cv2.boundingRect(c)[0] + cv2.boundingRect(c)[2] // 2
        for c in v_cnts if cv2.boundingRect(c)[3] > h // 4
    ]))

    # Çok yakın çizgileri birleştir (merge within 8px)
    h_ys = _merge_close(h_ys, threshold=8)
    v_xs = _merge_close(v_xs, threshold=8)

    return h_ys, v_xs


def _merge_close(vals, threshold=8):
    """Birbirine çok yakın değerleri birleştir (ortalama al)."""
    if not vals:
        return []
    merged = [vals[0]]
    for v in vals[1:]:
        if v - merged[-1] < threshold:
            merged[-1] = (merged[-1] + v) // 2
        else:
            merged.append(v)
    return merged


# ============================================================================
# COLUMN IDENTIFICATION (Hangi sütun hangisi?)
# ============================================================================

def _identify_columns(v_xs, w):
    """
    Dikey çizgi x-koordinatlarından sütun sınırlarını belirle.
    
    Cutting table sütun yapısı (gözlem):
      Col 0: HVT CODE     (~0.00 - 0.20)
      Col 1: SPOOL         (~0.20 - 0.36)
      Col 2: CUT NO        (~0.36 - 0.44)
      Col 3: CUT LENGHT    (~0.44 - 0.56)
      Col 4: DESCRIPTION   (~0.56 - 0.78)
      Col 5: DIA/INCH      (~0.78 - 1.00)
    
    Returns dict with keys: SPOOL, CUT_NO, CUT_LENGTH, each = (x_start, x_end)
    """
    if len(v_xs) < 4:
        # Dikey çizgi bulunamadıysa, sabit oranlarla tahmin et
        return {
            "SPOOL": (int(w * 0.20), int(w * 0.36)),
            "CUT_NO": (int(w * 0.36), int(w * 0.44)),
            "CUT_LENGTH": (int(w * 0.44), int(w * 0.56)),
        }

    # v_xs'i normalize et
    norms = [x / w for x in v_xs]

    # En iyi eşleşmeyi bul: sütunları tanımlayan 5 dikey çizgi bekleniyor
    # 0=sol kenar, 1=HVT/SPOOL sınırı, 2=SPOOL/CUT_NO, 3=CUT_NO/CUT_LENGTH,
    # 4=CUT_LENGTH/DESC, 5=DESC/DIA

    # Strateji: Bilinen normalize pozisyonlara en yakın çizgileri eşle
    target_boundaries = [0.20, 0.36, 0.44, 0.56, 0.78]
    best_matches = {}

    for i, target in enumerate(target_boundaries):
        closest = min(v_xs, key=lambda x: abs(x / w - target))
        if abs(closest / w - target) < 0.10:
            best_matches[i] = closest

    # SPOOL sınırları
    spool_start = best_matches.get(0, int(w * 0.20))
    spool_end = best_matches.get(1, int(w * 0.36))

    # CUT NO sınırları
    cutno_start = spool_end
    cutno_end = best_matches.get(2, int(w * 0.44))

    # CUT LENGTH sınırları
    cutlen_start = cutno_end
    cutlen_end = best_matches.get(3, int(w * 0.56))

    return {
        "SPOOL": (spool_start, spool_end),
        "CUT_NO": (cutno_start, cutno_end),
        "CUT_LENGTH": (cutlen_start, cutlen_end),
    }


# ============================================================================
# POST-PROCESSING RULES (Domain-specific düzeltmeler)
# ============================================================================

def _postprocess_spool(raw):
    """
    SPOOL değerini düzelt.
    Beklenen format: SP01, SP02, ..., SP10 veya " (tekrar işareti)
    """
    raw = raw.strip().upper()
    if not raw or raw in ['"', "''", '""', "DITTO", "II", "11"]:
        return '"'  # Tekrar (ditto) işareti

    # Yaygın OCR hataları
    raw = raw.replace("O", "0").replace("o", "0")
    raw = raw.replace("l", "1").replace("I", "1")
    raw = raw.replace("{", "1").replace("}", "1")
    raw = raw.replace("(", "1").replace(")", "1")

    # SP + sayı pattern'ini bul
    match = re.search(r'[S5][PpF]?\s*[O0]?(\d{1,2})', raw)
    if match:
        num = int(match.group(1))
        if 1 <= num <= 20:
            return f"SP{num:02d}"

    # Sadece sayı varsa
    digits = re.sub(r'[^0-9]', '', raw)
    if digits and 1 <= int(digits) <= 20:
        return f"SP{int(digits):02d}"

    return raw


def _postprocess_cut_no(raw):
    """
    CUT NO değerini düzelt.
    Beklenen format: 1-1, 1-2, ..., 1-10, 2-1, vb.
    """
    raw = raw.strip()
    if not raw:
        return ""

    # Yaygın OCR hataları
    raw = raw.replace("l", "1").replace("I", "1")
    raw = raw.replace("O", "0").replace("o", "0")
    raw = raw.replace("{", "1").replace("(", "1")
    raw = raw.replace("}", "1").replace(")", "1")
    raw = raw.replace("~", "-").replace("_", "-")
    raw = raw.replace(" ", "-")

    # X-Y formatını bul
    match = re.search(r'(\d{1,2})\s*[-–—]\s*(\d{1,2})', raw)
    if match:
        return f"{int(match.group(1))}-{int(match.group(2))}"

    # Sadece tek sayı
    digits = re.sub(r'[^0-9]', '', raw)
    if digits:
        num = int(digits)
        if num <= 20:
            return str(num)

    return raw


def _postprocess_cut_length(raw):
    """
    CUT LENGTH değerini düzelt.
    Beklenen format: 3-5 haneli tamsayı (mm cinsinden boru uzunluğu)
    """
    raw = raw.strip()
    if not raw:
        return ""

    # Yaygın OCR hataları
    raw = raw.replace("O", "0").replace("o", "0")
    raw = raw.replace("l", "1").replace("I", "1")
    raw = raw.replace("{", "1").replace("(", "1")
    raw = raw.replace("}", "1").replace(")", "1")
    raw = raw.replace(",", "").replace(".", "")
    raw = raw.replace(" ", "")

    # Sadece rakamları al
    digits = re.sub(r'[^0-9]', '', raw)
    if digits:
        num = int(digits)
        # Boru uzunluğu: 10mm - 12000mm arası mantıklı
        if 10 <= num <= 12000:
            return str(num)
        # Çok büyük sayılar muhtemelen başında ekstra rakam var
        if num > 12000 and len(digits) > 4:
            # Son 4 haneyi dene
            trimmed = int(digits[-4:])
            if 10 <= trimmed <= 12000:
                return str(trimmed)

    return digits if digits else raw


# ============================================================================
# HEADER ROW DETECTION
# ============================================================================

def _find_header_row(h_ys, img_np):
    """
    İlk 2-3 yatay çizgi arası header bölgesidir.
    Header'ın bittiği y koordinatını döndür.
    """
    if len(h_ys) >= 3:
        # Header genelde ilk 2-3 çizgi arası (~90px)
        # Genelde 3. çizgiden sonra veri başlar
        # Header yüksekliği yaklaşık %20-25
        for i in range(1, len(h_ys)):
            gap = h_ys[i] - h_ys[0]
            if gap > 40:
                return h_ys[i]
        return h_ys[2] if len(h_ys) > 2 else h_ys[1]
    elif len(h_ys) >= 2:
        return h_ys[1]
    else:
        # Tahmini: üstten %25
        return int(img_np.shape[0] * 0.25)


# ============================================================================
# MAIN: Read Cutting Table
# ============================================================================

def read_cutting_table(img_np, yolo_box=None):
    """
    Cutting table görüntüsünden SPOOL, CUT NO, CUT LENGTH verilerini oku.

    Args:
        img_np: numpy array (RGB) — tam PDF görüntüsü veya sadece cutting table crop
        yolo_box: (x1, y1, x2, y2) — YOLO'dan gelen bounding box. None ise img_np zaten crop.

    Returns:
        list of dict: [{"spool": "SP01", "cut_no": "1-3", "cut_length": "1648"}, ...]
    """
    # 1. Crop if YOLO box provided
    if yolo_box is not None:
        x1, y1, x2, y2 = yolo_box
        pad = 5
        cy1 = max(0, y1 - pad)
        cy2 = min(img_np.shape[0], y2 + pad)
        cx1 = max(0, x1 - pad)
        cx2 = min(img_np.shape[1], x2 + pad)
        cropped = img_np[cy1:cy2, cx1:cx2]
    else:
        cropped = img_np

    h, w = cropped.shape[:2]

    # 2. Mavi grid çizgilerini bul
    h_ys, v_xs = _detect_blue_grid(cropped)
    print(f"    Grid: {len(h_ys)} yatay, {len(v_xs)} dikey cizgi")
    print(f"    H: {h_ys}")
    print(f"    V: {v_xs}")

    # 3. Sütun sınırlarını belirle
    col_bounds = _identify_columns(v_xs, w)
    print(f"    Sütun sınırları: {col_bounds}")

    # 4. Header satırını bul
    header_end_y = _find_header_row(h_ys, cropped)
    print(f"    Header bitis: y={header_end_y}")

    # 5. Veri satırlarını belirle
    data_row_ys = [y for y in h_ys if y > header_end_y]

    if len(data_row_ys) < 1:
        # Yatay çizgi bulunamadıysa, piksel bazında satır tahmini yap
        row_height = 35  # ~35px per row typical
        y = header_end_y + 5
        data_row_ys = []
        while y < h - 10:
            data_row_ys.append(y)
            y += row_height

    # Satır başlangıç/bitiş çiftleri oluştur
    row_pairs = []
    for i in range(len(data_row_ys) - 1):
        y_start = data_row_ys[i]
        y_end = data_row_ys[i + 1]
        if y_end - y_start > 10:  # En az 10px yükseklik
            row_pairs.append((y_start, y_end))

    # Son satır - tablo alt sınırına kadar
    if data_row_ys:
        last_y = data_row_ys[-1]
        if h - last_y > 15:
            row_pairs.append((last_y, min(h, last_y + 40)))

    print(f"    {len(row_pairs)} veri satırı tespit edildi")

    # 6. Her satırda SPOOL, CUT NO, CUT LENGTH hücrelerini oku
    results = []
    for row_idx, (ry1, ry2) in enumerate(row_pairs):
        row_data = {}
        is_empty_row = True

        for col_name, (cx1, cx2) in col_bounds.items():
            # Hücreyi kırp
            cell = cropped[ry1 + 2:ry2 - 2, cx1 + 2:cx2 - 2]
            if cell.size == 0:
                row_data[col_name] = ""
                continue

            # Hücre boş mu kontrol et (çok az piksel varsa boş)
            gray_cell = cv2.cvtColor(cell, cv2.COLOR_RGB2GRAY)
            # İnvert: yazı beyaz, arka plan siyah (OTSU)
            _, binary = cv2.threshold(gray_cell, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            
            # Kenarlardan 4 piksel kırp ki artakalan tablo çizgileri hata yapmasın
            if binary.shape[0] > 8 and binary.shape[1] > 8:
                binary = binary[4:-4, 4:-4]
                
            dark_pixel_ratio = np.sum(binary > 0) / max(binary.size, 1)

            if dark_pixel_ratio < 0.01:
                row_data[col_name] = ""
                continue

            is_empty_row = False

            # Hücreyi TrOCR'a gönder (preprocessing)
            cell_for_ocr = _preprocess_cell(cell)
            raw_text = trocr_read_cell(cell_for_ocr)

            # Post-processing
            if col_name == "SPOOL":
                row_data[col_name] = _postprocess_spool(raw_text)
            elif col_name == "CUT_NO":
                row_data[col_name] = _postprocess_cut_no(raw_text)
            elif col_name == "CUT_LENGTH":
                row_data[col_name] = _postprocess_cut_length(raw_text)
            else:
                row_data[col_name] = raw_text

        if is_empty_row:
            continue

        # Ditto (") işareti kontrolü — SPOOL boşsa üst satırdan al
        if row_data.get("SPOOL") == '"' and results:
            row_data["SPOOL"] = results[-1].get("spool", "")

        results.append({
            "spool": row_data.get("SPOOL", ""),
            "cut_no": row_data.get("CUT_NO", ""),
            "cut_length": row_data.get("CUT_LENGTH", ""),
            "_raw": row_data,
        })
        print(f"    Satır {row_idx}: SPOOL={row_data.get('SPOOL', '')}, "
              f"CUT_NO={row_data.get('CUT_NO', '')}, "
              f"CUT_LENGTH={row_data.get('CUT_LENGTH', '')}")

    return results


def _preprocess_cell(cell_img):
    """
    Hücre görüntüsünü OCR için optimize et:
    TrOCR 384x384 kare görüntü bekler. 
    Aspect ratio bozulmaması için resmi kare olacak şekilde (beyaz) padliyoruz.
    """
    h, w = cell_img.shape[:2]
    result = cell_img.copy()

    # Hedef kare boyutu
    target_size = max(h, w)
    
    # Ne kadar padding gerekecek?
    pad_h = target_size - h
    pad_w = target_size - w
    
    top = pad_h // 2
    bottom = pad_h - top
    left = pad_w // 2
    right = pad_w - left
    
    # Ekstra 16px genel padding (kenarlara değmesin diye)
    extra_pad = 16
    top += extra_pad
    bottom += extra_pad
    left += extra_pad
    right += extra_pad
    
    padded = cv2.copyMakeBorder(result, top, bottom, left, right,
                                cv2.BORDER_CONSTANT, value=[255, 255, 255])

    return padded


# ============================================================================
# INTEGRATION: YOLO + Cutting Table Reader
# ============================================================================

def extract_cutting_table_from_pdf(img, yolo_model=None):
    """
    Tam PDF görüntüsünden cutting table'ı YOLO ile bul ve oku.

    Args:
        img: PIL Image — pdf_to_image() çıktısı
        yolo_model: YOLO model instance (None ise lazy load)

    Returns:
        list of dict veya None (cutting table bulunamazsa)
    """
    img_np = np.array(img)

    # YOLO ile cutting table tespit et
    if yolo_model is None:
        from ocr_utils import get_yolo_model
        yolo_model = get_yolo_model()

    if yolo_model is None:
        print("    UYARI: YOLO modeli bulunamadı!")
        return None

    results = yolo_model(img_np, conf=0.30, verbose=False)

    best_ct = None
    best_conf = 0
    for box in results[0].boxes:
        cls_name = yolo_model.names[int(box.cls[0])]
        conf_val = float(box.conf[0])
        if cls_name == "cutting_table" and conf_val > best_conf:
            x1, y1, x2, y2 = [int(v) for v in box.xyxy[0]]
            best_ct = (x1, y1, x2, y2)
            best_conf = conf_val

    if best_ct is None:
        print("    Cutting table bulunamadi!")
        return None

    print(f"    Cutting table: {best_ct}, conf={best_conf:.3f}")

    # Cutting table'ı oku
    return read_cutting_table(img_np, yolo_box=best_ct)
