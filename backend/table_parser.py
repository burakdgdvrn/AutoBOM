# -*- coding: utf-8 -*-
"""
Tablo bölgesi ayrıştırma motoru.
OCR sonuçlarını satırlara ve sütunlara atayarak yapısal tablo verisi üretir.
"""

import cv2
import numpy as np

from config import REGIONS, TABLE_CONFIGS
from ocr_utils import (
    reader, fix_ocr_text,
    is_category_text, is_table_title, is_column_header,
    is_garbled_header, is_noise_row, crop_region,
)


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
