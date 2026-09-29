import cv2
import numpy as np
import easyocr
import math

reader = None

def get_reader():
    global reader
    if reader is None:
        reader = easyocr.Reader(['en'], gpu=True, verbose=False)
    return reader

def calculate_distance(box1, box2):
    c1 = ((box1[0] + box1[2])/2, (box1[1] + box1[3])/2)
    c2 = ((box2[0] + box2[2])/2, (box2[1] + box2[3])/2)
    return math.hypot(c1[0]-c2[0], c1[1]-c2[1])

def detect_markers(img_np):
    """Detect circles and boxes in the drawing region."""
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blur, 50, 150)
    contours, _ = cv2.findContours(edges, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    shapes = []
    for c in contours:
        area = cv2.contourArea(c)
        if 150 < area < 5000:
            x, y, w, h = cv2.boundingRect(c)
            aspect = w / float(h)
            
            # Circle check
            (cx, cy), radius = cv2.minEnclosingCircle(c)
            circle_area = np.pi * (radius ** 2)
            if area / circle_area > 0.75 and 0.8 <= aspect <= 1.2:
                shapes.append({"type": "CIRCLE", "bbox": [x, y, x+w, y+h]})
                continue
                
            # Box check
            rect_area = w * h
            if area / rect_area > 0.8 and 0.5 <= aspect <= 2.0:
                shapes.append({"type": "BOX", "bbox": [x, y, x+w, y+h]})
    
    return shapes

def get_marker_for_token(t_bbox, t_text, shapes, all_tokens):
    """Finds what marker surrounds/is adjacent to the token."""
    if "<" in t_text or ">" in t_text:
        return {"type": "CHEVRON", "score": 1.0, "bbox": t_bbox}
        
    cx = (t_bbox[0] + t_bbox[2])/2
    cy = (t_bbox[1] + t_bbox[3])/2
    for other in all_tokens:
        if other["text"] in ["<", ">"]:
            ocx = (other["bbox"][0] + other["bbox"][2])/2
            ocy = (other["bbox"][1] + other["bbox"][3])/2
            if math.hypot(cx - ocx, cy - ocy) < 50:
                return {"type": "CHEVRON", "score": 1.0, "bbox": other["bbox"]}
                
    for s in shapes:
        sx1, sy1, sx2, sy2 = s["bbox"]
        if sx1 < cx < sx2 and sy1 < cy < sy2:
            return {"type": s["type"], "score": 1.0, "bbox": s["bbox"]}
            
    return {"type": "NONE", "score": 0.0, "bbox": None}

def run_spatial_engine(img_np, roi, yolo_spools, target_pls_qty_dict):
    """
    img_np: Full page numpy image
    roi: [x1, y1, x2, y2] indicating the drawing area
    yolo_spools: list of spool predictions [{"text": "SP01", "box": [...]}]
    target_pls_qty_dict: {"3": 1, "4": 2} -> dict mapping normalized Piece No to its QTY
    """
    x1_roi, y1_roi, x2_roi, y2_roi = roi
    drawing_img = img_np[y1_roi:y2_roi, x1_roi:x2_roi]
    
    r = get_reader()
    ocr_raw = r.readtext(drawing_img, detail=1, paragraph=False)
    
    all_tokens = []
    numeric_cands = []
    
    for bbox, text, conf in ocr_raw:
        if conf < 0.15: continue
        text = text.strip()
        x1 = min([p[0] for p in bbox]) + x1_roi
        x2 = max([p[0] for p in bbox]) + x1_roi
        y1 = min([p[1] for p in bbox]) + y1_roi
        y2 = max([p[1] for p in bbox]) + y1_roi
        full_bbox = [x1, y1, x2, y2]
        
        token = {"text": text, "bbox": full_bbox, "conf": conf}
        all_tokens.append(token)
        
        cleaned = "".join(c for c in text if c.isdigit())
        if cleaned in target_pls_qty_dict:
            token["cleaned"] = cleaned
            numeric_cands.append(token)
            
    raw_shapes = detect_markers(drawing_img)
    shapes = []
    for s in raw_shapes:
        bx1, by1, bx2, by2 = s["bbox"]
        shapes.append({"type": s["type"], "bbox": [bx1+x1_roi, by1+y1_roi, bx2+x1_roi, by2+y1_roi]})

    results = []
    
    for target, qty in target_pls_qty_dict.items():
        cands_for_target = [c for c in numeric_cands if c["cleaned"] == target]
        scored_cands = []
        
        for cand in cands_for_target:
            marker = get_marker_for_token(cand["bbox"], cand["text"], shapes, all_tokens)
            
            # Distance to nearest spool
            min_spool_dist = float('inf')
            nearest_spool = None
            for ys in yolo_spools:
                d = calculate_distance(cand["bbox"], ys["box"])
                if d < min_spool_dist:
                    min_spool_dist = d
                    nearest_spool = ys["text"]
                    
            spool_score = 0.5 * (1000 / (min_spool_dist + 1000)) if min_spool_dist != float('inf') else 0
            
            # Dimension penalty
            dim_penalty = 0
            cx = (cand["bbox"][0] + cand["bbox"][2])/2
            cy = (cand["bbox"][1] + cand["bbox"][3])/2
            for other in all_tokens:
                if other != cand:
                    ocx = (other["bbox"][0] + other["bbox"][2])/2
                    ocy = (other["bbox"][1] + other["bbox"][3])/2
                    if math.hypot(cx - ocx, cy - ocy) < 100:
                        if len(other["text"]) >= 3 and any(c.isdigit() for c in other["text"]):
                            dim_penalty -= 0.5
                            
            total_score = marker["score"] + spool_score + (cand["conf"] * 0.2) + dim_penalty
            
            scored_cands.append({
                "bbox": cand["bbox"],
                "text": cand["text"],
                "marker_type": marker["type"],
                "marker_score": marker["score"],
                "total_score": total_score,
                "spool": nearest_spool
            })
            
        scored_cands.sort(key=lambda x: x["total_score"], reverse=True)
        
        # Valid candidates must have a positive total score and ideally a marker
        valid_cands = [c for c in scored_cands if c["marker_type"] != "NONE" and c["total_score"] > 0]
        
        assigned = 0
        for cand in valid_cands:
            if assigned >= qty:
                break
            
            results.append({
                "pl_no": target,
                "spool": cand["spool"] if cand["spool"] else "BİLİNMEYEN (NO SPOOL NEARBY)",
                "method": f"SPATIAL_ADV_{cand['marker_type']}",
                "score": cand["total_score"],
                "status": "SPATIAL_HIGH_CONFIDENCE",
                "reason": f"Localized with {cand['marker_type']} marker, score {cand['total_score']:.2f}"
            })
            assigned += 1
            
        # Pad with UNRESOLVED
        while assigned < qty:
            results.append({
                "pl_no": target,
                "spool": "BİLİNMEYEN (SPATIAL FAIL)",
                "method": "SPATIAL_WEAK",
                "score": 0.0,
                "status": "UNRESOLVED",
                "reason": f"Insufficient valid markers found. Needed {qty}, found {assigned}."
            })
            assigned += 1
            
    return results
