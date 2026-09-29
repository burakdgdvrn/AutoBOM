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

Modül yapısı:
  - config.py         : Sabitler ve yapılandırma
  - ocr_utils.py      : OCR motoru, YOLO, yardımcı fonksiyonlar
  - table_parser.py   : Tablo bölgesi ayrıştırma
  - excel_writer.py   : Excel yazma
  - pdf_to_excel.py   : Ana işlem orchestration (bu dosya)
"""

import os
import sys
import re
import numpy as np

from config import PDF_FOLDER, OUTPUT_FILE
from ocr_utils import (
    pdf_to_image, detect_yolo_labels, extract_drawing_names,
    canonicalize_spool, validate_piece_no,
)
from table_parser import parse_table_region
from excel_writer import write_bom_excel


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
        "image": img,
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

    # 3. Collect Fabrication Requirements for Spatial Engine
    target_pls_qty_dict = {}
    fab_rows = []
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
                norm_pl = re.sub(r'[^0-9A-Z\-]', '', pl_no)
                
                desc = str(rdata[1]).strip()
                raw_item = str(rdata[3]).strip()
                qty_str = str(rdata[5]).strip()
                weight = str(rdata[6]).strip()
                
                if "SUPPORT" in current_category.upper() or raw_item.upper().startswith("X"):
                    continue
                    
                try:
                    total_qty = int(float(qty_str))
                except:
                    total_qty = 1
                    
                if norm_pl:
                    target_pls_qty_dict[norm_pl] = target_pls_qty_dict.get(norm_pl, 0) + total_qty
                    
                fab_rows.append({
                    "pl_no": pl_no,
                    "norm_pl": norm_pl,
                    "desc": desc,
                    "raw_item": raw_item,
                    "qty": total_qty,
                    "weight": weight,
                    "category": current_category
                })

    # 4. Run Advanced Spatial Engine
    import traceback
    try:
        from spatial_engine import run_spatial_engine
        img = result["image"]
        img_np = np.array(img)
        w, h = img.size
        roi = [int(w * 0.05), int(h * 0.10), int(w * 0.60), int(h * 0.80)]
        spatial_results = run_spatial_engine(img_np, roi, yolo_spools, target_pls_qty_dict)
    except Exception as e:
        print(f"Spatial Engine Error: {e}")
        traceback.print_exc()
        spatial_results = []
        
    spatial_lookup = {}
    for sr in spatial_results:
        pl = sr["pl_no"]
        if pl not in spatial_lookup: spatial_lookup[pl] = []
        spatial_lookup[pl].append(sr)

    # 5. Association & Row Generation
    rows = []
    
    def emit_row(assembly_spool, emit_qty, trace_dict, fab_data):
        row_dict = {
            "id": "",
            "technical_drawing": tech_drawing,
            "assembly": assembly_spool,
            "assembly_description": "",
            "assembly_weight": "",
            "assembly_qty": "1",
            "sub_assembly": re.sub(r'[^0-9A-Z\-]', '', str(fab_data["pl_no"])),
            "sub_assembly_defination": fab_data["desc"],
            "item_code": trace_dict["item_code_raw"],
            "qty": str(emit_qty) + ".0",
            "unit_weight": fab_data["weight"],
            "pose_no": "",
            "giris_yapan": "",
            "giris_tarihi": "",
            "_debug": trace_dict
        }
        rows.append(row_dict)

    for f in fab_rows:
        norm_data = normalize_item_code(f["raw_item"])
        lookup_code = norm_data["candidate"]
        
        trace = {
            "piece_no": f["pl_no"],
            "item_code_raw": norm_data["raw"],
            "item_code_normalized": lookup_code,
            "normalization_status": norm_data["status"],
            "category": f["category"],
            "spool": None,
            "method": None,
            "score": 0.0,
            "status": "UNRESOLVED",
            "reason": ""
        }
        
        if "PIPE" in f["category"].upper():
            candidates = cut_dict.get(lookup_code, [])
            if len(candidates) == 1:
                trace["spool"] = candidates[0]["spool_no"]
                trace["method"] = "DETERMINISTIC_CROSS_REF"
                trace["score"] = 1.0
                trace["status"] = "STRONG_DETERMINISTIC"
                trace["reason"] = f"Exact item match '{lookup_code}' in CUT PIPE with unique spool candidate"
                
                # Emit sub-piece (pipe) directly
                row_dict = {
                    "id": "",
                    "technical_drawing": tech_drawing,
                    "assembly": candidates[0]["spool_no"],
                    "assembly_description": "",
                    "assembly_weight": "",
                    "assembly_qty": "1",
                    "sub_assembly": candidates[0]["pt_no"],
                    "sub_assembly_defination": f["desc"],
                    "item_code": norm_data["raw"],
                    "qty": "1.0",
                    "unit_weight": f["weight"],
                    "pose_no": "",
                    "giris_yapan": "",
                    "giris_tarihi": "",
                    "_debug": trace
                }
                rows.append(row_dict)
                continue
                
            elif len(candidates) > 1:
                trace["spool"] = "BİLİNMEYEN (MULTIPLE CUT PIPE)"
                trace["method"] = "DETERMINISTIC_CROSS_REF"
                trace["score"] = 0.5
                trace["status"] = "AMBIGUOUS"
                trace["reason"] = f"Multiple CUT PIPE candidates ({len(candidates)}) found for item '{lookup_code}'"
                emit_row(trace["spool"], f["qty"], trace, f)
            else:
                trace["spool"] = "BİLİNMEYEN (PIPE NO CUT REF)"
                trace["method"] = "DETERMINISTIC_FAIL"
                trace["status"] = "UNRESOLVED"
                trace["reason"] = f"No CUT PIPE match for '{lookup_code}'"
                emit_row(trace["spool"], f["qty"], trace, f)
                
        elif any(c in f["category"].upper() for c in ["FITTNGS", "FITTINGS", "FLANGES"]):
            assigned_spools = spatial_lookup.get(f["norm_pl"], [])
            if not assigned_spools:
                trace["spool"] = "BİLİNMEYEN (SPATIAL FAIL)"
                trace["method"] = "SPATIAL_WEAK"
                trace["status"] = "UNRESOLVED"
                trace["reason"] = "No spatial candidates found"
                emit_row(trace["spool"], f["qty"], trace, f)
            else:
                for s_res in assigned_spools:
                    trace_copy = trace.copy()
                    trace_copy["spool"] = s_res["spool"]
                    trace_copy["method"] = s_res["method"]
                    trace_copy["score"] = s_res["score"]
                    trace_copy["status"] = s_res["status"]
                    trace_copy["reason"] = s_res["reason"]
                    emit_row(s_res["spool"], 1, trace_copy, f)
        else:
            trace["spool"] = "BİLİNMEYEN (UNKNOWN CAT)"
            trace["status"] = "UNRESOLVED"
            trace["reason"] = f"Unknown category '{f['category']}'"
            emit_row(trace["spool"], f["qty"], trace, f)

    return rows

def main():
    print("=" * 60)
    print("  PDF'den Tablo Cikarma Sistemi v3 (BOM & Spool Associations)")
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
    
    all_rows = []
    for fn in pdf_files:
        try:
            pdf_path = os.path.join(PDF_FOLDER, fn)
            pdf_name = fn.replace('.pdf', '')
            
            # Use the newly integrated BOM Extractor which includes the Spatial Engine
            rows = extract_bom_data_to_json(pdf_path, pdf_name)
            all_rows.extend(rows)
            
        except Exception as e:
            print(f"\n  HATA: {fn}: {e}")
            import traceback
            traceback.print_exc()
    
    if all_rows:
        write_bom_excel(all_rows)
        print(f"\n{'='*60}")
        print(f"  TAMAMLANDI! {len(pdf_files)} PDF islendi -> {OUTPUT_FILE}")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
