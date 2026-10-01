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
    
    from master_indexer import get_master_candidates, find_master_by_description

    def normalize_item_code(raw_code, description=""):
        raw_code = raw_code.strip()
        candidate = raw_code
        status = "RAW"
        rule = None
        evidence = []
        master_match = False
        master_candidates = []
        
        if not raw_code:
            desc_cands = find_master_by_description(description)
            if desc_cands:
                evidence.append(f"Dry-run description match found {len(desc_cands)} candidates.")
                status = "EMPTY_DESC_MATCH"
            else:
                evidence.append("Empty item code, no description match found.")
                status = "EMPTY_NO_MATCH"
            return {
                "raw": raw_code,
                "candidate": "",
                "rule": "EMPTY_DESC_DRY_RUN",
                "evidence": evidence,
                "status": status,
                "master_match": False,
                "master_candidates": desc_cands
            }
            
        cands = get_master_candidates(raw_code)
        if cands:
            candidate = raw_code
            master_candidates = cands
            master_match = True
            rule = "RAW_EXACT_MATCH"
        elif raw_code.startswith("1") and len(raw_code) > 1:
            hyp_candidate = raw_code[1:]
            hyp_cands = get_master_candidates(hyp_candidate)
            if hyp_cands:
                candidate = hyp_candidate
                master_candidates = hyp_cands
                master_match = True
                rule = "FABRICATION_PREFIX_HYPOTHESIS"
                evidence.append(f"Starts with 1. Candidate '{hyp_candidate}' matched.")
            else:
                candidate = raw_code
                rule = "FABRICATION_PREFIX_HYPOTHESIS_FAILED"
                evidence.append("Starts with 1.")
                evidence.append(f"Hypothesis '{hyp_candidate}' NOT found. Reverting to raw.")
                
        if master_match:
            if len(master_candidates) > 1:
                status = "MASTER_AMBIGUOUS"
            else:
                status = "VERIFIED_CANDIDATE"
        else:
            status = "RAW_FALLBACK" if rule else "RAW"
            
        return {
            "raw": raw_code,
            "candidate": candidate,
            "rule": rule,
            "evidence": evidence,
            "status": status,
            "master_match": master_match,
            "master_candidates": master_candidates
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

    # Collect unique spools to detect Single-Spool fallback
    valid_spools = set()
    for ys in yolo_spools:
        if ys.get("text") and "BİLİNMEYEN" not in ys["text"].upper():
            valid_spools.add(ys["text"])
    for c_pipe in cut_pipes:
        if c_pipe.get("spool_no") and "BİLİNMEYEN" not in c_pipe["spool_no"].upper():
            valid_spools.add(c_pipe["spool_no"])
    single_spool = list(valid_spools)[0] if len(valid_spools) == 1 else None

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
            "item_code": trace_dict.get("item_code_normalized") if trace_dict.get("item_code_normalized") else trace_dict.get("item_code_raw", ""),
            "qty": str(emit_qty) + ".0",
            "unit_weight": fab_data["weight"],
            "pose_no": "",
            "giris_yapan": "",
            "giris_tarihi": "",
            "_debug": trace_dict
        }
        rows.append(row_dict)

    for f in fab_rows:
        norm_data = normalize_item_code(f["raw_item"], f["desc"])
        lookup_code = norm_data["candidate"]
        
        trace = {
            "piece_no": f["pl_no"],
            "item_code_raw": norm_data["raw"],
            "item_code_normalized": lookup_code,
            "normalization_status": norm_data["status"],
            "normalization_rule": norm_data["rule"],
            "master_match": norm_data["master_match"],
            "master_candidates": len(norm_data["master_candidates"]),
            "category": f["category"],
            "spool": None,
            "method": None,
            "score": 0.0,
            "status": "UNRESOLVED",
            "reason": ""
        }
        
        # ---------------- Phase 4: Candidate Resolution ----------------
        spatial_cand = ""
        cut_pipe_cand = ""
        if "PIPE" in f["category"].upper():
            candidates_cp = cut_dict.get(lookup_code, [])
            if len(candidates_cp) >= 1:
                cut_pipe_cand = candidates_cp[0]["spool_no"]
        elif any(c in f["category"].upper() for c in ["FITTINGS", "FITTNGS", "FLANGE", "GASKET", "BOLT", "VALVE", "IN-LINE", "SPECIAL", "INSTRUMENT"]):
            assigned_spools = spatial_lookup.get(f["norm_pl"], [])
            valid_spatial = [s for s in assigned_spools if "SPATIAL FAIL" not in s["spool"]]
            if valid_spatial:
                spatial_cand = valid_spatial[0]["spool"]
            elif single_spool:
                spatial_cand = single_spool
                
        ctx = {
            "raw_item_code": norm_data["raw"],
            "normalized_item_code": lookup_code,
            "technical_drawing": tech_drawing,
            "description": f["desc"],
            "category": f["category"],
            "spatial_spool": spatial_cand,
            "cut_pipe_spool": cut_pipe_cand
        }
        
        from candidate_resolution import resolve_candidates, dry_run_description
        if lookup_code:
            res_trace = resolve_candidates(ctx)
            trace["phase4_resolution"] = res_trace
            
            # Auto-bind if High Confidence
            if res_trace["status"] == "HIGH_CONFIDENCE" and res_trace.get("top1"):
                trace["spool"] = res_trace["top1"]["candidate"]["assembly"]
                trace["sub_assembly_override"] = res_trace["top1"]["candidate"]["sub_assembly"]
                trace["status"] = "HIGH_CONFIDENCE_RESOLVED"
                trace["reason"] = f"Phase 4 resolution matched with margin {res_trace['margin']}"
        else:
            res_trace = dry_run_description(f["desc"])
            trace["phase4_resolution"] = res_trace
        # ---------------------------------------------------------------
        
        if trace.get("status") == "HIGH_CONFIDENCE_RESOLVED":
            f_copy = f.copy()
            f_copy["pl_no"] = trace.get("sub_assembly_override", f["pl_no"])
            emit_row(trace["spool"], f["qty"], trace, f_copy)
            continue

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
                    "item_code": norm_data["candidate"] if norm_data["candidate"] else norm_data["raw"],
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
                
        elif any(c in f["category"].upper() for c in ["FITTINGS", "FITTNGS", "FLANGE", "GASKET", "BOLT", "VALVE", "IN-LINE", "SPECIAL", "INSTRUMENT"]):
            assigned_spools = spatial_lookup.get(f["norm_pl"], [])
            valid_spatial = [s for s in assigned_spools if "SPATIAL FAIL" not in s["spool"]]
            
            if not valid_spatial and single_spool:
                trace["spool"] = single_spool
                trace["method"] = "SINGLE_SPOOL_FALLBACK"
                trace["status"] = "MATCHED"
                trace["reason"] = "Only 1 spool in drawing, bypassing spatial engine fail."
                emit_row(trace["spool"], f["qty"], trace, f)
            elif not assigned_spools:
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

    # 6. Aggregation
    aggregated = {}
    for r in rows:
        # Key based on Assembly, Sub Assembly and Item Code to aggregate Qtys
        key = (r["assembly"], r["sub_assembly"], r["item_code"])
        if key not in aggregated:
            aggregated[key] = r
        else:
            try:
                q1 = float(aggregated[key]["qty"])
                q2 = float(r["qty"])
                aggregated[key]["qty"] = str(q1 + q2)
                
                # Combine debug reasons
                if r["_debug"].get("reason") not in aggregated[key]["_debug"].get("reason", ""):
                    aggregated[key]["_debug"]["reason"] = str(aggregated[key]["_debug"].get("reason", "")) + " | " + str(r["_debug"].get("reason", ""))
            except:
                pass

    return list(aggregated.values())

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
