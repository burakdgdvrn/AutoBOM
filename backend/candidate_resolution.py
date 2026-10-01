import re
import difflib
from master_indexer import get_master_candidates

RESOLUTION_CONFIG = {
    "MIN_SCORE_THRESHOLD": 1.0,
    "MARGIN_THRESHOLD": 0.5
}

def normalize_description(desc):
    if not desc: return ""
    desc = str(desc).upper()
    desc = re.sub(r'[^\w\s]', ' ', desc)
    desc = re.sub(r'\s+', ' ', desc).strip()
    return desc

def get_base_drawing(dwg):
    dwg = str(dwg).upper()
    dwg = dwg.split('_')[0]
    dwg = dwg.split('-SP')[0]
    return dwg

def resolve_candidates(context):
    """
    context: {
        "raw_item_code": str,
        "normalized_item_code": str,
        "technical_drawing": str,
        "description": str,
        "category": str,
        "spatial_spool": str, # optional
        "cut_pipe_spool": str # optional
    }
    """
    item_code = context.get("normalized_item_code")
    if not item_code:
        return {"status": "UNRESOLVED", "reason": "No item code"}
        
    candidates = get_master_candidates(item_code)
    candidate_count_before = len(candidates)
    
    if not candidates:
        return {"status": "UNRESOLVED", "reason": "No master candidates"}
        
    # Progressive Filtering
    # 1. Technical Drawing Context
    pdf_dwg = get_base_drawing(context.get("technical_drawing", ""))
    filtered_by_dwg = []
    if pdf_dwg:
        for c in candidates:
            c_dwg = get_base_drawing(c.get("assembly", ""))
            # fallback to sub_assembly if assembly is empty
            if not c_dwg:
                c_dwg = get_base_drawing(c.get("sub_assembly", ""))
            
            # Substring match if exact match fails, but let's try exact first
            if pdf_dwg == c_dwg or (c_dwg and pdf_dwg.replace("-", "") == c_dwg.replace("-", "")):
                filtered_by_dwg.append(c)
                
        # Substring fallback
        if not filtered_by_dwg:
            for c in candidates:
                c_dwg = get_base_drawing(c.get("assembly", ""))
                if c_dwg and (pdf_dwg in c_dwg or c_dwg in pdf_dwg):
                    filtered_by_dwg.append(c)
                    
    pool = filtered_by_dwg if filtered_by_dwg else candidates
    count_after_dwg = len(pool)
    
    # 2. Category Context
    cat_str = str(context.get("category", "")).upper()
    filtered_by_cat = []
    if cat_str:
        for c in pool:
            c_cat = str(c.get("category", "")).upper()
            if c_cat and (cat_str in c_cat or c_cat in cat_str):
                filtered_by_cat.append(c)
                
    pool = filtered_by_cat if filtered_by_cat else pool
    count_after_cat = len(pool)
    
    # 3. Spatial / Cut Pipe Context
    spatial_spool = context.get("spatial_spool", "")
    cut_pipe_spool = context.get("cut_pipe_spool", "")
    
    filtered_by_spatial = []
    if spatial_spool or cut_pipe_spool:
        for c in pool:
            c_assembly = str(c.get("assembly", "")).upper()
            if spatial_spool and (spatial_spool in c_assembly or c_assembly in spatial_spool):
                filtered_by_spatial.append(c)
                continue
            if cut_pipe_spool and (cut_pipe_spool in c_assembly or c_assembly in cut_pipe_spool):
                filtered_by_spatial.append(c)
                
    pool = filtered_by_spatial if filtered_by_spatial else pool
    candidate_count_after = len(pool)
    
    # Scoring Remaining Pool
    scored_pool = []
    norm_pdf_desc = normalize_description(context.get("description", ""))
    
    for c in pool:
        score = 0.0
        evidence = {}
        
        # Description
        c_desc = normalize_description(c.get("description", ""))
        desc_sim = difflib.SequenceMatcher(None, norm_pdf_desc, c_desc).ratio()
        if desc_sim > 0.5:
            score += desc_sim
            evidence["description"] = round(desc_sim, 2)
            
        # Category (if matched)
        c_cat = str(c.get("category", "")).upper()
        if cat_str and c_cat and (cat_str in c_cat or c_cat in cat_str):
            score += 0.5
            evidence["category"] = 0.5
            
        # Context Spatial
        c_assembly = str(c.get("assembly", "")).upper()
        if spatial_spool and (spatial_spool in c_assembly or c_assembly in spatial_spool):
            score += 1.5
            evidence["spatial"] = 1.5
            
        if cut_pipe_spool and (cut_pipe_spool in c_assembly or c_assembly in cut_pipe_spool):
            score += 1.5
            evidence["cut_pipe"] = 1.5
            
        scored_pool.append({
            "candidate": c,
            "total_score": round(score, 2),
            "evidence": evidence
        })
        
    scored_pool.sort(key=lambda x: x["total_score"], reverse=True)
    
    top1 = scored_pool[0] if len(scored_pool) > 0 else None
    top2 = scored_pool[1] if len(scored_pool) > 1 else None
    
    status = "UNRESOLVED"
    margin = 0.0
    
    if top1:
        if top1["total_score"] >= RESOLUTION_CONFIG["MIN_SCORE_THRESHOLD"]:
            if top2:
                margin = round(top1["total_score"] - top2["total_score"], 2)
                if margin >= RESOLUTION_CONFIG["MARGIN_THRESHOLD"]:
                    status = "HIGH_CONFIDENCE"
                else:
                    status = "AMBIGUOUS"
            else:
                margin = top1["total_score"]
                status = "HIGH_CONFIDENCE"
        else:
            if len(pool) > 1:
                status = "AMBIGUOUS"
                if top2:
                    margin = round(top1["total_score"] - top2["total_score"], 2)
            else:
                # Top score is below threshold but it's the only one
                status = "HIGH_CONFIDENCE"
                margin = top1["total_score"]
                
    return {
        "status": status,
        "candidate_count_before": candidate_count_before,
        "candidate_count_after": candidate_count_after,
        "filter_trace": {
            "after_dwg": count_after_dwg,
            "after_cat": count_after_cat,
            "after_spatial": candidate_count_after
        },
        "top1": top1,
        "top2": top2,
        "margin": margin
    }

def dry_run_description(desc):
    from master_indexer import load_master_index
    _, d_idx = load_master_index()
    norm_desc = normalize_description(desc)
    
    best_score = 0
    best_cand = None
    second_best_score = 0
    second_best_cand = None
    
    for m_desc_raw, records in d_idx.items():
        m_desc = normalize_description(m_desc_raw)
        sim = difflib.SequenceMatcher(None, norm_desc, m_desc).ratio()
        
        if sim > best_score:
            second_best_score = best_score
            second_best_cand = best_cand
            best_score = sim
            best_cand = records[0] if records else None
        elif sim > second_best_score:
            second_best_score = sim
            second_best_cand = records[0] if records else None
            
    margin = round(best_score - second_best_score, 2)
    status = "HIGH_CONFIDENCE" if best_score > 0.8 and margin > 0.1 else "AMBIGUOUS"
    if best_score < 0.5:
        status = "UNRESOLVED"
        
    return {
        "status": status,
        "best_score": round(best_score, 2),
        "best_cand": best_cand,
        "margin": margin
    }
