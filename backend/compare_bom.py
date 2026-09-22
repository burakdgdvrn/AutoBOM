import pandas as pd
import re

def normalize_assembly(val):
    val = str(val).strip().upper()
    if val.lower() == 'nan' or val == 'NONE': return ''
    # Strip _1_1, _1_2 etc for comparison
    val = re.sub(r'_\d+_\d+', '', val)
    return val

def normalize_sub(val):
    val = str(val).strip().upper()
    if val.lower() == 'nan' or val == 'NONE' or not val: return ''
    # Eğer Excel'de "1-5" (PieceNo-PtNo) formatı varsa ve AI sadece "5" ürettiyse eşleştirebilmek için:
    if re.match(r'^\d+-\d+$', val):
        val = val.split('-')[-1]
    return val

def get_base_drawing(val):
    """
    Örn: '10-FG-502433-2001-P1401_1_2' -> '10-FG-502433-2001-P1401'
    """
    val = str(val).strip().upper()
    if pd.isna(val) or val == 'NAN' or val == 'NONE' or not val: return ''
    # Arkadaki revizyon vb ekleri at
    val = val.split('_')[0]
    # Bazen '-SP01' de olabiliyor assembly için
    val = val.split('-SP')[0]
    return val

def normalize_item(val):
    val = str(val).strip().upper()
    if val.lower() == 'nan' or val == 'NONE' or not val: return ''
    # I, 1, l öneklerini at (Örn: I5669 -> 5669, 15669 -> 5669)
    val = re.sub(r'^[1IL]0?', '', val)
    return val

def get_row_key(row, is_excel=False):
    """
    Excel ve JSON verileri için ortak Primary Key oluşturur.
    is_excel=True ise başlıklar farklı olabilir ama genelde aynı.
    """
    assembly = row.get('Assembly') or row.get('SPOOL (ASSEMBLY)') or row.get('assembly') or ''
    sub_assembly = row.get('Sub Assembly') or row.get('SUB ASSEMBLY') or row.get('sub_assembly') or ''
    item_code = row.get('Item Code') or row.get('ITEM CODE') or row.get('item_code') or ''
    
    a_norm = normalize_assembly(assembly)
    s_norm = normalize_sub(sub_assembly)
    i_norm = normalize_item(item_code)
    
    return f"{a_norm}|{s_norm}|{i_norm}"

def get_qty(row):
    qty = row.get('Qty') or row.get('QTY') or row.get('qty')
    if pd.isna(qty): return ''
    val = str(qty).strip()
    if val.endswith('.0'): val = val[:-2]
    return val

def safe_val(val):
    """Pandas NaN değerlerini boş string'e çevirir (JSON hatasını önler)."""
    if pd.isna(val): return ""
    return str(val) if val is not None else ""

def compare_excel_and_json(excel_path, json_data):
    """
    Orijinal Excel ile AI'ın bulduğu JSON verisini hücre bazlı heuristik mantıkla karşılaştırır.
    """
    df_excel = pd.read_excel(excel_path).fillna("")
    excel_rows = df_excel.to_dict('records')
    
    ai_rows = [r for r in json_data if r.get('assembly') or r.get('item_code')]
    
    ai_base_drawings = set()
    for r in ai_rows:
        dwg = r.get('technical_drawing') or r.get('TECHNICAL DRAWING') or r.get('Technical Drawing') or r.get('assembly')
        if dwg:
            b_dwg = get_base_drawing(dwg)
            if b_dwg:
                ai_base_drawings.add(b_dwg)
                
    filtered_excel_rows = []
    for r in excel_rows:
        ex_dwg = r.get('Technical Drawing')
        if ex_dwg:
            b_ex_dwg = get_base_drawing(ex_dwg)
            if b_ex_dwg in ai_base_drawings:
                filtered_excel_rows.append(r)
                
    # Heuristic Matching
    unmatched_ai = list(ai_rows)
    unmatched_ex = list(filtered_excel_rows)
    
    perfect_matches = []
    cell_discrepancies = []
    
    def calculate_score(ai_r, ex_r):
        score = 0
        
        # Normalize fields
        ai_item = normalize_item(ai_r.get('item_code'))
        ex_item = normalize_item(ex_r.get('Item Code'))
        ai_sub = normalize_sub(ai_r.get('sub_assembly'))
        ex_sub = normalize_sub(ex_r.get('Sub Assembly'))
        ai_assm = normalize_assembly(ai_r.get('assembly'))
        ex_assm = normalize_assembly(ex_r.get('Assembly'))
        ai_q = safe_val(get_qty(ai_r))
        ex_q = safe_val(get_qty(ex_r))
        
        ai_dwg = get_base_drawing(ai_r.get('technical_drawing') or ai_r.get('assembly'))
        ex_dwg = get_base_drawing(ex_r.get('Technical Drawing'))
        
        if ai_item and ai_item == ex_item:
            score += 40
        elif ai_item and ex_item and ai_item != ex_item:
            score -= 20
            
        if ai_assm and ai_assm == ex_assm:
            score += 40
        elif ai_assm and ex_assm and ai_assm != ex_assm:
            score -= 20
            
        if ai_sub and ai_sub == ex_sub:
            score += 20
            
        if ai_dwg and ai_dwg == ex_dwg:
            score += 20
            
        if ai_q and ai_q == ex_q:
            score += 10
            
        return score

    # First pass: Find matches
    for ex_r in list(unmatched_ex):
        best_ai = None
        best_score = -1
        
        for ai_r in unmatched_ai:
            score = calculate_score(ai_r, ex_r)
            if score > best_score:
                best_score = score
                best_ai = ai_r
                
        # Threshold for match
        if best_score >= 50:
            unmatched_ex.remove(ex_r)
            unmatched_ai.remove(best_ai)
            
            ai_item = normalize_item(best_ai.get('item_code'))
            ex_item = normalize_item(ex_r.get('Item Code'))
            ai_sub = normalize_sub(best_ai.get('sub_assembly'))
            ex_sub = normalize_sub(ex_r.get('Sub Assembly'))
            ai_assm = normalize_assembly(best_ai.get('assembly'))
            ex_assm = normalize_assembly(ex_r.get('Assembly'))
            ai_q = safe_val(get_qty(best_ai))
            ex_q = safe_val(get_qty(ex_r))
            ai_dwg = get_base_drawing(best_ai.get('technical_drawing') or best_ai.get('assembly'))
            ex_dwg = get_base_drawing(ex_r.get('Technical Drawing'))
            
            match_data = {
                "technical_drawing": ex_dwg or ai_dwg,
                "ai_assembly": ai_assm,
                "ex_assembly": ex_assm,
                "ai_sub_assembly": ai_sub,
                "ex_sub_assembly": ex_sub,
                "ai_item_code": ai_item,
                "ex_item_code": ex_item,
                "ai_qty": ai_q,
                "ex_qty": ex_q
            }
            
            # If everything matches perfectly
            if ai_item == ex_item and ai_sub == ex_sub and ai_assm == ex_assm and ai_q == ex_q:
                perfect_matches.append(match_data)
            else:
                cell_discrepancies.append(match_data)

    extra_by_ai = []
    for ai_r in unmatched_ai:
        extra_by_ai.append({
            "technical_drawing": get_base_drawing(ai_r.get('technical_drawing') or ai_r.get('assembly')),
            "ai_assembly": normalize_assembly(ai_r.get('assembly')),
            "ai_sub_assembly": normalize_sub(ai_r.get('sub_assembly')),
            "ai_item_code": normalize_item(ai_r.get('item_code')),
            "ai_qty": safe_val(get_qty(ai_r)),
            "desc": safe_val(ai_r.get('sub_assembly_defination') or ai_r.get('assembly_description'))
        })

    missed_by_ai = []
    for ex_r in unmatched_ex:
        missed_by_ai.append({
            "technical_drawing": get_base_drawing(ex_r.get('Technical Drawing')),
            "ex_assembly": normalize_assembly(ex_r.get('Assembly')),
            "ex_sub_assembly": normalize_sub(ex_r.get('Sub Assembly')),
            "ex_item_code": normalize_item(ex_r.get('Item Code')),
            "ex_qty": safe_val(get_qty(ex_r)),
            "desc": safe_val(ex_r.get('Sub Assembly Defination') or ex_r.get('Assembly Description'))
        })
        
    return {
        "stats": {
            "total_ai": len(ai_rows),
            "total_excel": len(filtered_excel_rows),
            "perfect_matches": len(perfect_matches),
            "cell_discrepancies": len(cell_discrepancies),
            "extra_by_ai": len(extra_by_ai),
            "missed_by_ai": len(missed_by_ai)
        },
        "perfect_matches": perfect_matches,
        "cell_discrepancies": cell_discrepancies,
        "extra_by_ai": extra_by_ai,
        "missed_by_ai": missed_by_ai,
        "raw_filtered_excel": filtered_excel_rows,
        "raw_ai_data": ai_rows
    }
