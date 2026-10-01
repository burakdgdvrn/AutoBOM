import pandas as pd
import re

MASTER_EXCEL_PATH = r"C:\Users\burak\Desktop\PDF okuma\data\master_excel\E-26-RNS-01_BOM_Listesi (1).xlsx"

_MASTER_INDEX = None
_MASTER_DESC_INDEX = None

def load_master_index():
    global _MASTER_INDEX, _MASTER_DESC_INDEX
    if _MASTER_INDEX is not None:
        return _MASTER_INDEX, _MASTER_DESC_INDEX
    
    _MASTER_INDEX = {}
    _MASTER_DESC_INDEX = {}
    
    try:
        df = pd.read_excel(MASTER_EXCEL_PATH)
        df = df.fillna("")
        for _, row in df.iterrows():
            item_code = str(row.get('Item Code') or row.get('ITEM CODE') or row.get('item_code', '')).strip()
            if item_code.endswith('.0'): item_code = item_code[:-2]
            desc = str(row.get('Description') or row.get('DESCRIPTION') or row.get('description', '')).strip()
            cat = str(row.get('Category') or row.get('CATEGORY') or row.get('category', '')).strip()
            assembly = str(row.get('Assembly') or row.get('SPOOL (ASSEMBLY)') or row.get('assembly', '')).strip()
            sub_assembly = str(row.get('Sub Assembly') or row.get('SUB ASSEMBLY') or row.get('sub_assembly', '')).strip()
            
            record = {
                "item_code": item_code,
                "description": desc,
                "category": cat,
                "assembly": assembly,
                "sub_assembly": sub_assembly
            }
            
            # Item Code Index (norm to upper string)
            if item_code:
                norm_ic = str(item_code).upper()
                if norm_ic not in _MASTER_INDEX:
                    _MASTER_INDEX[norm_ic] = []
                _MASTER_INDEX[norm_ic].append(record)
            
            # Description Index (for empty item codes)
            if desc:
                norm_desc = re.sub(r'\s+', ' ', desc.upper()).strip()
                if norm_desc not in _MASTER_DESC_INDEX:
                    _MASTER_DESC_INDEX[norm_desc] = []
                _MASTER_DESC_INDEX[norm_desc].append(record)
                
    except Exception as e:
        print(f"Master indexer load error: {e}")
        
    return _MASTER_INDEX, _MASTER_DESC_INDEX

def get_master_candidates(item_code):
    idx, _ = load_master_index()
    if not item_code:
        return []
    return idx.get(str(item_code).strip().upper(), [])

def find_master_by_description(description):
    _, d_idx = load_master_index()
    if not description:
        return []
    norm_search = re.sub(r'\s+', ' ', description.upper()).strip()
    # exact match
    if norm_search in d_idx:
        return d_idx[norm_search]
        
    # primitive fuzzy - if exact description exists in master desc
    for m_desc, records in d_idx.items():
        if norm_search in m_desc or m_desc in norm_search:
            return records
    return []
