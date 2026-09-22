import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Renk Paletleri
COLORS = {
    "HEADER_BG": "2F5496",
    "HEADER_FG": "FFFFFF",
    "MISSED_BG": "F8CBAD",
    "EXTRA_BG": "DDEBF7",
    "DIFF_BG": "FFE699",
    "PERFECT_BG": "E2EFDA"
}

def create_styled_sheet(wb, title, tab_color, headers, data, is_missed=False):
    ws = wb.create_sheet(title=title)
    ws.sheet_properties.tabColor = tab_color
    
    header_font = Font(bold=True, color=COLORS["HEADER_FG"])
    header_fill = PatternFill(start_color=COLORS["HEADER_BG"], end_color=COLORS["HEADER_BG"], fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    al_c = Alignment(horizontal='center', vertical='center')
    al_l = Alignment(horizontal='left', vertical='center')

    # Başlıkları yaz
    for col_idx, header in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=header)
        c.font = header_font
        c.fill = header_fill
        c.alignment = al_c
        c.border = border
        
    # Otomatik Filtre
    if len(headers) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Verileri yaz
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, val in enumerate(row_data, 1):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.border = border
            c.alignment = al_l if col_idx > 3 else al_c
            
        # Debug Notu Ekleme (Sadece Kaçırılanlar İçin)
        if is_missed:
            # Assembly'ye bakarak basit tahminler
            assembly = row_data[0]
            desc = row_data[4]
            note = ""
            if "PIPE" in str(desc).upper():
                note = "Olası Neden: OCR pipe length okumasında hata (Silik yazı / Kesinti)"
            else:
                note = "Olası Neden: Regex format uyuşmazlığı veya OCR satır birleştirme hatası"
            
            c_note = ws.cell(row=row_idx, column=len(headers)+1, value=note)
            c_note.font = Font(italic=True, color="7F7F7F")

    # Sütun Genişlikleri
    for i in range(1, len(headers) + (2 if is_missed else 1)):
        ws.column_dimensions[get_column_letter(i)].width = 25

    return ws

def generate_qa_excel(report_data, output_path):
    wb = openpyxl.Workbook()
    # Varsayılan "Sheet" sayfasını sil
    wb.remove(wb.active)

    stats = report_data.get("stats", {})

    # 1. ANALİZ ÖZETİ SAYFASI
    ws_summary = wb.create_sheet(title="ANALİZ ÖZETİ")
    ws_summary.sheet_properties.tabColor = "000000"
    
    ws_summary.column_dimensions['A'].width = 35
    ws_summary.column_dimensions['B'].width = 20
    ws_summary.column_dimensions['C'].width = 60

    title_font = Font(bold=True, size=16, color="000000")
    bold_font = Font(bold=True)
    
    ws_summary["A1"] = "BOM YAPAY ZEKA DEBUG & QA RAPORU"
    ws_summary["A1"].font = title_font

    summary_rows = [
        ("Toplam AI Buluntusu:", stats.get("total_ai", 0), ""),
        ("Toplam Master Excel Satırı:", stats.get("total_excel", 0), ""),
        ("", "", ""),
        ("Kusursuz Eşleşenler:", stats.get("perfect_matches", 0), "Sistemin mükemmel çalıştığı satırlar."),
        ("Kaçırılanlar (Debug Gerekli):", stats.get("missed_by_ai", 0), "Kodun/OCR'ın göremediği eksikler. Bu sekmeye odaklanın."),
        ("AI Ekstraları (Kazançlar):", stats.get("extra_by_ai", 0), "Manuel süreçte es geçilen ama AI'ın affetmediği detaylar."),
        ("Hücre Uyuşmazlıkları:", stats.get("cell_discrepancies", 0), "Kısmen eşleşen ama bazı hücreleri (Spool, Miktar vb) farklı olanlar.")
    ]

    for r_idx, (label, val, desc) in enumerate(summary_rows, 3):
        ws_summary.cell(row=r_idx, column=1, value=label).font = bold_font
        ws_summary.cell(row=r_idx, column=2, value=val)
        ws_summary.cell(row=r_idx, column=3, value=desc)

    # 2. KAÇIRILANLAR
    missed_data = [[r.get("technical_drawing"), r.get("ex_assembly"), r.get("ex_sub_assembly"), r.get("ex_item_code"), r.get("ex_qty"), r.get("desc")] for r in report_data.get("missed_by_ai", [])]
    create_styled_sheet(wb, "🔴 KAÇIRILANLAR", "FF0000", ["Technical Drawing", "Assembly", "Sub Assembly", "Item Code", "Excel Qty", "Açıklama"], missed_data, is_missed=True)
    
    if len(missed_data) > 0:
        ws = wb["🔴 KAÇIRILANLAR"]
        ws.cell(row=1, column=7, value="Otomize Debug Notu (Yapay Zeka Tahmini)").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=1, column=7).fill = PatternFill(start_color=COLORS["HEADER_BG"], end_color=COLORS["HEADER_BG"], fill_type="solid")

    # 3. AI EKSTRALARI
    extra_data = [[r.get("technical_drawing"), r.get("ai_assembly"), r.get("ai_sub_assembly"), r.get("ai_item_code"), r.get("ai_qty"), r.get("desc")] for r in report_data.get("extra_by_ai", [])]
    create_styled_sheet(wb, "🔵 AI ÜSTÜNLÜĞÜ", "0070C0", ["Technical Drawing", "Assembly", "Sub Assembly", "Item Code", "AI Qty", "Açıklama"], extra_data)

    # 4. HÜCRE UYUŞMAZLIKLARI
    diff_data = [[r.get("technical_drawing"), r.get("ex_assembly"), r.get("ai_assembly"), r.get("ex_sub_assembly"), r.get("ai_sub_assembly"), r.get("ex_item_code"), r.get("ai_item_code"), r.get("ex_qty"), r.get("ai_qty")] for r in report_data.get("cell_discrepancies", [])]
    create_styled_sheet(wb, "🟡 UYUŞMAZLIKLAR", "FFC000", ["Technical Drawing", "Ex Assembly", "AI Assembly", "Ex Sub", "AI Sub", "Ex Item", "AI Item", "Ex Qty", "AI Qty"], diff_data)

    # 5. KUSURSUZ EŞLEŞENLER
    perfect_data = [[r.get("technical_drawing"), r.get("ai_assembly"), r.get("ai_sub_assembly"), r.get("ai_item_code"), r.get("ai_qty")] for r in report_data.get("perfect_matches", [])]
    create_styled_sheet(wb, "🟢 EŞLEŞENLER", "00B050", ["Technical Drawing", "Assembly", "Sub Assembly", "Item Code", "Qty"], perfect_data)

    # 6. HAM VERİLER (İsteğe Bağlı eklenebilir ama şu an arayüzde var)
    
    wb.save(output_path)
