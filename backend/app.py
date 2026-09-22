import os
import glob
import tempfile
import json
import uuid
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from pdf_to_excel import extract_bom_data_to_json
from compare_bom import compare_excel_and_json
from export_qa import generate_qa_excel

app = Flask(__name__)
CORS(app)  # Allow frontend to make requests

# Use system temp directory to prevent VS Code Live Server from auto-refreshing the page
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "havatek_bom_uploads")
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/api/process-bom', methods=['POST'])
def process_bom_api():
    print("API called: Processing BOM PDF...")
    
    if 'pdf' not in request.files:
        return jsonify({"status": "error", "message": "PDF dosyasi bulunamadi"}), 400
        
    file = request.files['pdf']
    if file.filename == '':
        return jsonify({"status": "error", "message": "Dosya secilmedi"}), 400
        
    if file and file.filename.lower().endswith('.pdf'):
        # Dosyayı gecici olarak kaydet
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        pdf_name = os.path.splitext(file.filename)[0]
        
        try:
            # pdf_to_excel içindeki yeni metodu çağır
            rows = extract_bom_data_to_json(file_path, pdf_name)
            
            # İsteğe bağlı: Dosyayı islem bitince sil
            # os.remove(file_path)
            
            return jsonify({"status": "success", "rows": rows})
        except Exception as e:
            print(f"Error processing {pdf_name}: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({"status": "error", "message": str(e)}), 500
    
    return jsonify({"status": "error", "message": "Gecersiz dosya formati"}), 400

@app.route('/api/compare', methods=['POST'])
def compare_bom_api():
    print("API called: Comparing BOM...")
        
    ai_data_str = request.form.get('ai_data')
    if not ai_data_str:
        return jsonify({"status": "error", "message": "AI verisi bulunamadi"}), 400
        
    try:
        ai_data = json.loads(ai_data_str)
        
        file_path = None
        if 'excel' in request.files and request.files['excel'].filename != '':
            file = request.files['excel']
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
        else:
            # Kullanıcı excel yüklemezse varsayılan Master Excel'i kullan
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            file_path = os.path.join(base_dir, "data", "master_excel", "E-26-RNS-01_BOM_Listesi (1).xlsx")
            if not os.path.exists(file_path):
                return jsonify({"status": "error", "message": f"Ana Excel dosyasi bulunamadi: {file_path}"}), 400
        
        result = compare_excel_and_json(file_path, ai_data)
        
        return jsonify({"status": "success", "result": result})
    except Exception as e:
        print(f"Error comparing: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/export-qa-report', methods=['POST'])
def export_qa_report():
    print("API called: Exporting QA Report...")
    try:
        report_data_str = request.form.get('report_data')
        if not report_data_str:
            return jsonify({"status": "error", "message": "Rapor verisi bulunamadi"}), 400
            
        report_data = json.loads(report_data_str)
        
        # Benzersiz bir dosya adı oluştur
        filename = f"QA_Raporu_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(UPLOAD_FOLDER, filename)
        
        # Excel'i üret
        generate_qa_excel(report_data, output_path)
        
        # Dosyayı gönder
        return send_file(output_path, as_attachment=True, download_name="BOM_QA_Raporu.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
    except Exception as e:
        print(f"Error exporting report: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    print("==============================================")
    print(" Havatek ERP BOM API Server Started on 5000")
    print("==============================================")
    app.run(port=5000, debug=True)
