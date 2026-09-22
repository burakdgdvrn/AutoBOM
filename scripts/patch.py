import sys

with open('src/ERP Sistemi.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()
    
new_content = lines[:3349] + [
    '                    <!-- BOM UPLOAD ZONE -->\n',
    '                    <div id="bom-upload-zone" class="upload-zone mb-3 mt-3 text-center p-4 border border-2 border-dashed rounded-3 position-relative" style="background-color: var(--app-surface-2); border-color: var(--app-border) !important; cursor: pointer;">\n',
    '                        <input type="file" id="pdfUpload" class="position-absolute w-100 h-100 top-0 start-0 opacity-0" accept=".pdf" style="cursor: pointer;" onchange="handlePDFUpload(event)">\n',
    '                        <i class="fas fa-cloud-upload-alt fa-3x mb-3 text-primary"></i>\n',
    '                        <h5 class="text-white">PDF Dosyası Yükleyin veya Sürükleyin</h5>\n',
    '                        <p class="text-muted mb-0">Fabrication/Erection Materials ve Cut Pipe Length bilgilerini otomatik işlemek için bir PDF seçin.</p>\n',
    '                        <div id="upload-spinner" class="spinner-border text-primary mt-3 d-none" role="status">\n',
    '                            <span class="visually-hidden">Yükleniyor...</span>\n',
    '                        </div>\n',
    '                    </div>\n',
    '                    <!-- HANDSONTABLE CONTAINER -->\n',
    '                    <div id="hot-grid" class="grid-container rounded mt-3" style="height: 60vh; overflow: hidden; width: 100%;"></div>\n'
] + lines[4535:]

with open('src/ERP Sistemi.html', 'w', encoding='utf-8') as f:
    f.writelines(new_content)
print('Replaced successfully')
