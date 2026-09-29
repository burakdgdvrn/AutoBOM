let hot; // Handsontable instance

document.addEventListener('DOMContentLoaded', function() {
    console.log("Havatek BOM Engine — Initializing...");

    // Theme Toggle Logic
    const themeToggle = document.getElementById('themeToggle');
    const htmlEl = document.documentElement;
    
    // Check saved theme or system preference
    const savedTheme = localStorage.getItem('erp-theme') || 'dark';
    if (savedTheme === 'light') {
        htmlEl.setAttribute('data-theme', 'light');
        themeToggle.checked = true;
    } else {
        htmlEl.setAttribute('data-theme', 'dark');
        themeToggle.checked = false;
    }

    if (themeToggle) {
        themeToggle.addEventListener('change', (e) => {
            if (e.target.checked) {
                htmlEl.setAttribute('data-theme', 'light');
                localStorage.setItem('erp-theme', 'light');
            } else {
                htmlEl.setAttribute('data-theme', 'dark');
                localStorage.setItem('erp-theme', 'dark');
            }
            if (hot) {
                setTimeout(() => hot.render(), 50);
            }
        });
    }

    // Initialize Handsontable
    const container = document.getElementById('hot-grid');
    if (container) {
        hot = new Handsontable(container, {
            data: [], // Initially empty
            colHeaders: [
                'ID',
                'Technical Drawing',
                'Assembly',
                'Assembly Description',
                'Assembly Weight',
                'Assembly Qty',
                'Sub Assembly',
                'Sub Assembly Defination',
                'Item Code',
                'Qty',
                'Unit Weight',
                'Pose No',
                'Giriş Yapan',
                'Giriş Tarihi'
            ],
            columns: [
                { data: 'id', title: 'ID', readOnly: true, width: 50 },
                { data: 'technical_drawing', title: 'Teknik Çizim', width: 220 },
                { 
                    data: 'assembly', 
                    title: 'Spool (Assembly)', 
                    width: 250,
                    type: 'autocomplete',
                    source: function (query, process) {
                        if (!hot) { process([]); return; }
                        let allData = hot.getSourceData();
                        let uniqueAssemblies = [...new Set(allData.map(r => r.assembly).filter(a => a && a.trim() !== ''))];
                        process(uniqueAssemblies);
                    },
                    strict: false
                },
                { 
                    data: 'assembly_description', 
                    title: 'Assembly Desc.', 
                    width: 250,
                    type: 'autocomplete',
                    source: function (query, process) {
                        if (!hot) { process([]); return; }
                        let allData = hot.getSourceData();
                        let uniqueAssemblies = [...new Set(allData.map(r => r.assembly).filter(a => a && a.trim() !== ''))];
                        process(uniqueAssemblies);
                    },
                    strict: false
                },
                { data: 'assembly_weight', type: 'text' },
                { data: 'assembly_qty', type: 'text' },
                { data: 'sub_assembly', type: 'text' },
                { data: 'sub_assembly_defination', type: 'text' },
                { data: 'item_code', type: 'text' },
                { data: 'qty', type: 'text' },
                { data: 'unit_weight', type: 'text' },
                { data: 'pose_no', type: 'text' },
                { data: 'giris_yapan', type: 'text' },
                { data: 'giris_tarihi', type: 'text' }
            ],
            rowHeaders: true,
            manualColumnResize: true,
            manualRowResize: true,
            filters: true,
            dropdownMenu: false,
            width: '100%',
            height: '100%', 
            stretchH: 'all',
            autoWrapRow: true,
            autoWrapCol: true,
            minRows: 40,
            licenseKey: 'non-commercial-and-evaluation',
            afterChange: updateBadge,
            afterLoadData: updateBadge,
            cells: function (row, col, prop) {
                var cellProperties = {};
                cellProperties.renderer = function (instance, td, row, col, prop, value, cellProperties) {
                    Handsontable.renderers.TextRenderer.apply(this, arguments);
                    
                    const rowData = instance.getSourceDataAtRow(row);
                    if (rowData && typeof rowData.assembly === 'string' && rowData.assembly.includes('BİLİNMEYEN')) {
                        td.classList.add('unknown-row-highlight');
                    } else {
                        td.classList.remove('unknown-row-highlight');
                    }
                };
                return cellProperties;
            }
        });
    }

    // Full-screen Drag and Drop logic
    const dragOverlay = document.getElementById('drag-overlay');
    let dragCounter = 0;

    document.body.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dragCounter++;
        if (dragCounter === 1) {
            dragOverlay.classList.remove('hidden');
        }
    });

    document.body.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
            dragOverlay.classList.add('hidden');
        }
    });

    document.body.addEventListener('dragover', (e) => {
        e.preventDefault();
    });

    document.body.addEventListener('drop', (e) => {
        e.preventDefault();
        dragCounter = 0;
        dragOverlay.classList.add('hidden');
        
        if (e.dataTransfer.files.length > 0) {
            const fileInput = document.getElementById('pdfUpload');
            fileInput.files = e.dataTransfer.files;
            handlePDFUpload({ target: fileInput });
        }
    });
});

function updateBadge() {
    if (!hot) return;
    const badge = document.getElementById('row-count-badge');
    if (badge) {
        const data = hot.getSourceData();
        const validRows = data.filter(r => r && (r.technical_drawing || r.assembly || r.item_code));
        const count = validRows.length;
        badge.innerHTML = `<i class="fas fa-bars-staggered"></i> <span class="stat-value">${count}</span> Satır`;
    }
}

// SweetAlert2 Toast Definition
const Toast = Swal.mixin({
    toast: true,
    position: 'bottom-end',
    showConfirmButton: false,
    timer: 3000,
    timerProgressBar: true,
    didOpen: (toast) => {
        toast.addEventListener('mouseenter', Swal.stopTimer)
        toast.addEventListener('mouseleave', Swal.resumeTimer)
    }
});

// --- PDF Upload ---
window.handlePDFUpload = async function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const spinner = document.getElementById('upload-spinner');
    if (spinner) spinner.classList.remove('hidden');

    const formData = new FormData();
    formData.append('pdf', file);

    try {
        console.log("Uploading file to local API...", file.name);
        const response = await fetch('http://localhost:5000/api/process-bom', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error("Sunucu hatası: " + response.statusText);
        }

        const data = await response.json();
        console.log("Sunucudan dönen veri:", data);
        
        if (data.status === 'success' && data.rows) {
            const currentData = hot.getSourceData();
            const filteredData = currentData.filter(r => r && (r.technical_drawing || r.assembly || r.sub_assembly || r.item_code));
            
            const newRows = data.rows;
            hot.loadData([...filteredData, ...newRows]);
            
            setTimeout(() => hot.render(), 100);
            
            Toast.fire({
                icon: 'success',
                title: `${newRows.length} satır başarıyla eklendi.`
            });
        } else {
            Toast.fire({
                icon: 'error',
                title: 'İşlem Başarısız',
                text: "Sunucu bir hata döndürdü: " + (data.message || "Bilinmeyen hata"),
            });
        }

    } catch (error) {
        console.error("PDF işleme hatası:", error);
        Toast.fire({
            icon: 'error',
            title: 'Bağlantı Hatası',
            text: "Sunucunun (app.py) çalıştığından emin olun. \nDetay: " + error.message,
        });
    } finally {
        if (spinner) spinner.classList.add('hidden');
        event.target.value = '';
    }
};

// --- Clear Table ---
window.clearTable = function() {
    if (!hot) return;
    
    Swal.fire({
        title: 'Emin misiniz?',
        text: "Tüm tablodaki veriler kalıcı olarak silinecektir!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#6366f1',
        cancelButtonColor: '#f87171',
        confirmButtonText: 'Evet, Sil!',
        cancelButtonText: 'İptal',
    }).then((result) => {
        if (result.isConfirmed) {
            hot.loadData([]);
            setTimeout(() => {
                hot.render();
                updateBadge();
            }, 50);
            
            Toast.fire({
                icon: 'success',
                title: 'Tablo temizlendi.'
            });
        }
    });
};

// --- Export to Excel ---
window.exportToExcel = function() {
    if (!hot) return;
    
    const colHeaders = hot.getColHeader();
    const allData = hot.getSourceData();
    const validData = allData.filter(r => r && (r.technical_drawing || r.assembly || r.sub_assembly || r.item_code));
    
    if (validData.length === 0) {
        Toast.fire({
            icon: 'info',
            title: 'Tablo Boş',
            text: 'Dışa aktarılacak veri bulunamadı! Lütfen önce PDF yükleyin.',
        });
        return;
    }
    
    const dataArray = [colHeaders];
    
    validData.forEach(row => {
        dataArray.push([
            row.id || '',
            row.technical_drawing || '',
            row.assembly || '',
            row.assembly_description || '',
            row.assembly_weight || '',
            row.assembly_qty || '',
            row.sub_assembly || '',
            row.sub_assembly_defination || '',
            row.item_code || '',
            row.qty || '',
            row.unit_weight || '',
            row.pose_no || '',
            row.giris_yapan || '',
            row.giris_tarihi || ''
        ]);
    });
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(dataArray);
    
    const colWidths = colHeaders.map(h => ({ wch: Math.max(h.length + 5, 15) }));
    ws['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(wb, ws, "BOM_Listesi");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `BOM_Listesi_${dateStr}.xlsx`);
};

// --- QA Report Download ---
window.lastQAReportData = null;

window.downloadQAReport = async function() {
    if (!window.lastQAReportData) {
        Toast.fire({
            icon: 'warning',
            title: 'Hata',
            text: 'İndirilecek bir rapor bulunamadı. Lütfen önce kıyaslama yapın.',
        });
        return;
    }
    
    Swal.fire({
        title: 'Rapor Hazırlanıyor...',
        text: 'Akıllı Excel analiz raporunuz oluşturuluyor.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });
    
    try {
        const formData = new FormData();
        formData.append('report_data', JSON.stringify(window.lastQAReportData));
        
        const response = await fetch('http://127.0.0.1:5000/api/export-qa-report', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`Sunucu hatası: ${response.status}`);
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'BOM_QA_Raporu.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        
        Toast.fire({
            icon: 'success',
            title: 'Başarılı!',
            text: 'Rapor başarıyla indirildi.',
            timer: 2000,
            showConfirmButton: false
        });
        
    } catch (error) {
        console.error("Export Error:", error);
        Toast.fire({
            icon: 'error',
            title: 'İndirme Hatası',
            text: 'Rapor oluşturulurken bir hata oluştu: ' + error.message,
        });
    }
};

// --- Excel Compare ---
window.handleExcelCompare = async function(event) {
    let file = null;
    if (event && event.target && event.target.files) {
        file = event.target.files[0];
    }

    if (!hot) return;
    const allData = hot.getSourceData();
    const validData = allData.filter(r => r && (r.technical_drawing || r.assembly || r.item_code));
    
    if (validData.length === 0) {
        Toast.fire({
            icon: 'warning', 
            title: 'Tablo Boş', 
            text: 'Karşılaştırma yapmak için önce PDF yükleyip AI tablosunu doldurun.',
        });
        if (event && event.target) event.target.value = '';
        return;
    }

    const spinner = document.getElementById('upload-spinner');
    if (spinner) {
        spinner.classList.remove('hidden');
        spinner.querySelector('h4').innerText = "Excel Karşılaştırılıyor...";
        spinner.querySelector('p').innerText = "Master Excel ile AI bulguları eşleştiriliyor.";
    }

    const formData = new FormData();
    if (file) {
        formData.append('excel', file);
    }
    formData.append('ai_data', JSON.stringify(validData));

    try {
        const response = await fetch('http://localhost:5000/api/compare', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        
        if (data.status === 'success') {
            const res = data.result;
            window.lastQAReportData = res;
            
            // Update score cards
            document.getElementById('stat-perfect').innerText = res.stats.perfect_matches;
            document.getElementById('stat-extra').innerText = res.stats.extra_by_ai;
            document.getElementById('stat-diff').innerText = res.stats.cell_discrepancies;
            document.getElementById('stat-missed').innerText = res.stats.missed_by_ai;
            
            // Table fill helper
            const fillTable = (id, arr, type) => {
                const tbody = document.getElementById(id);
                tbody.innerHTML = '';
                if(arr.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4" style="color: var(--text-tertiary);">Bu kategoride kayıt bulunamadı.</td></tr>';
                    return;
                }
                
                arr.forEach(row => {
                    let tr = document.createElement('tr');
                    let html = `<td>${row.technical_drawing || '-'}</td>`;
                    
                    if (type === 'missed') {
                        html += `<td>${row.ex_assembly || '-'}</td><td>${row.ex_sub_assembly || '-'}</td><td>${row.ex_item_code || '-'}</td><td>${row.ex_qty || '-'}</td><td><small style="color: var(--text-tertiary);">${row.desc || '-'}</small></td>`;
                    } else if (type === 'extra') {
                        html += `<td>${row.ai_assembly || '-'}</td><td>${row.ai_sub_assembly || '-'}</td><td>${row.ai_item_code || '-'}</td><td>${row.ai_qty || '-'}</td><td><small style="color: var(--text-tertiary);">${row.desc || '-'}</small></td>`;
                    } else if (type === 'perfect') {
                        html += `<td>${row.ai_assembly || '-'}</td><td>${row.ai_sub_assembly || '-'}</td><td>${row.ai_item_code || '-'}</td><td>${row.ai_qty || '-'}</td>`;
                    }
                    tr.innerHTML = html;
                    tbody.appendChild(tr);
                });
            };
            
            // Cell discrepancies table
            const fillDiffTable = (id, arr) => {
                const tbody = document.getElementById(id);
                tbody.innerHTML = '';
                if(arr.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4" style="color: var(--text-tertiary);">Uyuşmazlık bulunamadı.</td></tr>';
                    return;
                }
                
                arr.forEach(row => {
                    let tr = document.createElement('tr');
                    
                    const a_class = row.ai_assembly !== row.ex_assembly ? 'style="background: var(--warning-dim); color: var(--warning); font-weight: 700;"' : '';
                    const s_class = row.ai_sub_assembly !== row.ex_sub_assembly ? 'style="background: var(--warning-dim); color: var(--warning); font-weight: 700;"' : '';
                    const i_class = row.ai_item_code !== row.ex_item_code ? 'style="background: var(--warning-dim); color: var(--warning); font-weight: 700;"' : '';
                    const q_class = row.ai_qty !== row.ex_qty ? 'style="background: var(--warning-dim); color: var(--warning); font-weight: 700;"' : '';
                    
                    const formatCell = (ai, ex, cls) => {
                        if (cls) {
                            return `<td ${cls}><div style="font-size: 10px; color: var(--text-tertiary); text-decoration: line-through;">${ex||'-'}</div><div>${ai||'-'}</div></td>`;
                        }
                        return `<td>${ai||'-'}</td>`;
                    };
                    
                    let html = `<td>${row.technical_drawing || '-'}</td>`;
                    html += formatCell(row.ai_assembly, row.ex_assembly, a_class);
                    html += formatCell(row.ai_sub_assembly, row.ex_sub_assembly, s_class);
                    html += formatCell(row.ai_item_code, row.ex_item_code, i_class);
                    html += `<td ${q_class ? 'style="text-decoration: line-through; color: var(--text-tertiary);"' : ''}>${row.ex_qty || '-'}</td>`;
                    html += `<td ${q_class}>${row.ai_qty || '-'}</td>`;
                    
                    tr.innerHTML = html;
                    tbody.appendChild(tr);
                });
            };
            
            // Raw data tables
            const fillRawTable = (id, arr, isExcel) => {
                const tbody = document.getElementById(id);
                tbody.innerHTML = '';
                if(arr.length === 0) return;
                
                arr.forEach(r => {
                    let tr = document.createElement('tr');
                    if (isExcel) {
                        tr.innerHTML = `<td>${r['Technical Drawing'] || '-'}</td><td>${r['Assembly'] || '-'}</td><td>${r['Sub Assembly'] || '-'}</td><td>${r['Item Code'] || '-'}</td><td>${r['Qty'] || '-'}</td>`;
                    } else {
                        tr.innerHTML = `<td>${r.technical_drawing || '-'}</td><td>${r.assembly || '-'}</td><td>${r.sub_assembly || '-'}</td><td>${r.item_code || '-'}</td><td>${r.qty || '-'}</td>`;
                    }
                    tbody.appendChild(tr);
                });
            };
            
            fillTable('tbody-missed', res.missed_by_ai, 'missed');
            fillTable('tbody-extra', res.extra_by_ai, 'extra');
            fillDiffTable('tbody-diff', res.cell_discrepancies);
            fillTable('tbody-perfect', res.perfect_matches, 'perfect');
            
            fillRawTable('tbody-raw-ex', res.raw_filtered_excel || [], true);
            fillRawTable('tbody-raw-ai', res.raw_ai_data || [], false);
            
            // Show modal
            const modal = new bootstrap.Modal(document.getElementById('compareModal'));
            modal.show();
            
        } else {
            Toast.fire({
                icon: 'error', 
                title: 'Hata', 
                text: "Sunucu hatası: " + (data.message || "Bilinmeyen hata"),
            });
        }
    } catch (error) {
        console.error(error);
        Toast.fire({
            icon: 'error', 
            title: 'Bağlantı Hatası', 
            text: error.message,
        });
    } finally {
        if (spinner) {
            spinner.classList.add('hidden');
            spinner.querySelector('h4').innerText = "PDF İşleniyor...";
            spinner.querySelector('p').innerText = "Yapay Zeka tabloları okuyor ve Spool'ları eşleştiriyor.";
        }
        if (event && event.target) {
            event.target.value = '';
        }
    }
};

// --- YOLO AI Test ---
window.openYoloModal = function() {
    const modal = new bootstrap.Modal(document.getElementById('yoloModal'));
    modal.show();
};

window.runYoloTest = async function() {
    const fileInput = document.getElementById('yoloImageInput');
    const file = fileInput.files[0];
    if (!file) {
        Toast.fire({ icon: 'warning', title: 'Lütfen bir resim seçin' });
        return;
    }

    const spinner = document.getElementById('yolo-spinner');
    const resultContainer = document.getElementById('yolo-result-container');
    const resultImg = document.getElementById('yolo-result-img');
    
    spinner.classList.remove('hidden');
    resultContainer.classList.add('hidden');

    const formData = new FormData();
    formData.append('image', file);

    try {
        const response = await fetch('http://127.0.0.1:5000/api/test-yolo', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        
        if (data.status === 'success') {
            resultImg.src = 'data:image/jpeg;base64,' + data.image_base64;
            resultContainer.classList.remove('hidden');
            if (typeof resetYoloZoom === 'function') resetYoloZoom();
            Toast.fire({ icon: 'success', title: 'Yapay Zeka Analizi Tamamlandı' });
        } else {
            Toast.fire({ icon: 'error', title: 'Hata', text: data.message });
        }
    } catch (error) {
        console.error("YOLO Test Hatası:", error);
        Toast.fire({
            icon: 'error',
            title: 'Bağlantı Hatası',
            text: 'Sunucuya ulaşılamadı. Backend API çalışıyor mu? (app.py)',
        });
    } finally {
        spinner.classList.add('hidden');
    }
};

// --- YOLO Zoom Controls ---
let currentZoomPercent = 100;

window.zoomYoloImage = function(step) {
    const img = document.getElementById('yolo-result-img');
    if (!img) return;
    currentZoomPercent += step;
    if (currentZoomPercent < 20) currentZoomPercent = 20;
    if (currentZoomPercent > 1000) currentZoomPercent = 1000;
    img.style.width = currentZoomPercent + '%';
};

window.resetYoloZoom = function() {
    const img = document.getElementById('yolo-result-img');
    if (!img) return;
    currentZoomPercent = 100;
    img.style.width = '100%';
};

// --- YOLO Image Pan & Scroll-Zoom ---
document.addEventListener('DOMContentLoaded', function() {
    const yoloContainer = document.getElementById('yolo-scroll-container');
    const yoloImg = document.getElementById('yolo-result-img');
    
    if (yoloContainer && yoloImg) {
        let isDown = false;
        let startX, startY, scrollLeft, scrollTop;

        yoloContainer.addEventListener('mousedown', (e) => {
            isDown = true;
            yoloContainer.style.cursor = 'grabbing';
            startX = e.pageX - yoloContainer.offsetLeft;
            startY = e.pageY - yoloContainer.offsetTop;
            scrollLeft = yoloContainer.scrollLeft;
            scrollTop = yoloContainer.scrollTop;
            e.preventDefault();
        });

        yoloContainer.addEventListener('mouseleave', () => {
            isDown = false;
            yoloContainer.style.cursor = 'grab';
        });

        yoloContainer.addEventListener('mouseup', () => {
            isDown = false;
            yoloContainer.style.cursor = 'grab';
        });

        yoloContainer.addEventListener('mousemove', (e) => {
            if (!isDown) return;
            e.preventDefault();
            const x = e.pageX - yoloContainer.offsetLeft;
            const y = e.pageY - yoloContainer.offsetTop;
            const walkX = (x - startX) * 1.5;
            const walkY = (y - startY) * 1.5;
            yoloContainer.scrollLeft = scrollLeft - walkX;
            yoloContainer.scrollTop = scrollTop - walkY;
        });
        
        yoloContainer.addEventListener('wheel', (e) => {
            e.preventDefault();
            const delta = e.deltaY > 0 ? -10 : 10;
            zoomYoloImage(delta);
        });
    }
});
