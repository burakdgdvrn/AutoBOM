let hot; // Handsontable instance

document.addEventListener('DOMContentLoaded', function() {
    console.log("Havatek ERP BOM Automation - Initializing...");

    // Theme Toggle Logic
    const themeToggle = document.getElementById('themeToggle');
    const htmlEl = document.documentElement;
    
    // Check saved theme or system preference
    const savedTheme = localStorage.getItem('erp-theme');
    if (savedTheme === 'dark') {
        htmlEl.setAttribute('data-bs-theme', 'dark');
        themeToggle.checked = true;
    } else {
        htmlEl.setAttribute('data-bs-theme', 'light');
        themeToggle.checked = false;
    }

    themeToggle.addEventListener('change', (e) => {
        if (e.target.checked) {
            htmlEl.setAttribute('data-bs-theme', 'dark');
            localStorage.setItem('erp-theme', 'dark');
        } else {
            htmlEl.setAttribute('data-bs-theme', 'light');
            localStorage.setItem('erp-theme', 'light');
        }
        // Force handsontable redraw to apply new CSS variables properly
        if (hot) {
            setTimeout(() => hot.render(), 50);
        }
    });

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
            dropdownMenu: false, // Disabled to remove ugly black boxes on headers
            width: '100%',
            height: '100%', 
            stretchH: 'all',
            autoWrapRow: true,
            autoWrapCol: true,
            minRows: 40, // Increased to 40 to fill the black void with spreadsheet cells
            licenseKey: 'non-commercial-and-evaluation',
            afterChange: updateBadge,
            afterLoadData: updateBadge
        });
    }

    // Full-screen Drag and Drop logic
    const dragOverlay = document.getElementById('drag-overlay');
    let dragCounter = 0; // To prevent flicker with child elements

    document.body.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dragCounter++;
        if (dragCounter === 1) {
            dragOverlay.classList.remove('d-none');
        }
    });

    document.body.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
            dragOverlay.classList.add('d-none');
        }
    });

    document.body.addEventListener('dragover', (e) => {
        e.preventDefault(); // Necessary to allow dropping
    });

    document.body.addEventListener('drop', (e) => {
        e.preventDefault();
        dragCounter = 0;
        dragOverlay.classList.add('d-none');
        
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
        badge.innerText = `${validRows.length} Satır`;
    }
}

window.downloadQAReport = async function() {
    if (!window.lastQAReportData) {
        Toast.fire({
            icon: 'warning',
            title: 'Hata',
            text: 'İndirilecek bir rapor bulunamadı. Lütfen önce kıyaslama yapın.',
            confirmButtonColor: '#2F5496'
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
        
        // Dosyayı blob olarak indir
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
            confirmButtonColor: '#d33'
        });
    }
};

// SweetAlert2 Toast Definition for non-intrusive notifications
const Toast = Swal.mixin({
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 3000,
    timerProgressBar: true,
    background: 'var(--bg-surface)',
    color: 'var(--text-main)',
    didOpen: (toast) => {
        toast.addEventListener('mouseenter', Swal.stopTimer)
        toast.addEventListener('mouseleave', Swal.resumeTimer)
    }
});

// Window global function for input onchange
window.handlePDFUpload = async function(event) {
    const file = event.target.files[0];
    if (!file) return;

    const spinner = document.getElementById('upload-spinner');
    if (spinner) spinner.classList.remove('d-none');

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
            // Filter out empty rows safely
            const filteredData = currentData.filter(r => r && (r.technical_drawing || r.assembly || r.sub_assembly || r.item_code));
            
            const newRows = data.rows;
            hot.loadData([...filteredData, ...newRows]);
            
            // Re-render to ensure grid paints correctly
            setTimeout(() => hot.render(), 100);
            
            // Modern Success Notification
            Toast.fire({
                icon: 'success',
                title: `${newRows.length} satır başarıyla eklendi.`
            });
        } else {
            Toast.fire({
                icon: 'error',
                title: 'İşlem Başarısız',
                text: "Sunucu bir hata döndürdü: " + (data.message || "Bilinmeyen hata"),
                background: 'var(--bg-surface)',
                color: 'var(--text-main)'
            });
        }

    } catch (error) {
        console.error("PDF işleme hatası:", error);
        Toast.fire({
            icon: 'error',
            title: 'Bağlantı Hatası',
            text: "Sunucunun (app.py) çalıştığından emin olun. \nDetay: " + error.message,
            background: 'var(--bg-surface)',
            color: 'var(--text-main)'
        });
    } finally {
        if (spinner) spinner.classList.add('d-none');
        // Reset file input
        event.target.value = '';
    }
};

window.clearTable = function() {
    if (!hot) return;
    
    Swal.fire({
        title: 'Emin misiniz?',
        text: "Tüm tablodaki veriler kalıcı olarak silinecektir!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: 'var(--primary)',
        cancelButtonColor: '#ef4444',
        confirmButtonText: 'Evet, Sil!',
        cancelButtonText: 'İptal',
        background: 'var(--bg-surface)',
        color: 'var(--text-main)'
    }).then((result) => {
        if (result.isConfirmed) {
            hot.loadData([]);
            // Force a re-render and badge update
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

window.exportToExcel = function() {
    if (!hot) return;
    
    // Get headers
    const colHeaders = hot.getColHeader();
    
    // Get valid data (filter out empty rows)
    const allData = hot.getSourceData();
    const validData = allData.filter(r => r && (r.technical_drawing || r.assembly || r.sub_assembly || r.item_code));
    
    if (validData.length === 0) {
        Toast.fire({
            icon: 'info',
            title: 'Tablo Boş',
            text: 'Dışa aktarılacak veri bulunamadı! Lütfen önce PDF yükleyin.',
            background: 'var(--bg-surface)',
            color: 'var(--text-main)'
        });
        return;
    }
    
    // Format data as 2D Array for SheetJS
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
    
    // Create a new workbook and add the worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(dataArray);
    
    // Make columns wider automatically (approximate)
    const colWidths = colHeaders.map(h => ({ wch: Math.max(h.length + 5, 15) }));
    ws['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(wb, ws, "BOM_Listesi");
    
    // Write out the true .xlsx file
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `BOM_Listesi_${dateStr}.xlsx`);
};

// QA State
window.lastQAReportData = null;

// --- Eşleştirme ve Doğrulama Sistemi (BOM QA) ---
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
            text: 'Karşılaştırma yapmak için önce sistemin PDF üzerinden AI tablosunu doldurması gerekiyor.',
            background: 'var(--bg-surface)',
            color: 'var(--text-main)'
        });
        event.target.value = '';
        return;
    }

    const spinner = document.getElementById('upload-spinner');
    if (spinner) {
        spinner.classList.remove('d-none');
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
            window.lastQAReportData = res; // Rapor indirme için kaydet
            
            // Skorları Güncelle
            document.getElementById('stat-perfect').innerText = res.stats.perfect_matches;
            document.getElementById('stat-extra').innerText = res.stats.extra_by_ai;
            document.getElementById('stat-diff').innerText = res.stats.cell_discrepancies;
            document.getElementById('stat-missed').innerText = res.stats.missed_by_ai;
            
            // Tablo Doldurma Yardımcısı (Standart listeler için)
            const fillTable = (id, arr, type) => {
                const tbody = document.getElementById(id);
                tbody.innerHTML = '';
                if(arr.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Bu kategoride kayıt bulunamadı.</td></tr>';
                    return;
                }
                
                arr.forEach(row => {
                    let tr = document.createElement('tr');
                    let html = `<td>${row.technical_drawing || '-'}</td>`;
                    
                    if (type === 'missed') {
                        html += `<td>${row.ex_assembly || '-'}</td><td>${row.ex_sub_assembly || '-'}</td><td>${row.ex_item_code || '-'}</td><td>${row.ex_qty || '-'}</td><td><small class="text-muted">${row.desc || '-'}</small></td>`;
                    } else if (type === 'extra') {
                        html += `<td>${row.ai_assembly || '-'}</td><td>${row.ai_sub_assembly || '-'}</td><td>${row.ai_item_code || '-'}</td><td>${row.ai_qty || '-'}</td><td><small class="text-muted">${row.desc || '-'}</small></td>`;
                    } else if (type === 'perfect') {
                        // perfect match has ai_assembly etc since it's from cell_discrepancies logic
                        html += `<td>${row.ai_assembly || '-'}</td><td>${row.ai_sub_assembly || '-'}</td><td>${row.ai_item_code || '-'}</td><td>${row.ai_qty || '-'}</td>`;
                    }
                    tr.innerHTML = html;
                    tbody.appendChild(tr);
                });
            };
            
            // Hücre Bazlı Farklılıklar (Cell Discrepancies) Tablosu
            const fillDiffTable = (id, arr) => {
                const tbody = document.getElementById(id);
                tbody.innerHTML = '';
                if(arr.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">Uyuşmazlık bulunamadı.</td></tr>';
                    return;
                }
                
                arr.forEach(row => {
                    let tr = document.createElement('tr');
                    
                    const a_class = row.ai_assembly !== row.ex_assembly ? 'bg-warning-subtle text-warning fw-bold' : '';
                    const s_class = row.ai_sub_assembly !== row.ex_sub_assembly ? 'bg-warning-subtle text-warning fw-bold' : '';
                    const i_class = row.ai_item_code !== row.ex_item_code ? 'bg-warning-subtle text-warning fw-bold' : '';
                    const q_class = row.ai_qty !== row.ex_qty ? 'bg-warning-subtle text-warning fw-bold' : '';
                    
                    const formatCell = (ai, ex, cls) => {
                        if (cls) {
                            return `<td class="${cls}"><div class="small text-muted text-decoration-line-through">${ex||'-'}</div><div>${ai||'-'}</div></td>`;
                        }
                        return `<td>${ai||'-'}</td>`;
                    };
                    
                    let html = `<td>${row.technical_drawing || '-'}</td>`;
                    html += formatCell(row.ai_assembly, row.ex_assembly, a_class);
                    html += formatCell(row.ai_sub_assembly, row.ex_sub_assembly, s_class);
                    html += formatCell(row.ai_item_code, row.ex_item_code, i_class);
                    // Qty for ex and ai
                    html += `<td class="${q_class ? 'text-decoration-line-through text-muted' : ''}">${row.ex_qty || '-'}</td>`;
                    html += `<td class="${q_class ? 'text-warning fw-bold' : ''}">${row.ai_qty || '-'}</td>`;
                    
                    tr.innerHTML = html;
                    tbody.appendChild(tr);
                });
            };
            
            // Ham Veri (Raw) Tabloları
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
            
            // Modal'ı Göster
            const modal = new bootstrap.Modal(document.getElementById('compareModal'));
            modal.show();
            
        } else {
            Toast.fire({
                icon: 'error', 
                title: 'Hata', 
                text: "Sunucu hatası: " + (data.message || "Bilinmeyen hata"),
                background: 'var(--bg-surface)',
                color: 'var(--text-main)'
            });
        }
    } catch (error) {
        console.error(error);
        Toast.fire({
            icon: 'error', 
            title: 'Bağlantı Hatası', 
            text: error.message,
            background: 'var(--bg-surface)',
            color: 'var(--text-main)'
        });
    } finally {
        if (spinner) {
            spinner.classList.add('d-none');
            spinner.querySelector('h4').innerText = "PDF İşleniyor...";
            spinner.querySelector('p').innerText = "Yapay Zeka tabloları okuyor ve Spool'ları eşleştiriyor.";
        }
        if (event && event.target) {
            event.target.value = '';
        }
    }
};

