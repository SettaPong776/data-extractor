/**
 * DataExtractorApp — Main Application Controller
 * จัดการ flow: Upload → Extract → Preview → Map → Export
 */
(function () {
    'use strict';

    // ===== STATE =====
    const state = {
        currentStep: 1,
        file: null,
        tables: [],           // All extracted tables
        activeTableIndex: 0,  // Currently viewed table
        previewPage: 1,
        rowsPerPage: 50,
    };

    // ===== MODULES =====
    const pdfExtractor = new PDFExtractor();
    const wordExtractor = new WordExtractor();
    const columnMapper = new ColumnMapper('mapperColumns');
    const excelExporter = new ExcelExporter();

    // ===== DOM ELEMENTS =====
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    const els = {
        uploadZone: $('#uploadZone'),
        fileInput: $('#fileInput'),
        fileInfo: $('#fileInfo'),
        fileName: $('#fileName'),
        fileSize: $('#fileSize'),
        removeFile: $('#removeFile'),
        progressCard: $('#progressCard'),
        progressFill: $('#progressFill'),
        progressText: $('#progressText'),
        progressDetail: $('#progressDetail'),
        tableSelector: $('#tableSelector'),
        tableSelect: $('#tableSelect'),
        previewHead: $('#previewHead'),
        previewBody: $('#previewBody'),
        pagination: $('#pagination'),
        extractionSummary: $('#extractionSummary'),
        mappedPreviewHead: $('#mappedPreviewHead'),
        mappedPreviewBody: $('#mappedPreviewBody'),
        exportSummary: $('#exportSummary'),
        exportFilename: $('#exportFilename'),
        stepLineFill: $('#stepLineFill'),
        templateSelect: $('#templateSelect'),
        templateSettings: $('#templateSettings'),
        exportTitle: $('#exportTitle'),
        exportSubtitle: $('#exportSubtitle'),
        startDate: $('#startDate'),
        endDate: $('#endDate'),
    };

    // ===== INITIALIZATION =====
    function init() {
        setupUpload();
        setupNavigation();
        setupColumnMapper();
        setupExport();
    }

    // ===== UPLOAD =====
    function setupUpload() {
        const zone = els.uploadZone;

        // Click to select file
        zone.addEventListener('click', () => els.fileInput.click());

        // File input change
        els.fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handleFile(e.target.files[0]);
        });

        // Drag & Drop
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('drag-over');
        });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
        });

        // Remove file
        els.removeFile.addEventListener('click', (e) => {
            e.stopPropagation();
            resetAll();
        });
    }

    function handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext !== 'docx') {
            showToast('รองรับเฉพาะไฟล์ DOCX', 'error');
            return;
        }

        state.file = file;
        els.fileName.textContent = file.name;
        els.fileSize.textContent = formatFileSize(file.size);
        els.fileInfo.style.display = 'flex';
        els.uploadZone.style.display = 'none';

        // Start extraction
        startExtraction(file, ext);
    }

    async function startExtraction(file, ext) {
        els.progressCard.style.display = 'block';
        els.progressFill.style.width = '0%';
        els.progressText.textContent = 'กำลังวิเคราะห์เอกสาร...';
        els.progressDetail.textContent = 'เตรียมพร้อม...';

        try {
            let tables;
            const onProgress = (current, total) => {
                const pct = Math.round((current / total) * 100);
                els.progressFill.style.width = pct + '%';
                els.progressDetail.textContent = `หน้า ${current} / ${total}`;
            };

            // Default to 'egp' mode since the UI was removed
            const extractMode = 'egp';

            if (ext === 'pdf') {
                if (extractMode === 'egp') {
                    tables = await pdfExtractor.extractEGP(file, onProgress);
                    // Auto-select procurement template
                    els.templateSelect.value = 'procurement';
                } else {
                    tables = await pdfExtractor.extract(file, onProgress);
                }
            } else {
                if (extractMode === 'egp') {
                    tables = await wordExtractor.extractEGP(file, onProgress);
                    els.templateSelect.value = 'procurement';
                } else {
                    tables = await wordExtractor.extract(file, onProgress);
                }
            }

            els.progressFill.style.width = '100%';

            if (!tables || tables.length === 0) {
                els.progressText.textContent = 'ไม่พบตารางในเอกสาร';
                els.progressDetail.textContent = 'ลองใช้ไฟล์อื่นหรือตรวจสอบว่าเอกสารมีตารางข้อมูล';
                showToast('ไม่พบตารางในเอกสาร', 'error');
                return;
            }

            state.tables = tables;

            // Merge all tables into 1 single table
            if (state.tables.length > 1) {
                // Find the table with the most columns to use as base
                let baseTable = state.tables.reduce((a, b) => a.columnCount >= b.columnCount ? a : b);
                let mergedRows = [];
                
                state.tables.forEach(t => {
                    t.rows.forEach(row => {
                        // Pad rows to match base column count
                        while (row.length < baseTable.columnCount) row.push('');
                        mergedRows.push(row.slice(0, baseTable.columnCount));
                    });
                });

                state.tables = [{
                    headers: baseTable.headers,
                    rows: mergedRows,
                    columnCount: baseTable.columnCount,
                    pageNumber: 1
                }];
            }

            state.activeTableIndex = 0;

            const totalRows = state.tables.reduce((s, t) => s + t.rows.length, 0);
            showToast(`พบ ${totalRows} แถวข้อมูล (รวมเป็น 1 ตาราง)`, 'success');

            // Auto go to preview
            setTimeout(() => goToStep(2), 600);

        } catch (err) {
            console.error('Extraction error:', err);
            els.progressText.textContent = 'เกิดข้อผิดพลาด';
            els.progressDetail.textContent = err.message || 'ไม่สามารถอ่านไฟล์ได้';
            showToast('เกิดข้อผิดพลาดในการอ่านไฟล์', 'error');
        }
    }

    // ===== NAVIGATION =====
    function setupNavigation() {
        $('#btnBackToUpload').addEventListener('click', () => goToStep(1));
        $('#btnGoToMapping').addEventListener('click', () => goToStep(3));
        $('#btnBackToPreview').addEventListener('click', () => goToStep(2));
        $('#btnGoToExport').addEventListener('click', () => goToStep(4));
        $('#btnBackToMapping').addEventListener('click', () => goToStep(3));
        $('#btnStartOver').addEventListener('click', () => {
            resetAll();
            goToStep(1);
        });
    }

    function goToStep(step) {
        state.currentStep = step;

        // Update step panels
        $$('.step-panel').forEach(p => p.classList.remove('active'));
        $(`#step${step}`).classList.add('active');

        // Update step indicators
        $$('.step-indicator .step').forEach(s => {
            const sNum = parseInt(s.dataset.step);
            s.classList.remove('active', 'completed');
            if (sNum === step) s.classList.add('active');
            else if (sNum < step) s.classList.add('completed');
        });

        // Update progress line
        const pct = ((step - 1) / 3) * 100;
        els.stepLineFill.style.width = pct + '%';

        // Step-specific setup
        if (step === 2) renderPreview();
        if (step === 3) setupMappingStep();
        if (step === 4) setupExportStep();
    }

    // ===== PREVIEW (STEP 2) =====
    function renderPreview() {
        const tables = state.tables;
        if (tables.length === 0) return;

        // Table selector
        if (tables.length > 1) {
            els.tableSelector.style.display = 'flex';
            els.tableSelect.innerHTML = tables.map((t, i) =>
                `<option value="${i}">ตาราง ${i + 1} (${t.rows.length} แถว, ${t.columnCount} คอลัมน์)</option>`
            ).join('');
            els.tableSelect.value = state.activeTableIndex;
            els.tableSelect.onchange = () => {
                state.activeTableIndex = parseInt(els.tableSelect.value);
                state.previewPage = 1;
                renderPreviewTable();
            };
        } else {
            els.tableSelector.style.display = 'none';
        }

        // Summary
        const total = tables.reduce((s, t) => s + t.rows.length, 0);
        els.extractionSummary.textContent =
            `พบ ${tables.length} ตาราง, รวม ${total} แถวข้อมูล — จาก ${state.file.name}`;

        renderPreviewTable();
    }

    function renderPreviewTable() {
        const table = state.tables[state.activeTableIndex];
        if (!table) return;

        // Header
        els.previewHead.innerHTML = '<tr>' +
            table.headers.map((h, i) => `<th>${escHtml(h) || `Col ${i + 1}`}</th>`).join('') +
            '</tr>';

        // Paginated rows
        const totalRows = table.rows.length;
        const totalPages = Math.ceil(totalRows / state.rowsPerPage) || 1;
        if (state.previewPage > totalPages) state.previewPage = totalPages;

        const start = (state.previewPage - 1) * state.rowsPerPage;
        const pageRows = table.rows.slice(start, start + state.rowsPerPage);

        els.previewBody.innerHTML = pageRows.map(row =>
            '<tr>' + row.map(cell => `<td>${escHtml(cell)}</td>`).join('') + '</tr>'
        ).join('');

        // Pagination
        renderPagination(totalPages, totalRows);
    }

    function renderPagination(totalPages, totalRows) {
        if (totalPages <= 1) {
            els.pagination.innerHTML = `<small class="text-muted">แสดงทั้งหมด ${totalRows} แถว</small>`;
            return;
        }

        let html = `<small class="text-muted" style="margin-right:8px">หน้า ${state.previewPage}/${totalPages} (${totalRows} แถว)</small>`;
        html += `<button ${state.previewPage === 1 ? 'disabled' : ''} data-page="${state.previewPage - 1}">‹</button>`;

        const maxButtons = 7;
        let startPage = Math.max(1, state.previewPage - Math.floor(maxButtons / 2));
        let endPage = Math.min(totalPages, startPage + maxButtons - 1);
        if (endPage - startPage < maxButtons - 1) startPage = Math.max(1, endPage - maxButtons + 1);

        for (let p = startPage; p <= endPage; p++) {
            html += `<button class="${p === state.previewPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
        }

        html += `<button ${state.previewPage === totalPages ? 'disabled' : ''} data-page="${state.previewPage + 1}">›</button>`;

        els.pagination.innerHTML = html;
        els.pagination.querySelectorAll('button').forEach(btn => {
            btn.addEventListener('click', () => {
                state.previewPage = parseInt(btn.dataset.page);
                renderPreviewTable();
            });
        });
    }

    // ===== COLUMN MAPPING (STEP 3) =====
    function setupMappingStep() {
        const table = state.tables[state.activeTableIndex];
        if (!table) return;

        applyTemplateToMapper(table.headers);
        renderMappedPreview();
        
        els.templateSelect.addEventListener('change', () => {
            applyTemplateToMapper(table.headers);
            renderMappedPreview();
        });
    }

    function applyTemplateToMapper(originalHeaders) {
        const template = els.templateSelect.value;
        
        if (template === 'procurement') {
            const procurementHeaders = [
                "ลำดับที่",
                "งานที่จัดซื้อหรือจัดจ้าง",
                "วงเงินที่จะซื้อหรือจ้าง",
                "ราคากลาง",
                "วิธีซื้อหรือจ้าง",
                "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
                "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
                "เหตุผลที่คัดเลือกโดยสรุป",
                "เลขที่สัญญา/ข้อตกลง",
                "วันที่ของสัญญา/ข้อตกลง"
            ];
            
            // Map original headers to procurement headers
            // We assume the first 9 extracted columns correspond to these
            const mapping = originalHeaders.map((h, i) => ({
                sourceIndex: i,
                sourceName: h || `คอลัมน์ ${i + 1}`,
                targetName: procurementHeaders[i] || h || `คอลัมน์ ${i + 1}`,
                enabled: i < procurementHeaders.length,
                order: i
            }));
            
            // Set directly in column mapper
            columnMapper.mapping = mapping;
            columnMapper.render();
        } else {
            // Custom - default mapping
            columnMapper.setHeaders(originalHeaders);
        }
    }

    function setupColumnMapper() {
        columnMapper.onChange(() => renderMappedPreview());
    }

    function renderMappedPreview() {
        const table = state.tables[state.activeTableIndex];
        if (!table) return;

        const mapped = columnMapper.applyMapping(table.rows);

        els.mappedPreviewHead.innerHTML = '<tr>' +
            mapped.headers.map(h => `<th>${escHtml(h)}</th>`).join('') +
            '</tr>';

        // Show first 10 rows as preview
        const previewRows = mapped.rows.slice(0, 10);
        els.mappedPreviewBody.innerHTML = previewRows.map(row =>
            '<tr>' + row.map(cell => `<td>${escHtml(cell)}</td>`).join('') + '</tr>'
        ).join('');

        if (mapped.rows.length > 10) {
            els.mappedPreviewBody.innerHTML +=
                `<tr><td colspan="${mapped.headers.length}" style="text-align:center;color:var(--text-muted);font-style:italic">... อีก ${mapped.rows.length - 10} แถว</td></tr>`;
        }
    }

    // ===== EXPORT (STEP 4) =====
    function setupExportStep() {
        const table = state.tables[state.activeTableIndex];
        if (!table) return;

        const template = els.templateSelect.value;
        if (template === 'procurement') {
            els.templateSettings.style.display = 'block';
        } else {
            els.templateSettings.style.display = 'none';
        }

        const mapped = columnMapper.applyMapping(table.rows);
        els.exportSummary.textContent =
            `${mapped.headers.length} คอลัมน์ × ${mapped.rows.length} แถว — พร้อมดาวน์โหลด`;

        // Default filename from source file
        const baseName = state.file.name.replace(/\.[^.]+$/, '');
        els.exportFilename.value = baseName + '_extracted';
    }

    function setupExport() {
        // Auto-update subtitle from date pickers
        function updateSubtitle() {
            const thaiMonths = ['', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.']; // Fix month array, Jan is index 1 or 12?
            // Actually new Date().getMonth() returns 0 for Jan, 1 for Feb
            const thaiMonthsReal = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
            
            const formatThai = (dateStr) => {
                if (!dateStr) return '';
                const d = new Date(dateStr);
                if (isNaN(d)) return '';
                return `${d.getDate()} ${thaiMonthsReal[d.getMonth()]} ${d.getFullYear() + 543}`;
            };

            const start = formatThai(els.startDate.value);
            const end = formatThai(els.endDate.value);

            if (start && end) {
                els.exportSubtitle.value = `วันที่ ${start} ถึง วันที่ ${end}`;
            } else if (start) {
                els.exportSubtitle.value = `ตั้งแต่วันที่ ${start}`;
            } else if (end) {
                els.exportSubtitle.value = `ถึงวันที่ ${end}`;
            }
        }

        if (els.startDate && els.endDate) {
            els.startDate.addEventListener('change', updateSubtitle);
            els.endDate.addEventListener('change', updateSubtitle);
        }

        $('#btnExport').addEventListener('click', () => {
            try {
                const table = state.tables[state.activeTableIndex];
                const mapped = columnMapper.applyMapping(table.rows);
                const filename = els.exportFilename.value.trim() || 'extracted_data';
                
                const template = els.templateSelect.value;
                const templateConfig = template === 'procurement' ? {
                    type: 'procurement',
                    title: els.exportTitle.value.trim(),
                    subtitle: els.exportSubtitle.value.trim()
                } : null;

                // Always merge all tables into 1 sheet
                if (state.tables.length > 1) {
                    let allRows = [];
                    let finalHeaders = mapped.headers;
                    
                    state.tables.forEach((t) => {
                        if (t.columnCount === table.columnCount) {
                            const m = columnMapper.applyMapping(t.rows);
                            allRows.push(...m.rows);
                        } else {
                            allRows.push(...t.rows);
                        }
                    });
                    
                    excelExporter.export(
                        { headers: finalHeaders, rows: allRows },
                        filename,
                        templateConfig
                    );
                } else {
                    excelExporter.export(mapped, filename, templateConfig);
                }

                showToast('ดาวน์โหลดไฟล์ Excel สำเร็จ!', 'success');
            } catch (err) {
                console.error('Export error:', err);
                showToast('เกิดข้อผิดพลาดในการส่งออก', 'error');
            }
        });
    }

    // ===== RESET =====
    function resetAll() {
        state.file = null;
        state.tables = [];
        state.activeTableIndex = 0;
        state.previewPage = 1;

        els.fileInput.value = '';
        els.fileInfo.style.display = 'none';
        els.uploadZone.style.display = '';
        els.progressCard.style.display = 'none';
        els.previewHead.innerHTML = '';
        els.previewBody.innerHTML = '';
        els.pagination.innerHTML = '';
    }

    // ===== UTILITIES =====
    function escHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1048576).toFixed(1) + ' MB';
    }

    function showToast(message, type = 'info') {
        const container = $('#toastContainer');
        const icons = { success: 'check_circle', error: 'error', info: 'info' };

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `
            <span class="material-icons-round">${icons[type] || 'info'}</span>
            <span>${message}</span>
        `;

        container.appendChild(toast);

        setTimeout(() => {
            toast.style.animation = 'toastOut 0.3s ease forwards';
            setTimeout(() => toast.remove(), 300);
        }, 3500);
    }

    // ===== START =====
    document.addEventListener('DOMContentLoaded', init);
})();
