/**
 * WordExtractor — สกัดตารางจากไฟล์ .docx โดยใช้ mammoth.js
 */
class WordExtractor {
    /**
     * Extract tables from a .docx file
     * @param {File} file
     * @param {Function} onProgress
     * @returns {Promise<Array>} array of table objects
     */
    async extract(file, onProgress) {
        if (onProgress) onProgress(1, 2);

        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;

        if (onProgress) onProgress(2, 2);

        // Parse HTML and extract tables
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const htmlTables = doc.querySelectorAll('table');

        if (htmlTables.length === 0) {
            // Try to detect tab-separated data in paragraphs
            return this._extractFromParagraphs(doc);
        }

        const tables = [];
        htmlTables.forEach((table, index) => {
            const tableData = this._parseHTMLTable(table);
            if (tableData && tableData.rows.length > 0) {
                tables.push({
                    headers: tableData.headers,
                    rows: tableData.rows,
                    columnCount: tableData.headers.length,
                    pageNumber: index + 1,
                    pageRange: [index + 1],
                    source: 'word-table'
                });
            }
        });

        return tables;
    }

    /**
     * Parse an HTML table element into headers and rows
     */
    _parseHTMLTable(tableElement) {
        const allRows = tableElement.querySelectorAll('tr');
        if (allRows.length === 0) return null;

        const data = [];
        let maxCols = 0;

        allRows.forEach(tr => {
            const cells = tr.querySelectorAll('td, th');
            const rowData = [];
            cells.forEach(cell => {
                rowData.push(cell.textContent.trim());
            });
            if (rowData.length > maxCols) maxCols = rowData.length;
            data.push(rowData);
        });

        // Normalize row lengths
        data.forEach(row => {
            while (row.length < maxCols) row.push('');
        });

        return {
            headers: data[0] || [],
            rows: data.slice(1)
        };
    }

    /**
     * Fallback: try to extract table-like data from paragraphs
     * (e.g., tab-separated or consistently structured text)
     */
    _extractFromParagraphs(doc) {
        const paragraphs = doc.querySelectorAll('p');
        const lines = [];

        paragraphs.forEach(p => {
            const text = p.textContent.trim();
            if (text) lines.push(text);
        });

        // Try tab-separated detection
        const tabLines = lines.filter(l => l.includes('\t'));
        if (tabLines.length >= 2) {
            const data = tabLines.map(l => l.split('\t').map(c => c.trim()));
            const maxCols = Math.max(...data.map(r => r.length));
            data.forEach(row => {
                while (row.length < maxCols) row.push('');
            });

            return [{
                headers: data[0],
                rows: data.slice(1),
                columnCount: maxCols,
                pageNumber: 1,
                pageRange: [1],
                source: 'word-paragraphs'
            }];
        }

        // Try pipe-separated or other delimiters
        const delimiters = ['|', ';'];
        for (const delim of delimiters) {
            const delimLines = lines.filter(l => l.includes(delim));
            if (delimLines.length >= 2) {
                const data = delimLines.map(l =>
                    l.split(delim).map(c => c.trim()).filter(c => c !== '')
                );
                const maxCols = Math.max(...data.map(r => r.length));
                data.forEach(row => {
                    while (row.length < maxCols) row.push('');
                });

                return [{
                    headers: data[0],
                    rows: data.slice(1),
                    columnCount: maxCols,
                    pageNumber: 1,
                    pageRange: [1],
                    source: 'word-delimited'
                }];
            }
        }

        return [];
    }

    /**
     * Smart e-GP Extractor for DOCX files (Multi-page)
     * Scans ALL paragraphs and tables sequentially
     * Splits forms by detecting repeating "1." section pattern
     */
    async extractEGP(file, onProgress) {
        if (onProgress) onProgress(1, 3);

        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;

        if (onProgress) onProgress(2, 3);

        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Collect ALL elements in document order using TreeWalker
        const items = []; // { type: 'text'|'table', content: string|parsedTable }
        const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_ELEMENT);
        const visited = new Set();

        let node;
        while (node = walker.nextNode()) {
            if (visited.has(node)) continue;

            if (node.tagName === 'TABLE') {
                visited.add(node);
                // Mark all children as visited
                node.querySelectorAll('*').forEach(c => visited.add(c));
                const td = this._parseHTMLTable(node);
                if (td && td.rows.length > 0) {
                    items.push({ type: 'table', content: td });
                }
            } else if (node.tagName === 'P') {
                visited.add(node);
                const t = node.textContent.trim();
                if (t) items.push({ type: 'text', content: t });
            }
        }

        // Split items into forms by detecting section "1." pattern
        const forms = [];
        let currentForm = [];

        for (const item of items) {
            if (item.type === 'text') {
                const t = item.content;
                // Detect form boundary: "1. หน่วยงาน" or title "ข้อมูลสาระสำคัญ"
                const isSection1 = /^1\s*\./.test(t);
                const isTitle = t.includes('สาระสำคัญ') || t.includes('สาระส');

                if ((isSection1 || isTitle) && currentForm.length > 2) {
                    forms.push(currentForm);
                    currentForm = [];
                }
            }
            currentForm.push(item);
        }
        if (currentForm.length > 0) forms.push(currentForm);

        // Process each form
        const egpRows = [];

        for (const form of forms) {
            const paragraphs = form.filter(i => i.type === 'text').map(i => i.content);
            const tables = form.filter(i => i.type === 'table').map(i => i.content);

            // Parse sections
            const sections = {};
            let currentSection = 0;

            for (const line of paragraphs) {
                const m = line.match(/^(\d)\s*\./);
                if (m) {
                    const num = parseInt(m[1]);
                    if (num >= 1 && num <= 7) {
                        currentSection = num;
                        sections[num] = line.replace(/^\d\s*\.\s*/, '').trim();
                        continue;
                    }
                }
                if (currentSection >= 1 && currentSection <= 5) {
                    sections[currentSection] = (sections[currentSection] || '') + ' ' + line;
                }
            }
            for (const k in sections) sections[k] = (sections[k] || '').trim();

            // Skip if no meaningful sections found (likely a title-only block)
            if (!sections[3] && !sections[4] && tables.length === 0) continue;

            // Section 3: Project Name (Yellow)
            let projName = sections[3] || '';
            projName = projName.replace(/^.*?โครงการ\s*/, '');
            let method = '';
            const mm = projName.match(/\s*โดยวิธี(.*?)$/);
            if (mm) {
                method = mm[1].trim();
                projName = projName.replace(/\s*โดยวิธี.*$/, '').trim();
            }

            // Section 4: Budget (Dark Green)
            let budget = sections[4] || '';
            budget = budget.replace(/^.*?[มณ]\s+/g, '').replace(/\s*บาท.*$/, '').trim();

            // Section 5: Median Price (Light Green)
            let medianPrice = sections[5] || '';
            medianPrice = medianPrice.replace(/^.*?กลาง\s*/, '').replace(/\s*บาท.*$/, '').trim();

            // Table 6: Bidders (Light Blue)
            let biddersStr = '-';
            const table6 = tables.find(t => t.headers.length >= 3 && t.headers.length <= 6);
            if (table6) {
                const bidders = table6.rows.map(r => {
                    const name = r.length >= 2 ? r[r.length - 2] : '';
                    const price = r.length >= 1 ? r[r.length - 1] : '';
                    return `${name}/ ${price} บาท`.trim();
                }).filter(b => b.length > 5);
                if (bidders.length > 0) biddersStr = bidders.join('\n');
            }

            // Table 7: Winners
            let winnersStr = '-';
            let reason = '';
            let contractId = '';
            let contractDate = '';

            const table7 = tables.find(t => t.headers.length >= 7);
            if (table7) {
                const dataRows = table7.rows.filter(r => r.some(c => /\d{10,}/.test(c)));
                if (dataRows.length > 0) {
                    const r = dataRows[0];
                    const rowStr = r.join(' ');

                    const nameCol = r.findIndex(c => /บริษัท|ห้าง|ร้าน|สหกรณ์/.test(c));
                    const name = nameCol >= 0 ? r[nameCol] : (r.length > 2 ? r[2] : '');
                    const priceCol = r.findIndex(c => /^[\d,]+\.\d{2}$/.test(c.trim()));
                    const price = priceCol >= 0 ? r[priceCol] : '';

                    winnersStr = `${name}/ ${price} บาท`.trim();
                    reason = r[r.length - 1] || '';

                    const eGpMatch = rowStr.match(/(6\d{11})/);
                    contractId = eGpMatch ? eGpMatch[1] : '';
                    const dateMatch = rowStr.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
                    contractDate = dateMatch ? dateMatch[1] : '';
                }
            }

            egpRows.push([
                egpRows.length + 1,
                projName || '(ไม่พบชื่อโครงการ)',
                budget, medianPrice, method,
                biddersStr, winnersStr, reason,
                contractId, contractDate
            ]);
        }

        if (onProgress) onProgress(3, 3);

        return [{
            headers: [
                "ลำดับที่", "งานที่จัดซื้อหรือจัดจ้าง",
                "วงเงินที่จะซื้อหรือจ้าง", "ราคากลาง", "วิธีซื้อหรือจ้าง",
                "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
                "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
                "เหตุผลที่คัดเลือกโดยสรุป",
                "เลขที่และวันที่ของสัญญาหรือข้อตกลง", ""
            ],
            rows: egpRows,
            columnCount: 10,
            pageNumber: 1
        }];
    }
}
