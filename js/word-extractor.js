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
     * Smart e-GP Extractor for DOCX files
     * Parses text paragraphs for sections 1-5 and tables 6-7
     * Returns same 10-column format as PDF extractEGP
     */
    async extractEGP(file, onProgress) {
        if (onProgress) onProgress(1, 3);

        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;

        if (onProgress) onProgress(2, 3);

        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Get all paragraphs (for sections 1-5)
        const paragraphs = doc.querySelectorAll('p');
        const allText = [];
        paragraphs.forEach(p => {
            const t = p.textContent.trim();
            if (t) allText.push(t);
        });

        // Get all tables (for sections 6 and 7)
        const htmlTables = doc.querySelectorAll('table');
        const tables = [];
        htmlTables.forEach(table => {
            const td = this._parseHTMLTable(table);
            if (td && td.rows.length > 0) tables.push(td);
        });

        // Parse sections from paragraphs
        const sections = {};
        let currentSection = 0;

        for (const line of allText) {
            const m = line.match(/^(\d)\s*\./);
            if (m) {
                const num = parseInt(m[1]);
                if (num >= 1 && num <= 7) {
                    currentSection = num;
                    sections[currentSection] = line.replace(/^\d\s*\.\s*/, '').trim();
                    continue;
                }
            }
            if (currentSection >= 1 && currentSection <= 5) {
                sections[currentSection] = (sections[currentSection] || '') + ' ' + line;
            }
        }

        // Clean sections
        for (const k in sections) sections[k] = (sections[k] || '').trim();

        // === Build the row ===
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

        // Table 7: Winners (Blue + Grey + Pink + Red)
        let winnersStr = '-';
        let reason = '';
        let contractId = '';
        let contractDate = '';

        const table7 = tables.find(t => t.headers.length >= 7);
        if (table7) {
            const dataRows = table7.rows.filter(r => r.some(c => /\d{10,}/.test(c)));
            if (dataRows.length > 0) {
                const r = dataRows[0];
                // Find columns by content pattern
                const rowStr = r.join(' ');

                // ชื่อผู้ขาย (Blue) - column with Thai company name
                const nameCol = r.findIndex(c => /บริษัท|ห้าง|ร้าน|สหกรณ์/.test(c));
                const name = nameCol >= 0 ? r[nameCol] : (r.length > 2 ? r[2] : '');

                // จำนวนเงิน - column with price pattern
                const priceCol = r.findIndex(c => /^[\d,]+\.\d{2}$/.test(c.trim()));
                const price = priceCol >= 0 ? r[priceCol] : '';

                winnersStr = `${name}/ ${price} บาท`.trim();

                // เหตุผลที่คัดเลือก (Grey) - last column
                reason = r[r.length - 1] || '';

                // เลขคุมสัญญา e-GP (Pink) - 12-digit starting with 6
                const eGpMatch = rowStr.match(/(6\d{11})/);
                contractId = eGpMatch ? eGpMatch[1] : '';

                // วันที่ทำสัญญา (Red)
                const dateMatch = rowStr.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
                contractDate = dateMatch ? dateMatch[1] : '';
            }
        }

        if (onProgress) onProgress(3, 3);

        // Return single consolidated table with 10 columns
        return [{
            headers: [
                "ลำดับที่",
                "งานที่จัดซื้อหรือจัดจ้าง",
                "วงเงินที่จะซื้อหรือจ้าง",
                "ราคากลาง",
                "วิธีซื้อหรือจ้าง",
                "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
                "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
                "เหตุผลที่คัดเลือกโดยสรุป",
                "เลขที่และวันที่ของสัญญาหรือข้อตกลง",
                ""
            ],
            rows: [[
                1,
                projName || '(ไม่พบชื่อโครงการ)',
                budget,
                medianPrice,
                method,
                biddersStr,
                winnersStr,
                reason,
                contractId,
                contractDate
            ]],
            columnCount: 10,
            pageNumber: 1
        }];
    }
}
