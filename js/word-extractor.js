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
}
