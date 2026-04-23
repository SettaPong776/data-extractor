/**
 * PDFExtractor — สกัดตารางจากไฟล์ PDF โดยใช้ PDF.js
 * ใช้ text positions (x, y) เพื่อจัดกลุ่มเป็น rows/columns
 */
class PDFExtractor {
    constructor() {
        // Set PDF.js worker
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        this.Y_TOLERANCE = 5;   // pixels tolerance for same row
        this.X_TOLERANCE = 15;  // pixels tolerance for same column
        this.MIN_ROW_ITEMS = 2; // minimum items per row to be considered table
        this.MIN_TABLE_ROWS = 2; // minimum rows to form a table
    }

    /**
     * Extract tables from a PDF file
     * @param {File} file
     * @param {Function} onProgress — callback(pageNum, totalPages)
     * @returns {Promise<Array>} array of table objects
     */
    async extract(file, onProgress) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const totalPages = pdf.numPages;
        const allTables = [];

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            if (onProgress) onProgress(pageNum, totalPages);

            try {
                const page = await pdf.getPage(pageNum);
                const viewport = page.getViewport({ scale: 1 });
                const textContent = await page.getTextContent();

                const items = textContent.items
                    .filter(item => item.str && item.str.trim() !== '')
                    .map(item => ({
                        text: item.str.trim(),
                        x: Math.round(item.transform[4]),
                        y: Math.round(viewport.height - item.transform[5]),
                        width: item.width,
                        height: Math.abs(item.transform[3]) || 12
                    }));

                if (items.length < this.MIN_ROW_ITEMS) continue;

                const tables = this._extractTablesFromPage(items, pageNum);
                allTables.push(...tables);
            } catch (err) {
                console.warn(`Error processing page ${pageNum}:`, err);
            }
        }

        return this._mergeTables(allTables);
    }

    /**
     * Extract tables from a single page's text items
     */
    _extractTablesFromPage(items, pageNum) {
        // Step 1: Cluster into rows by Y
        const rows = this._clusterByY(items);

        // Step 2: Identify table regions (consecutive rows with multiple items)
        const tableRegions = this._findTableRegions(rows);

        const tables = [];
        for (const region of tableRegions) {
            // Step 3: Detect column positions
            const columns = this._detectColumns(region);
            if (columns.length < 2) continue;

            // Step 4: Build table grid
            const grid = this._buildGrid(region, columns);
            if (grid.length < this.MIN_TABLE_ROWS) continue;

            tables.push({
                headers: grid[0],
                rows: grid.slice(1),
                columnCount: columns.length,
                pageNumber: pageNum
            });
        }

        return tables;
    }

    /**
     * Cluster text items into rows by Y coordinate
     */
    _clusterByY(items) {
        if (items.length === 0) return [];

        const sorted = [...items].sort((a, b) => a.y - b.y);
        const rows = [];
        let currentRow = [sorted[0]];

        for (let i = 1; i < sorted.length; i++) {
            const avgY = currentRow.reduce((s, it) => s + it.y, 0) / currentRow.length;
            if (Math.abs(sorted[i].y - avgY) <= this.Y_TOLERANCE) {
                currentRow.push(sorted[i]);
            } else {
                currentRow.sort((a, b) => a.x - b.x);
                rows.push(currentRow);
                currentRow = [sorted[i]];
            }
        }

        currentRow.sort((a, b) => a.x - b.x);
        rows.push(currentRow);

        return rows;
    }

    /**
     * Find consecutive row groups that look like tables
     */
    _findTableRegions(rows) {
        const regions = [];
        let currentRegion = [];

        for (const row of rows) {
            if (row.length >= this.MIN_ROW_ITEMS) {
                currentRegion.push(row);
            } else {
                if (currentRegion.length >= this.MIN_TABLE_ROWS) {
                    regions.push(currentRegion);
                }
                currentRegion = [];
            }
        }

        if (currentRegion.length >= this.MIN_TABLE_ROWS) {
            regions.push(currentRegion);
        }

        // If no table regions found, try treating all rows as one table
        if (regions.length === 0 && rows.length >= this.MIN_TABLE_ROWS) {
            const multiItemRows = rows.filter(r => r.length >= this.MIN_ROW_ITEMS);
            if (multiItemRows.length >= this.MIN_TABLE_ROWS) {
                regions.push(multiItemRows);
            }
        }

        return regions;
    }

    /**
     * Detect column positions by clustering X coordinates
     */
    _detectColumns(rows) {
        const xPositions = [];
        rows.forEach(row => {
            row.forEach(item => xPositions.push(item.x));
        });

        xPositions.sort((a, b) => a - b);
        if (xPositions.length === 0) return [];

        const clusters = [];
        let cluster = [xPositions[0]];

        for (let i = 1; i < xPositions.length; i++) {
            if (xPositions[i] - cluster[cluster.length - 1] <= this.X_TOLERANCE) {
                cluster.push(xPositions[i]);
            } else {
                clusters.push({
                    x: cluster.reduce((a, b) => a + b) / cluster.length,
                    count: cluster.length
                });
                cluster = [xPositions[i]];
            }
        }
        clusters.push({
            x: cluster.reduce((a, b) => a + b) / cluster.length,
            count: cluster.length
        });

        // Filter out columns that appear in very few rows (noise)
        const threshold = Math.max(2, rows.length * 0.2);
        const validClusters = clusters.filter(c => c.count >= threshold);

        // If filtering removed too many, use all
        const finalClusters = validClusters.length >= 2 ? validClusters : clusters;

        return finalClusters.map(c => c.x).sort((a, b) => a - b);
    }

    /**
     * Build a 2D grid from rows and column positions
     */
    _buildGrid(rows, columnPositions) {
        const grid = [];

        for (const row of rows) {
            const gridRow = new Array(columnPositions.length).fill('');

            for (const item of row) {
                // Find nearest column
                let minDist = Infinity;
                let colIdx = 0;

                for (let i = 0; i < columnPositions.length; i++) {
                    const dist = Math.abs(item.x - columnPositions[i]);
                    if (dist < minDist) {
                        minDist = dist;
                        colIdx = i;
                    }
                }

                // Append text if cell already has content
                if (gridRow[colIdx]) {
                    gridRow[colIdx] += ' ' + item.text;
                } else {
                    gridRow[colIdx] = item.text;
                }
            }

            grid.push(gridRow);
        }

        return grid;
    }

    /**
     * Merge consecutive tables that have the same column structure
     */
    _mergeTables(tables) {
        if (tables.length <= 1) return tables;

        const merged = [];
        let current = { ...tables[0], rows: [...tables[0].rows], pageRange: [tables[0].pageNumber] };

        for (let i = 1; i < tables.length; i++) {
            const next = tables[i];

            if (next.columnCount === current.columnCount) {
                // Check if headers match (repeated header on new page)
                const headersMatch = next.headers.every((h, j) =>
                    h.trim().toLowerCase() === (current.headers[j] || '').trim().toLowerCase()
                );

                if (headersMatch) {
                    // Same table continued — just append rows
                    current.rows = current.rows.concat(next.rows);
                } else {
                    // Same column count but different headers — treat as continuation
                    current.rows.push(next.headers);
                    current.rows = current.rows.concat(next.rows);
                }
                current.pageRange.push(next.pageNumber);
            } else {
                // Different structure — separate table
                merged.push(current);
                current = { ...next, rows: [...next.rows], pageRange: [next.pageNumber] };
            }
        }

        merged.push(current);
        return merged;
    }

    /**
     * Smart e-GP Extractor: Extracts specific fields from Thai e-GP forms
     * Uses LINE-BY-LINE section detection (immune to garbled Thai fonts)
     * 1 Page = 1 Row
     */
    async extractEGP(file, onProgress) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const totalPages = pdf.numPages;
        const egpRows = [];

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            if (onProgress) onProgress(pageNum, totalPages);

            try {
                const page = await pdf.getPage(pageNum);
                const viewport = page.getViewport({ scale: 1 });
                const textContent = await page.getTextContent();

                const items = textContent.items
                    .filter(item => item.str && item.str.trim() !== '')
                    .map(item => ({
                        text: item.str.trim(),
                        x: Math.round(item.transform[4]),
                        y: Math.round(viewport.height - item.transform[5]),
                        width: item.width,
                        height: Math.abs(item.transform[3]) || 12
                    }));

                if (items.length < 3) continue;

                // Group items into lines (sorted by Y then X)
                const lines = this._clusterByY(items);

                // Build sections by detecting leading numbers "1." through "7."
                const sections = {};
                let currentSection = 0;

                for (const line of lines) {
                    const lineText = line.map(i => i.text).join(' ');

                    // Check if line starts a new numbered section
                    // Look for patterns: "3.", "3 .", "3.ชื่อ", etc.
                    let newSection = 0;
                    const firstText = line[0].text;
                    
                    // Case 1: First item is just a digit "3" and second is "." or starts with "."
                    if (/^\d$/.test(firstText) && line.length > 1) {
                        const second = line[1].text;
                        if (second === '.' || second.startsWith('.')) {
                            newSection = parseInt(firstText);
                        }
                    }
                    // Case 2: First item is "3." or "3.ชื่อ"
                    if (!newSection) {
                        const m = firstText.match(/^(\d)\s*\./);
                        if (m) newSection = parseInt(m[1]);
                    }
                    // Case 3: lineText starts with "3." pattern
                    if (!newSection) {
                        const m = lineText.match(/^(\d)\s*\./);
                        if (m) newSection = parseInt(m[1]);
                    }

                    if (newSection >= 1 && newSection <= 7) {
                        currentSection = newSection;
                        // Remove the "N." prefix from content
                        const afterNum = lineText.replace(/^\d\s*\.\s*/, '');
                        sections[currentSection] = afterNum;
                    } else if (currentSection >= 1 && currentSection <= 5) {
                        // Append multi-line content to sections 1-5 only
                        sections[currentSection] = (sections[currentSection] || '') + ' ' + lineText;
                    }
                    // Sections 6 and 7 are tables — handled separately below
                }

                // Clean sections
                for (const k in sections) {
                    sections[k] = (sections[k] || '').trim();
                }

                // === Section 3: Project Name (Yellow) ===
                let projName = sections[3] || '';
                // Strip leading label (handles garbled "ชอื่โครงการ" etc.)
                projName = projName.replace(/^.*?[กา]ร\s+/, function(match) {
                    // Only strip if it looks like "ชื่อโครงการ" label (max 15 chars)
                    return match.length <= 15 ? '' : match;
                });
                // If still has label, try simpler strip
                if (projName.length > 5) {
                    projName = projName.replace(/^.*?โครงการ\s*/, '');
                }

                // Extract method from project name
                let method = '';
                const methodMatch = projName.match(/\s*โดยวิธี(.*?)$/);
                if (methodMatch) {
                    method = methodMatch[1].trim();
                    projName = projName.replace(/\s*โดยวิธี.*$/, '').trim();
                }
                // Fallback method detection from full page text
                if (!method) {
                    const fullText = lines.map(l => l.map(i => i.text).join(' ')).join(' ');
                    if (fullText.match(/วิธี\s*เฉพาะ/)) method = 'เฉพาะเจาะจง';
                    else if (fullText.match(/ประกวดราคา/)) method = 'ประกวดราคา';
                    else if (fullText.match(/วิธี\s*คัดเลือก/)) method = 'คัดเลือก';
                }

                // === Section 4: Budget (Dark Green) ===
                let budget = sections[4] || '';
                budget = budget.replace(/^.*?[มณ]\s+/, function(m) { return m.length <= 15 ? '' : m; });
                budget = budget.replace(/\s*บาท\s*$/, '').trim();

                // === Section 5: Median Price (Light Green) ===
                let medianPrice = sections[5] || '';
                medianPrice = medianPrice.replace(/^.*?กลาง\s*/, '');
                medianPrice = medianPrice.replace(/\s*บาท\s*$/, '').trim();

                // === Tables 6 & 7: Bidders and Winners ===
                const tables = this._extractTablesFromPage(items, pageNum);

                // Identify tables by column count and Tax ID presence
                let table6 = null;
                let table7 = null;
                for (const t of tables) {
                    const hasData = t.rows.some(r => r.some(c => /\d{13}/.test(c)));
                    if (!hasData) continue;
                    if (t.columnCount >= 7 && !table7) {
                        table7 = t;
                    } else if (t.columnCount >= 3 && !table6) {
                        table6 = t;
                    }
                }
                // Fallback: first table = table6, second = table7
                if (!table6 && !table7) {
                    if (tables.length >= 2) { table6 = tables[0]; table7 = tables[1]; }
                    else if (tables.length === 1) { table6 = tables[0]; }
                }

                // Format Table 6 (Bidders — Light Blue)
                let bidders = [];
                if (table6) {
                    const dataRows = table6.rows.filter(r =>
                        r.some(cell => /\d{13}/.test(cell)) || r.some(cell => /[\d,]+\.\d{2}/.test(cell))
                    );
                    bidders = dataRows.map(r => {
                        const name = r.length >= 2 ? r[r.length - 2] : '';
                        const price = r.length >= 1 ? r[r.length - 1] : '';
                        return `${name}/ ${price} บาท`.trim();
                    });
                }
                const biddersStr = bidders.join('\n') || '-';

                // Format Table 7 (Winners — Blue + Grey + Pink + Red)
                let winnersStr = '-';
                let reason = '';
                let contractId = '';
                let contractDate = '';

                if (table7) {
                    const dataRows = table7.rows.filter(r => r.some(cell => /\d{13}/.test(cell)));
                    if (dataRows.length > 0) {
                        const r = dataRows[0];
                        // Column mapping for Table 7:
                        // [0]=ลำดับ [1]=เลขภาษี [2]=ชื่อผู้ขาย [3]=เลขคุมe-GP [4]=เลขสัญญา [5]=วันที่ [6]=จำนวนเงิน [7]=สถานะ [8]=เหตุผล
                        const name = r.length > 2 ? r[2] : '';
                        const price = r.length > 6 ? r[6] : (r.length > 1 ? r[r.length - 2] : '');
                        winnersStr = `${name}/ ${price} บาท`.trim();

                        reason = r[r.length - 1] || '';

                        // e-GP contract ID (Pink) — 12-digit starting with 6
                        const rowStr = r.join(' ');
                        const eGpMatch = rowStr.match(/(6\d{11})/);
                        contractId = eGpMatch ? eGpMatch[1] : (r.length > 3 ? r[3] : '');

                        // Contract date (Red)
                        const dateMatch = rowStr.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
                        contractDate = dateMatch ? dateMatch[1] : (r.length > 5 ? r[5] : '');
                    }
                }

                // Always push a row so the user can see results
                if (projName || budget || items.length > 10) {
                    egpRows.push([
                        egpRows.length + 1,
                        projName || '(ไม่พบชื่อโครงการ)',
                        budget,
                        medianPrice,
                        method,
                        biddersStr,
                        winnersStr,
                        reason,
                        contractId,
                        contractDate
                    ]);
                }

            } catch (err) {
                console.warn(`Error processing page ${pageNum} in e-GP mode:`, err);
                egpRows.push([
                    egpRows.length + 1,
                    `⚠️ ข้อผิดพลาดหน้า ${pageNum}: ${err.message}`,
                    '', '', '', '', '', '', '', ''
                ]);
            }
        }

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
                "เลขที่สัญญา/ข้อตกลง",
                "วันที่ของสัญญา/ข้อตกลง"
            ],
            rows: egpRows,
            columnCount: 10,
            pageNumber: 1
        }];
    }
}
