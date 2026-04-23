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

                if (items.length < this.MIN_ROW_ITEMS) continue;
                
                // Sorting for fullText (Reading order: top-bottom, left-right)
                const sortedItems = [...items].sort((a, b) => {
                    if (Math.abs(a.y - b.y) <= this.Y_TOLERANCE) {
                        return a.x - b.x;
                    }
                    return a.y - b.y;
                });
                const fullText = sortedItems.map(i => i.text).join(' ');

                // Extract Fields using more permissive Regex
                const pProj = fullText.match(/ชื่อโครงการ\s*(.*?)\s*(?:งบประมาณ|4\.|ราคากลาง)/);
                let projName = pProj ? pProj[1].trim() : '';

                const pBudg = fullText.match(/งบประมาณ\s*(.*?)\s*(?:ราคากลาง|5\.|รายชื่อผู้เสนอ|6\.)/);
                let budget = pBudg ? pBudg[1].trim() : '';

                const pMed = fullText.match(/ราคากลาง\s*(.*?)\s*(?:รายชื่อ|6\.|ผู้ได้รับ|7\.)/);
                let medianPrice = pMed ? pMed[1].trim() : '';

                let method = '';
                const pMethod = projName.match(/วิธี([ก-๙a-zA-Z]+)/);
                if (pMethod) {
                    method = pMethod[1].trim();
                } else if (fullText.includes('เฉพาะเจาะจง')) {
                    method = 'เฉพาะเจาะจง';
                } else if (fullText.includes('ประกวดราคา')) {
                    method = 'ประกวดราคา';
                } else if (fullText.includes('คัดเลือก')) {
                    method = 'คัดเลือก';
                }

                // Extract Tables to find Bidders and Winners
                const tables = this._extractTablesFromPage(items, pageNum);
                
                // Find Table 6 and 7
                let table6 = tables.find(t => t.headers.join(' ').includes('ราคาที่เสนอ') || t.headers.join(' ').includes('รายชื่อผู้เสนอราคา'));
                if (!table6) table6 = tables.find(t => t.rows.length > 0 && t.rows[0].join(' ').includes('รายชื่อผู้เสนอราคา'));

                let table7 = tables.find(t => t.headers.join(' ').includes('เหตุผลที่คัดเลือก') || t.headers.join(' ').includes('ชื่อผู้ขาย'));
                if (!table7) table7 = tables.find(t => t.rows.length > 0 && t.rows[0].join(' ').includes('เหตุผลที่คัดเลือก'));

                // Format Table 6 (Bidders)
                let bidders = [];
                if (table6) {
                    const dataRows = table6.rows.filter(r => !r.join(' ').includes('รายชื่อผู้เสนอราคา') && r.join('').trim().length > 0);
                    bidders = dataRows.map(r => {
                        const name = r.length >= 2 ? r[r.length - 2] : '';
                        const price = r.length >= 1 ? r[r.length - 1] : r.join(' ');
                        return `${name} / ${price}`.trim();
                    });
                }
                const biddersStr = bidders.join('\n') || '-';

                // Format Table 7 (Winners)
                let winners = [];
                let reason = '';
                let contractInfo = '';
                
                if (table7) {
                    const dataRows = table7.rows.filter(r => !r.join(' ').includes('เหตุผลที่คัดเลือก') && r.join('').trim().length > 0);
                    if (dataRows.length > 0) {
                        const r = dataRows[0]; // Take first winner
                        const name = r.length > 2 ? r[2] : '';
                        const price = r.length > 6 ? r[6] : (r.length > 1 ? r[r.length - 2] : '');
                        winners.push(`${name} / ${price}`);
                        
                        reason = r[r.length - 1] || '';
                        
                        const contractNo = r.length > 4 ? r[4] : '';
                        const contractDate = r.length > 5 ? r[5] : '';
                        contractInfo = `${contractNo} ${contractDate}`.trim();
                    }
                }
                const winnersStr = winners.join('\n') || '-';

                // We push everything so user sees what happened
                if (projName || budget || fullText.length > 50) {
                    if (!projName && !budget) {
                        projName = "⚠️ สกัดข้อมูลไม่สำเร็จ";
                        budget = fullText.substring(0, 150) + "... (โปรดตรวจสอบไฟล์ต้นฉบับว่าใช่รูปแบบ e-GP หรือไม่ หรือเป็นไฟล์สแกน)";
                    }

                    egpRows.push([
                        egpRows.length + 1, // index
                        projName,
                        budget,
                        medianPrice,
                        method,
                        biddersStr,
                        winnersStr,
                        reason,
                        contractInfo
                    ]);
                } else if (fullText.trim().length === 0) {
                    egpRows.push([
                        egpRows.length + 1,
                        "⚠️ ไฟล์นี้ไม่มีข้อความ (อาจเป็นไฟล์รูปภาพ/สแกน โปรแกรมไม่สามารถอ่านได้)",
                        "", "", "", "", "", "", ""
                    ]);
                }

            } catch (err) {
                console.warn(`Error processing page ${pageNum} in e-GP mode:`, err);
            }
        }

        // Return a single consolidated table
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
                "เลขที่และวันที่ของสัญญาหรือข้อตกลงในการซื้อหรือจ้าง"
            ],
            rows: egpRows,
            columnCount: 9,
            pageNumber: 1
        }];
    }
}
