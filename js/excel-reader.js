/**
 * ExcelReader — อ่านข้อมูลจากไฟล์ Excel (.xlsx, .xls) รูปแบบ สขร.1
 * ใช้ SheetJS (xlsx.full.min.js) ที่ติดตั้งอยู่แล้วในโปรเจกต์
 * 
 * Source Excel columns (สขร.1):
 *   A = ลำดับ
 *   B = ชื่อผู้ประกอบการ
 *   C = รายการพัสดุที่จัดซื้อจัดจ้าง
 *   D = จำนวนเงินรวมที่จัดซื้อจัดจ้าง
 *   E = วันที่
 *   F = เลขที่
 *   G = เหตุผลสนับสนุน
 *
 * Output: 10-column procurement format
 */
class ExcelReader {

    /**
     * Standard 10-column procurement headers
     */
    static PROCUREMENT_HEADERS = [
        "ลำดับที่",
        "งานที่จัดซื้อหรือจัดจ้าง",
        "วงเงินที่จะซื้อหรือจ้าง",
        "ราคากลาง",
        "วิธีซื้อหรือจ้าง",
        "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
        "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
        "เหตุผลที่คัดเลือกโดยสรุป",
        "เลขที่สัญญา/ข้อตกลง",
        "วันที่ทำสัญญา/ข้อตกลง"
    ];

    /**
     * Keywords used to detect the header row in Excel
     */
    static HEADER_KEYWORDS = [
        'ลำดับ', 'ลําดับ', 'ผู้ประกอบการ', 'ผู้ประกอบกำร',
        'รายการพัสดุ', 'รำยกำรพัสดุ', 'จำนวนเงิน', 'จํานวนเงิน', 'จำนวนเงินรวม',
        'วันที่', 'เลขที่', 'เหตุผล', 'จัดซื้อ', 'จัดจ้าง',
        'สนับสนุน', 'อ้างอิง', 'เอกสาร'
    ];

    /**
     * Extract tables from an Excel file (raw mode — no mapping)
     */
    async extract(file, onProgress) {
        if (onProgress) onProgress(1, 3);

        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

        if (onProgress) onProgress(2, 3);

        const tables = [];

        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) return;

            const allData = this._readAllRows(worksheet);
            if (!allData || allData.length < 2) return;

            // Find header row
            const headerIdx = this._findHeaderRow(allData);
            const headers = allData[headerIdx];
            const rows = allData.slice(headerIdx + 1).filter(r => r.some(c => c !== ''));

            if (rows.length === 0) return;

            tables.push({
                headers: headers,
                rows: rows,
                columnCount: headers.length,
                pageNumber: index + 1,
                pageRange: [index + 1],
                source: 'excel-sheet',
                sheetName: sheetName
            });
        });

        if (onProgress) onProgress(3, 3);
        console.log(`[ExcelReader] Raw extract: ${tables.length} sheet(s)`);
        return tables;
    }

    /**
     * e-GP / สขร.1 extraction mode for Excel files
     * Smart detect header → map columns → group by item → output 10-column format
     */
    async extractEGP(file, onProgress) {
        if (onProgress) onProgress(1, 3);

        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

        if (onProgress) onProgress(2, 3);

        const allProcurementRows = [];

        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;

            const allData = this._readAllRows(worksheet);
            if (!allData || allData.length < 2) continue;

            console.log(`[ExcelReader] Sheet "${sheetName}": ${allData.length} total rows`);

            // Step 1: Find header row
            const headerIdx = this._findHeaderRow(allData);
            const headerRow = allData[headerIdx];
            console.log(`[ExcelReader] Header at row ${headerIdx + 1}:`, headerRow);

            // Step 2: Detect column indices by keywords
            const colMap = this._detectColumns(headerRow);
            console.log(`[ExcelReader] Column mapping:`, colMap);

            // Step 3: Get data rows (after header, skip sub-headers and empty rows)
            const dataRows = allData.slice(headerIdx + 1).filter(r => {
                // Must have at least some content
                if (r.every(c => !c || c.trim() === '')) return false;
                
                // Skip repeated header rows (e.g. spanning multiple pages in Excel)
                // If it looks exactly like a header (has "ผู้ประกอบการ" and "จำนวนเงิน")
                const str = r.join(' ').replace(/\s+/g, '');
                if (str.includes('ผู้ประกอบการ') && str.includes('จำนวนเงิน')) return false;
                
                // Skip rows like "(1)", "(2)" etc in the first column
                if (r[0] && /^\s*\(\s*\d+\s*\)\s*$/.test(r[0]) && r.slice(1).every(c => !c || c.trim() === '')) return false;
                
                return true;
            });

            console.log(`[ExcelReader] Data rows after filtering: ${dataRows.length}`);

            // Step 4: Group rows by item (Column C)
            // Rows with the same item description are multiple bidders for the same procurement
            const groups = this._groupByItem(dataRows, colMap);
            console.log(`[ExcelReader] Grouped into ${groups.length} procurement items`);

            // Step 5: Transform each group into a procurement row
            for (const group of groups) {
                const row = this._buildProcurementRow(group, colMap, allProcurementRows.length + 1);
                if (row) allProcurementRows.push(row);
            }
        }

        if (onProgress) onProgress(3, 3);

        console.log(`[ExcelReader] Final output: ${allProcurementRows.length} procurement rows`);

        if (allProcurementRows.length === 0) {
            // Fallback to raw extraction
            console.log(`[ExcelReader] No procurement data found, falling back to raw mode`);
            return this.extract(file, null);
        }

        return [{
            headers: ExcelReader.PROCUREMENT_HEADERS,
            rows: allProcurementRows,
            columnCount: 10,
            pageNumber: 1,
            source: 'excel-procurement'
        }];
    }

    // ==========================================
    // Private helpers
    // ==========================================

    /**
     * Read ALL rows from worksheet as normalized string arrays
     */
    _readAllRows(worksheet) {
        const ref = worksheet['!ref'];
        if (!ref) return null;

        const data = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: '',
            raw: false,
            dateNF: 'd/m/yyyy'
        });

        if (!data || data.length === 0) return null;

        // Find max columns
        let maxCols = 0;
        data.forEach(row => { if (row.length > maxCols) maxCols = row.length; });
        if (maxCols === 0) return null;

        // Normalize: convert all to trimmed strings, pad to maxCols
        return data.map(row => {
            const normalized = [];
            for (let i = 0; i < maxCols; i++) {
                const cell = row[i];
                normalized.push(cell === null || cell === undefined ? '' : String(cell).trim());
            }
            return normalized;
        });
    }

    /**
     * Find the header row by scanning top rows for keyword matches
     */
    _findHeaderRow(allData) {
        const scanLimit = Math.min(allData.length, 20);
        let bestIdx = 0;
        let bestScore = 0;

        for (let i = 0; i < scanLimit; i++) {
            const rowStr = allData[i].join(' ');
            const nonEmpty = allData[i].filter(c => c !== '').length;
            if (nonEmpty < 2) continue;

            const score = ExcelReader.HEADER_KEYWORDS.filter(kw => rowStr.includes(kw)).length;
            if (score > bestScore) {
                bestScore = score;
                bestIdx = i;
            }
        }

        if (bestScore >= 2) {
            console.log(`[ExcelReader] Header row at index ${bestIdx} (score: ${bestScore})`);
            return bestIdx;
        }

        // Fallback: first row with 3+ non-empty cells
        for (let i = 0; i < scanLimit; i++) {
            if (allData[i].filter(c => c !== '').length >= 3) return i;
        }

        return 0;
    }

    /**
     * Detect which column index maps to which data field
     * Returns: { company: idx, item: idx, amount: idx, date: idx, refNo: idx, reason: idx }
     */
    _detectColumns(headerRow) {
        const map = {
            seq: -1,       // A: ลำดับ
            company: -1,   // B: ชื่อผู้ประกอบการ
            item: -1,      // C: รายการพัสดุ
            amount: -1,    // D: จำนวนเงินรวม
            date: -1,      // E: วันที่
            refNo: -1,     // F: เลขที่
            reason: -1     // G: เหตุผล
        };

        for (let i = 0; i < headerRow.length; i++) {
            const h = headerRow[i];
            if (!h) continue;

            // Seq column
            if (map.seq < 0 && /ลำดับ|ลําดับ|ที่/.test(h) && !/วันที่|เลขที่/.test(h)) {
                map.seq = i;
            }
            // Company name column
            if (map.company < 0 && /ผู้ประกอบ|ชื่อผู้|ผู้ขาย|ผู้รับจ้าง|ผู้ประกอบกำร/.test(h)) {
                map.company = i;
            }
            // Item column
            else if (map.item < 0 && /รายการ|พัสดุ|รำยกำร|จัดซื้อ|จัดจ้าง/.test(h)) {
                map.item = i;
            }
            // Amount column
            else if (map.amount < 0 && /จำนวนเงิน|จํานวนเงิน|เงินรวม|ราคา|วงเงิน/.test(h)) {
                map.amount = i;
            }
            // Date column
            else if (map.date < 0 && /วันที่/.test(h)) {
                map.date = i;
            }
            // Reference number column
            else if (map.refNo < 0 && /เลขที่|สัญญา|ข้อตกลง|อ้างอิง|เอกสาร/.test(h)) {
                map.refNo = i;
            }
            // Reason column
            else if (map.reason < 0 && /เหตุผล|สนับสนุน/.test(h)) {
                map.reason = i;
            }
        }

        // Fallback: if keywords didn't match, use default สขร.1 positions (A=0,B=1,C=2,D=3,E=4,F=5,G=6)
        if (map.seq < 0) map.seq = 0;
        if (map.company < 0) map.company = 1;
        if (map.item < 0) map.item = 2;
        if (map.amount < 0) map.amount = 3;
        if (map.date < 0) map.date = 4;
        if (map.refNo < 0) map.refNo = 5;
        if (map.reason < 0) map.reason = 6;

        return map;
    }

    /**
     * Group consecutive rows by item description (Column C)
     * Rows with the same item (or empty item = same as previous) form a group
     *
     * Returns array of groups: [ [row1, row2, ...], ... ]
     */
    _groupByItem(dataRows, colMap) {
        const groups = [];
        let currentGroup = [];

        for (const row of dataRows) {
            const seq = (row[colMap.seq] || '').trim();
            const item = (row[colMap.item] || '').trim();
            const company = (row[colMap.company] || '').trim();
            const amount = (row[colMap.amount] || '').trim();
            const date = (row[colMap.date] || '').trim();
            const refNo = (row[colMap.refNo] || '').trim();

            // Skip rows that have absolutely no useful data
            if (!seq && !company && !item && !amount && !date && !refNo) continue;

            // Determine if this is a NEW procurement item
            // It's a new item if it has a sequence number, or if it has a date/refNo, 
            // or if it has an item description AND currentGroup is empty
            const isNewItem = (seq && /^\d+$/.test(seq)) || 
                              (date && date.length > 4) || 
                              (refNo && refNo.length > 3) ||
                              (item && currentGroup.length === 0);

            if (isNewItem && currentGroup.length > 0) {
                // Push previous group and start new
                groups.push(currentGroup);
                currentGroup = [row];
            } else if (isNewItem) {
                // First group
                currentGroup = [row];
            } else {
                // Continuation (additional bidder for the same procurement)
                // Ensure it has at least company or amount to be meaningful
                if (company || amount) {
                    currentGroup.push(row);
                }
            }
        }

        // Don't forget the last group
        if (currentGroup.length > 0) {
            groups.push(currentGroup);
        }

        return groups;
    }

    /**
     * Build a 10-column procurement row from a group of source rows
     */
    _buildProcurementRow(group, colMap, seqNumber) {
        if (group.length === 0) return null;

        const firstRow = group[0];

        // Column 1: งานที่จัดซื้อหรือจัดจ้าง — from item column
        const item = (firstRow[colMap.item] || '').trim();

        // Column 2 & 3: วงเงิน / ราคากลาง — from amount column (use first row's amount)
        let amount = (firstRow[colMap.amount] || '').trim();
        amount = this._formatCurrency(amount);

        // Column 4: วิธีซื้อหรือจ้าง — leave empty for สขร.1
        const method = '';

        // Column 5: รายชื่อผู้เสนอราคาและราคาที่เสนอ — ALL bidders combined
        const allBidders = [];
        for (const row of group) {
            const company = (row[colMap.company] || '').trim();
            let price = (row[colMap.amount] || '').trim();
            if (company && price) {
                price = this._formatCurrency(price);
                allBidders.push(`${company} / ${price} บาท`);
            } else if (company) {
                allBidders.push(company);
            }
        }
        const biddersStr = allBidders.length > 0 ? allBidders.join('\n') : '-';

        // Column 6: ผู้ได้รับการคัดเลือก — first bidder (winner)
        const winnerStr = allBidders.length > 0 ? allBidders[0] : '-';

        // Column 7: เหตุผลที่คัดเลือกโดยสรุป
        let reason = (firstRow[colMap.reason] || '').trim();

        // Column 8: เลขที่สัญญา/ข้อตกลง
        let refNo = (firstRow[colMap.refNo] || '').trim();

        // Column 9: วันที่ทำสัญญา/ข้อตกลง
        let dateStr = (firstRow[colMap.date] || '').trim();
        dateStr = this._formatThaiDate(dateStr);
        
        // If firstRow didn't have reason/refNo/date, try to find it in the group
        // This handles cases where merged cells cause data to fall into the 2nd row
        if (!reason) reason = (group.find(r => r[colMap.reason] && r[colMap.reason].trim() !== '') || [])[colMap.reason] || '';
        if (!refNo) refNo = (group.find(r => r[colMap.refNo] && r[colMap.refNo].trim() !== '') || [])[colMap.refNo] || '';
        if (!dateStr) {
            let rawDate = (group.find(r => r[colMap.date] && r[colMap.date].trim() !== '') || [])[colMap.date] || '';
            dateStr = this._formatThaiDate(rawDate);
        }

        return [
            seqNumber,       // 0: ลำดับที่
            item,            // 1: งานที่จัดซื้อหรือจัดจ้าง
            amount,          // 2: วงเงินที่จะซื้อหรือจ้าง
            amount,          // 3: ราคากลาง (same as วงเงิน for สขร.1)
            method,          // 4: วิธีซื้อหรือจ้าง
            biddersStr,      // 5: รายชื่อผู้เสนอราคาและราคาที่เสนอ
            winnerStr,       // 6: ผู้ได้รับการคัดเลือก
            reason,          // 7: เหตุผลที่คัดเลือกโดยสรุป
            refNo,           // 8: เลขที่สัญญา/ข้อตกลง
            dateStr           // 9: วันที่ทำสัญญา/ข้อตกลง
        ];
    }

    /**
     * Format number to currency with commas and 2 decimals
     * e.g. "3840" → "3,840.00", "633.00" → "633.00"
     */
    _formatCurrency(text) {
        if (!text) return '';
        const str = String(text).trim();
        const match = str.match(/[\d,]+(\.\d+)?/);
        if (match) {
            const num = parseFloat(match[0].replace(/,/g, ''));
            if (!isNaN(num) && num > 0) {
                return num.toLocaleString('en-US', {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                });
            }
        }
        return str;
    }

    /**
     * Format date to Thai format if it's DD/MM/YYYY
     * e.g. "15/03/2568" → "15 มี.ค. 2568"
     */
    _formatThaiDate(dateStr) {
        if (!dateStr) return '';
        
        const thaiMonths = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
                            'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

        // Match DD/MM/YYYY
        const match = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (match) {
            const d = parseInt(match[1], 10);
            const m = parseInt(match[2], 10);
            const y = match[3];
            if (m >= 1 && m <= 12) {
                return `${d} ${thaiMonths[m]} ${y}`;
            }
        }

        // Already in Thai format or other format — keep as-is
        return dateStr;
    }
}
