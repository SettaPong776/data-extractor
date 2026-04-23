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
     * Strategy: Each form has exactly 2 tables (Table 6 + Table 7)
     * So we pair tables and find text sections between them
     */
    async extractEGP(file, onProgress) {
        if (onProgress) onProgress(1, 3);

        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;

        if (onProgress) onProgress(2, 3);

        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        // Parse ALL tables from the document
        const htmlTables = doc.querySelectorAll('table');
        const allTables = [];
        htmlTables.forEach(table => {
            // Skip nested tables
            if (table.parentElement && table.parentElement.closest('table')) return;
            const td = this._parseHTMLTable(table);
            if (td && (td.rows.length > 0 || td.headers.length > 0)) {
                allTables.push(td);
            }
        });

        // Split HTML by <table> tags to get text BETWEEN tables
        const htmlParts = html.split(/<table[\s>]/i);
        // htmlParts[0] = text before first table
        // htmlParts[1] = table1 content + text after table1
        // We need to extract text after each </table> closing tag

        const textBetweenTables = []; // text chunks: [before_t0, between_t0_t1, between_t1_t2, ...]
        
        // First chunk: everything before first table
        textBetweenTables.push(this._htmlToText(htmlParts[0]));
        
        // Remaining chunks: text AFTER each </table>
        for (let i = 1; i < htmlParts.length; i++) {
            const afterClose = htmlParts[i].split(/<\/table>/i);
            // The text after the closing </table> tag
            const textAfter = afterClose.length > 1 ? afterClose.slice(1).join('') : '';
            textBetweenTables.push(this._htmlToText(textAfter));
        }

        console.log(`[e-GP DOCX] Found ${allTables.length} tables, ${textBetweenTables.length} text chunks`);

        // Auto-detect Non-eGP table format (e.g., สขร.1)
        const isNonEGP = allTables.some(t => {
            const firstFewRowsText = t.rows.slice(0, 3).map(r => r.join(' ')).join(' ');
            return /(ชื่อผู้ประกอบการ|ชื่อผู้ประกอบกำร)/.test(firstFewRowsText) && 
                   /(รายการพัสดุ|รำยกำรพัสดุ)/.test(firstFewRowsText);
        });

        if (isNonEGP) {
            return this._extractNonEGP(allTables, onProgress);
        }

        // Strategy: Pair tables as (table6, table7) for each form
        const numForms = Math.floor(allTables.length / 2);
        console.log(`[e-GP DOCX] Detected ${numForms} forms`);

        const egpRows = [];

        for (let fi = 0; fi < numForms; fi++) {
            const t6 = allTables[fi * 2];     // Table 6 (bidders)
            const t7 = allTables[fi * 2 + 1]; // Table 7 (winners)

            // Text before table 6 = textBetweenTables[fi * 2]
            // Text between table 6 and table 7 = textBetweenTables[fi * 2 + 1]
            const sectionText = textBetweenTables[fi * 2] || '';
            const fullText = sectionText.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
            if (fi === 0) console.log(`[e-GP DOCX] Form 1 text:`, fullText.substring(0, 200) + '...');

            // Extract by splitting the full text into numbered sections
            // This correctly handles cases where mammoth.js merges everything into a single line
            const sectionChunks = fullText.split(/(?=(?:^|\s)[1-7]\s*\.\s*)/);
            
            const sections = {};
            for (const chunk of sectionChunks) {
                const m = chunk.trim().match(/^([1-7])\s*\.\s*(.*)/s);
                if (m) {
                    sections[m[1]] = m[2].trim();
                }
            }

            if (fi === 0) console.log(`[e-GP DOCX] Form 1 sections:`, JSON.stringify(sections));

            // Section 3: Project Name (Yellow)
            let projName = sections[3] || '';
            projName = projName.replace(/^.*?โครงการ\s*/, '').trim();
            if (!projName) {
                // Try to match "งานที่จัดซื้อหรือจัดจ้าง" or "ชื่อโครงการ"
                const pMatch = fullText.match(/(?:ชื่อโครงการ|งานที่จัดซื้อหรือจัดจ้าง)\s*[:：]?\s*(.*?)(?=\s*(?:วงเงิน|งบประมาณ|ราคากลาง|วิธี|เลขที่)|$)/);
                if (pMatch) projName = pMatch[1].trim();
            }
            
            let method = '';
            const mm = projName.match(/\s*โดยวิธี(.*?)$/);
            if (mm) {
                method = mm[1].trim();
                projName = projName.replace(/\s*โดยวิธี.*$/, '').trim();
            }
            
            // Remove "ซื้อ" or "จ้าง" at the beginning to match user's exact example
            projName = projName.replace(/^(ซื้อ|จ้าง)\s*/, '').trim();

            // Function to safely extract money amounts, ignoring e-GP IDs (10+ digits) and years (256x)
            const extractMoney = (text, keywordRegex) => {
                const matches = [...text.matchAll(new RegExp(keywordRegex.source + '[^\\d]{0,100}?([\\d,]+\\.\\d{2}|[\\d,]+)', 'g'))];
                for (const m of matches) {
                    const val = m[1].replace(/,/g, '');
                    // Ignore years and e-GP IDs
                    if (val.length < 10 && val !== '2568' && val !== '2569') {
                        return m[1];
                    }
                }
                return '';
            };

            // Section 4: Budget (Dark Green)
            let budget = '';
            if (sections[4]) {
                const bMatch = sections[4].match(/[\d,]+\.\d{2}|[\d,]+/);
                if (bMatch) budget = bMatch[0];
            }
            if (!budget) {
                budget = extractMoney(fullText, /(?:งบประมาณ|วงเงิน)/);
            }

            // Section 5: Median Price (Light Green)
            let medianPrice = '';
            if (sections[5]) {
                const mMatch = sections[5].match(/[\d,]+\.\d{2}|[\d,]+/);
                if (mMatch) medianPrice = mMatch[0];
            }
            if (!medianPrice) {
                medianPrice = extractMoney(fullText, /ราคากลาง/);
            }

            // Table 6: Bidders (Light Blue)
            let biddersStr = '-';
            let blob = '';
            if (t6 && t6.rows.length > 0) {
                const bidders = t6.rows.map(r => {
                    const name = r.length >= 2 ? r[r.length - 2] : '';
                    const price = r.length >= 1 ? r[r.length - 1] : '';
                    return `${name}/ ${price} บาท`.trim();
                }).filter(b => b.length > 5);
                if (bidders.length > 0) biddersStr = bidders.join('\n');
                
                // Keep the raw text for fallback parsing
                blob = t6.rows.map(r => r.join(' ')).join(' ');
            }
            
            // --- NEW: Table 6 / Layout Table Fallback Strategy ---
            
            // 1. Extract Project Name from Table 6
            if (!projName || projName === '(ไม่พบชื่อโครงการ)' || projName.length > 100 || /^\d+$/.test(projName)) {
                // Case A: Table parsed into proper columns
                if (t6 && t6.rows.length > 0) {
                    const firstRow = t6.rows[0];
                    const taxIdIndex = firstRow.findIndex(c => /\d{13}/.test(c));
                    if (taxIdIndex >= 1) {
                        projName = firstRow[taxIdIndex - 1].trim();
                        projName = projName.replace(/^\d+\s+/, '').trim();
                    } else if (firstRow.length >= 4) {
                        // If no tax ID found but we have enough columns, column 1 is usually the project name
                        const possibleName = firstRow.length >= 5 ? firstRow[1] : firstRow[0];
                        if (possibleName && !/^\d+$/.test(possibleName)) {
                            projName = possibleName.replace(/^\d+\s+/, '').trim();
                        }
                    }
                }
                
                // Case B: Table mashed into a blob
                if (!projName || projName === '(ไม่พบชื่อโครงการ)') {
                    if (blob.length > 10) {
                        let pMatch = blob.match(/ราคาที่เสนอ\s*\d+\s*(.*?)(?=\d{13})/);
                        if (!pMatch) {
                            pMatch = blob.match(/(?:^\d+\s*|\s+\d+\s*)(\S.*?)(?=\d{13})/);
                        }
                        if (pMatch) {
                            projName = pMatch[1].replace(/^\d+\s*/, '').trim();
                        }
                    }
                }
            }

            // 2. Extract Budget, Median, and clean Bidder from the proposed price
            // User requested: "ดูเลขจาก column รายชื่อผู้เสนอราคา... ให้ดึงแต่ตัวเลขเงินมา เช่น 8,198.00 บาท จะแสดง ใน วงเงิน... กับ ราคากลาง"
            let proposedPrice = '';
            
            // Try from the clean biddersStr first
            if (biddersStr && biddersStr !== '-') {
                const priceMatch = biddersStr.match(/([\d,]+\.\d{2})/);
                if (priceMatch) proposedPrice = priceMatch[1];
            }
            
            // Try from the blob if it's a huge mashed string
            if (!proposedPrice && blob.length > 20) {
                const bidderMatch = blob.match(/\d{13}\s*(.*?)\s*([\d,]+\.\d{2})/);
                if (bidderMatch) {
                    proposedPrice = bidderMatch[2]; // e.g., "8,198.00"
                    // Clean up biddersStr
                    if (biddersStr.length > 150) {
                        biddersStr = `${bidderMatch[1].trim()}/ ${proposedPrice} บาท`;
                    }
                }
            }

            // Apply the proposed price to Budget and Median Price ALWAYS
            if (proposedPrice) {
                budget = proposedPrice;
                medianPrice = proposedPrice;
            }

            // Table 7: Winners
            let winnersStr = '-';
            let reason = '';
            let contractId = '';
            let contractDate = '';

            if (t7 && t7.rows.length > 0) {
                const dataRows = t7.rows.filter(r => r.some(c => /\d{10,}/.test(c)));
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
                    
                    // Try to extract the PO/Contract Number (e.g. 593/2569) which usually follows the e-GP number
                    const contractMatch = rowStr.match(/6\d{11}\s*([ก-ฮA-Za-z0-9.-]+\/\d{4})/);
                    if (contractMatch) {
                        contractId = contractMatch[1];
                    } else {
                        // Fallback: look for any pattern like XX/25XX that is not a date
                        const fallbackMatch = rowStr.match(/(?:^|\s)(?!(?:\d{1,2}\/\d{1,2}\/\d{4}))([ก-ฮA-Za-z0-9.-]+\/\d{4})/);
                        if (fallbackMatch) contractId = fallbackMatch[1];
                    }
                    
                    const dateMatch = rowStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                    if (dateMatch) {
                        const d = parseInt(dateMatch[1], 10);
                        const m = parseInt(dateMatch[2], 10);
                        const y = dateMatch[3];
                        const thaiMonths = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
                        if (m >= 1 && m <= 12) {
                            contractDate = `${d} ${thaiMonths[m]} ${y}`;
                        } else {
                            contractDate = dateMatch[0];
                        }
                    } else {
                        contractDate = '';
                    }
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
        console.log(`[e-GP DOCX] Extracted ${egpRows.length} rows`);

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

    /**
     * Strip HTML tags and return plain text, preserving paragraph/br newlines
     */
    _htmlToText(html) {
        if (!html) return '';
        // Replace closing block tags with newlines so textContent doesn't mush text together
        const withNewlines = html.replace(/<\/p>|<br\s*\/?>|<\/div>|<\/li>/gi, '\n');
        const tmp = document.createElement('div');
        tmp.innerHTML = withNewlines;
        return tmp.textContent || tmp.innerText || '';
    }

    /**
     * Extract data from Non-eGP summary tables (e.g. สขร.1)
     */
    _extractNonEGP(tables, onProgress) {
        console.log(`[Non-eGP DOCX] Extracting from ${tables.length} tables`);
        const egpRows = [];
        let seq = 1;
        let lastCompanyName = "ไม่ระบุ";

        for (const table of tables) {
            for (const r of table.rows) {
                // If less than 3 columns, it's probably a fully merged separator row
                if (r.length < 3) continue;
                
                const rowStr = r.join(' ');
                // Skip header rows (accounting for TH Sarabun garbled characters)
                if (/(?:ลำดับ|ผู้ประกอบ|รายการ|รำยกำร|จำนวน|จํำนวน|เอกสาร|เหตุผล)/.test(rowStr)) continue;
                
                // Skip sub-header rows that contain literal (1), (2), (3) etc.
                if (r[0] && /^\s*\(\s*[1-9]\s*\)\s*$/.test(r[0])) continue;
                if (r[1] && /^\s*\(\s*[1-9]\s*\)\s*$/.test(r[1])) continue;
                
                let name = r[1] ? r[1].trim() : '';
                
                // If name is empty, it might be a vertically merged cell, reuse the last known name
                if (!name) {
                    if (!r[2] || r[2].trim() === '') continue; // Skip if item is also empty
                    name = lastCompanyName;
                } else {
                    lastCompanyName = name;
                }

                const item = r[2] ? r[2].trim() : '';
                let amount = r[3] ? r[3].trim() : '';
                let dateStr = '';
                let refNo = '';
                let reason = '';

                // Mammoth might parse as 7 cols or squash date/refno into col 4
                if (r.length >= 7) {
                    dateStr = r[4] ? r[4].trim() : '';
                    refNo = r[5] ? r[5].trim() : '';
                    reason = r[6] ? r[6].trim() : '';
                } else if (r.length === 6) {
                    const str = r[4] ? r[4].trim() : '';
                    const dMatch = str.match(/(\d{1,2}\s+[ก-ฮ.A-Za-z]+\s+\d{4})/);
                    if (dMatch) {
                        dateStr = dMatch[1];
                        refNo = str.replace(dateStr, '').trim();
                    } else {
                        refNo = str;
                    }
                    reason = r[5] ? r[5].trim() : '';
                }

                // Clean up amount format
                const amtMatch = amount.match(/([\d,]+\.\d{2})/);
                if (amtMatch) amount = amtMatch[1];

                // Format date if it's DD/MM/YYYY, else keep as is (e.g. 2 ส.ค. 2568)
                let contractDate = dateStr;
                const dMatch2 = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                if (dMatch2) {
                    const thaiMonthsReal = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
                    const d = parseInt(dMatch2[1], 10);
                    const m = parseInt(dMatch2[2], 10);
                    const y = dMatch2[3];
                    if (m >= 1 && m <= 12) {
                        contractDate = `${d} ${thaiMonthsReal[m]} ${y}`;
                    }
                }

                const biddersStr = `${name}/ ${amount} บาท`.trim();

                egpRows.push([
                    seq++,
                    item,            // 2. งานที่จัดซื้อหรือจัดจ้าง
                    amount,          // 3. วงเงิน
                    '',              // 4. ราคากลาง (เว้นว่าง)
                    '',              // 5. วิธีซื้อหรือจ้าง (เว้นว่าง)
                    biddersStr,      // 6. รายชื่อผู้เสนอราคา
                    biddersStr,      // 7. ผู้ได้รับการคัดเลือก
                    reason,          // 8. เหตุผล
                    refNo,           // 9. เลขที่สัญญา
                    contractDate     // 10. วันที่ทำสัญญา
                ]);
            }
        }

        if (onProgress) onProgress(3, 3);
        console.log(`[Non-eGP DOCX] Extracted ${egpRows.length} rows`);

        return [{
            headers: [
                "ลำดับที่", "งานที่จัดซื้อหรือจัดจ้าง",
                "วงเงินที่จะซื้อหรือจ้าง", "ราคากลาง", "วิธีซื้อหรือจ้าง",
                "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
                "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
                "เหตุผลที่คัดเลือกโดยสรุป",
                "เลขที่สัญญา/ข้อตกลง", "วันที่ทำสัญญา/ข้อตกลง"
            ],
            rows: egpRows,
            rowCount: egpRows.length,
            columnCount: 10
        }];
    }
}
