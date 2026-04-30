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
     * Format number string to currency format with commas and 2 decimals
     */
    _formatCurrency(text) {
        if (!text) return '';
        const match = text.match(/[\d,]+(\.\d+)?/);
        if (match) {
            const num = parseFloat(match[0].replace(/,/g, ''));
            if (!isNaN(num)) {
                return text.replace(match[0], num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
            }
        }
        return text;
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

            // Combine text before table 6 AND between table 6/7 for robust section extraction
            const textBeforeT6 = textBetweenTables[fi * 2] || '';
            const textBetweenT6T7 = textBetweenTables[fi * 2 + 1] || '';
            const fullText = (textBeforeT6 + ' ' + textBetweenT6T7).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
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

            // Function to safely extract money amounts, ignoring e-GP IDs (10+ digits) and years
            const extractMoney = (text, keywordRegex) => {
                // Try to match value with "บาท" first for accuracy
                const patStrBaht = keywordRegex.source + '[^0-9]{0,100}?([0-9][0-9,]*[.][0-9]{1,2}|[0-9][0-9,]*)\\s*บาท';
                let matches = [...text.matchAll(new RegExp(patStrBaht, 'g'))];
                
                if (matches.length === 0) {
                    // Fallback to without "บาท", but require at least 3 digits or a decimal to avoid grabbing standalone "6"
                    const patStr = keywordRegex.source + '[^0-9]{0,100}?([0-9][0-9,]*[.][0-9]{1,2}|[0-9][0-9,]{3,})';
                    matches = [...text.matchAll(new RegExp(patStr, 'g'))];
                }
                
                for (const m of matches) {
                    const val = m[1].replace(/,/g, '');
                    if (val.length < 10 && !/^25[6-7][0-9]$/.test(val)) {
                        return m[1];
                    }
                }
                return '';
            };

            const extractSectionMoney = (text) => {
                if (!text) return '';
                // 1. Try to find a number followed by "บาท"
                const m1 = text.match(/([\d,]+(?:\.\d{1,2})?)\s*บาท/);
                if (m1) return m1[1];
                // 2. Try to find a number with a decimal point (like 52,943.60 or 52943.6)
                const m2 = text.match(/([\d,]+\.\d{1,2})/);
                if (m2) return m2[1];
                // 3. Fallback to any number that is at least 3 digits (to avoid matching "6" in "6 รายการ")
                const m3 = text.match(/([1-9][\d,]{2,})/);
                if (m3) return m3[1];
                // 4. Last resort: any number
                const m4 = text.match(/[\d,]+/);
                return m4 ? m4[0] : '';
            };

            // Section 4: Budget (Dark Green) — ดึงจาก "4. งบประมาณ XXX บาท"
            let budget = extractSectionMoney(sections[4]);
            if (!budget) {
                budget = extractMoney(fullText, /(?:งบประมาณ|วงเงิน)/);
            }

            // Section 5: Median Price (Light Green) — ดึงจาก "5. ราคากลาง XXX บาท"
            let medianPrice = extractSectionMoney(sections[5]);
            if (!medianPrice) {
                medianPrice = extractMoney(fullText, /ราคากลาง/);
            }

            // FALLBACK: When Thai keywords are garbled, use STRUCTURAL approach
            // In e-GP documents, the last two "NUMBER บาท" in text BEFORE table 6
            // are always budget (section 4) and median price (section 5)
            if (!budget || !medianPrice) {
                const preTableText = textBeforeT6.replace(/\n/g, ' ');
                const allMoneyMatches = [];
                const moneyPattern = /(\d[\d,]*(?:\.\d{2})?)\s*บาท/g;
                let mm;
                while ((mm = moneyPattern.exec(preTableText)) !== null) {
                    const val = mm[1].replace(/,/g, '');
                    // Skip years (2565-2579) and e-GP IDs (10+ digits)
                    if (val.length < 10 && !/^25[6-7][0-9]$/.test(val)) {
                        allMoneyMatches.push(mm[1]);
                    }
                }
                if (fi < 3) console.log(`[e-GP DOCX] Form ${fi+1} money-in-text:`, allMoneyMatches);
                if (allMoneyMatches.length >= 2) {
                    // Last two = budget, median price
                    if (!budget) budget = allMoneyMatches[allMoneyMatches.length - 2];
                    if (!medianPrice) medianPrice = allMoneyMatches[allMoneyMatches.length - 1];
                } else if (allMoneyMatches.length === 1) {
                    if (!budget) budget = allMoneyMatches[0];
                    if (!medianPrice) medianPrice = allMoneyMatches[0];
                }
            }

            if (fi < 3 || !budget) {
                console.log(`[e-GP DOCX] Form ${fi+1}: budget="${budget}", median="${medianPrice}"`);
            }

            budget = this._formatCurrency(budget);
            medianPrice = this._formatCurrency(medianPrice);

            // Table 6: Bidders (Light Blue)
            let biddersStr = '-';
            let blob = '';
            let biddersList = []; // structured data for cross-referencing with Table 7
            if (t6 && t6.rows.length > 0) {
                if (fi < 3) console.log(`[e-GP DOCX] Form ${fi+1} T6: ${t6.rows.length} rows`, t6.rows.map(r => r.join(' | ')));
                const bidders = t6.rows.map(r => {
                    // Smart detect: find company name and price by content pattern
                    let bName = '';
                    let bPrice = '';
                    
                    // Scan for company name (left to right)
                    for (const cell of r) {
                        const c = (cell || '').trim();
                        if (!c) continue;
                        if (!bName && /บริษัท|ห้าง|ร้าน|สหกรณ์|หจก|นาย|นาง|กิจการ/.test(c)) {
                            bName = c;
                            break;
                        }
                    }
                    
                    // Scan for price from RIGHT to LEFT (price is always last numeric column)
                    for (let ci = r.length - 1; ci >= 0; ci--) {
                        const c = (r[ci] || '').trim();
                        if (!c) continue;
                        // Skip tax IDs and e-GP IDs (10+ digits)
                        const digitsOnly = c.replace(/[,.\s]/g, '');
                        if (/^\d+$/.test(digitsOnly) && digitsOnly.length >= 10) continue;
                        // Match decimal price like 732,499.00 or 52943.6
                        const pm = c.match(/([\d,]+\.\d{1,2})/);
                        if (pm) { bPrice = pm[1]; break; }
                        // Match integer price like 770000 (at least 3 digits, not a tax ID)
                        if (/^[\d,]+$/.test(c) && digitsOnly.length >= 3 && digitsOnly.length < 10) {
                            bPrice = c; break;
                        }
                    }
                    
                    // Fallback: last 2 columns
                    if (!bName && r.length >= 2) {
                        const cand = r[r.length - 2].trim();
                        if (cand && !/^\d+$/.test(cand) && !/^\d{13}$/.test(cand)) bName = cand;
                    }
                    if (!bPrice && r.length >= 1) {
                        const cand = r[r.length - 1].trim();
                        const dOnly = cand.replace(/[,.\s]/g, '');
                        if (dOnly.length < 10) {
                            const pm = cand.match(/([\d,]+\.\d{1,2})/);
                            if (pm) bPrice = pm[1];
                            else if (/^[\d,]+$/.test(cand) && dOnly.length >= 3) bPrice = cand;
                        }
                    }
                    if (bName && bPrice && /\d/.test(bPrice)) {
                        bPrice = this._formatCurrency(bPrice);
                        biddersList.push({ name: bName, price: bPrice });
                        return `${bName} / ${bPrice} บาท`;
                    }
                    return null;
                }).filter(b => b !== null);
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

            // 2. Fallback: use bidder price for budget/median ONLY when sections 4/5 had no data
            let proposedPrice = '';
            if (biddersStr && biddersStr !== '-') {
                const priceMatch = biddersStr.match(/([\d,]+\.\d{1,2})/);
                if (priceMatch) proposedPrice = priceMatch[1];
            }
            if (!proposedPrice && blob.length > 20) {
                const bidderMatch = blob.match(/\d{13}\s*(.*?)\s*([\d,]+\.\d{1,2})/);
                if (bidderMatch) {
                    proposedPrice = bidderMatch[2];
                    if (biddersStr.length > 150) {
                        biddersStr = `${bidderMatch[1].trim()} / ${proposedPrice} บาท`;
                    }
                }
            }
            // Only apply as fallback — never overwrite values from section 4/5
            if (proposedPrice) {
                if (!budget) budget = proposedPrice;
                if (!medianPrice) medianPrice = proposedPrice;
            }

            // Table 7: Winners — cross-reference with biddersList from Table 6
            let winnersStr = '-';
            let reason = '';
            let contractId = '';
            let contractDate = '';

            if (t7 && t7.rows.length > 0) {
                const dataRows = t7.rows.filter(r => r.some(c => /\d{10,}/.test(c)));
                if (dataRows.length > 0) {
                    const r = dataRows[0];
                    const rowStr = r.join(' ');

                    // Find price in Table 7
                    const priceCol = r.findIndex(c => /^[\d,]+\.\d{2}$/.test(c.trim()));
                    let winPrice = priceCol >= 0 ? r[priceCol].trim() : '';

                    // Cross-reference: match Table 7 price against bidders from Table 6
                    let winName = '';
                    if (winPrice && biddersList.length > 1) {
                        const normWin = winPrice.replace(/,/g, '');
                        const matched = biddersList.find(b => b.price.replace(/,/g, '') === normWin);
                        if (matched) winName = matched.name;
                    }
                    // Fallback: find name directly from Table 7
                    if (!winName) {
                        const nameCol = r.findIndex(c => /บริษัท|ห้าง|ร้าน|สหกรณ์/.test(c));
                        winName = nameCol >= 0 ? r[nameCol] : (r.length > 2 ? r[2] : '');
                    }

                    if (winPrice) winPrice = this._formatCurrency(winPrice);
                    winnersStr = winPrice ? `${winName} / ${winPrice} บาท`.trim() : (winName || '-');
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
                amount = this._formatCurrency(amount);

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
                    amount,          // 4. ราคากลาง (ใช้ amount เหมือนวงเงิน)
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
