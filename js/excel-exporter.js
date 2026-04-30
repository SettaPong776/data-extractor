/**
 * ExcelExporter — ส่งออกข้อมูลเป็นไฟล์ .xlsx โดยใช้ SheetJS
 * ใช้ Blob + Anchor tag แทน XLSX.writeFile() เพื่อให้ชื่อไฟล์ถูกต้อง
 */
class ExcelExporter {
    /**
     * Export data to Excel file and trigger download
     * @param {Object} data — { headers: string[], rows: string[][] }
     * @param {string} filename — ชื่อไฟล์ (ไม่ต้องมีนามสกุล)
     * @param {Object|null} templateConfig — { type, title, subtitle }
     */
    export(data, filename = 'extracted_data', templateConfig = null) {
        const wb = XLSX.utils.book_new();

        let aoa = [];
        if (templateConfig && templateConfig.type === 'procurement') {
            // Row 1: Title (place in column E / index 4 for centering approx)
            const titleRow = [];
            titleRow[4] = templateConfig.title;
            aoa.push(titleRow);
            
            // Row 2: Subtitle
            const subTitleRow = [];
            subTitleRow[4] = templateConfig.subtitle;
            aoa.push(subTitleRow);
            
            // Row 3: headers and data
            aoa.push(data.headers);
            aoa.push(...data.rows);
        } else {
            // Build array of arrays: [headers, ...rows]
            aoa = [data.headers, ...data.rows];
        }

        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Auto-size columns
        const colWidths = data.headers.map((h, i) => {
            let maxLen = h ? h.length : 10;
            data.rows.forEach(row => {
                const cellLen = (row[i] || '').toString().length;
                if (cellLen > maxLen) maxLen = cellLen;
            });
            return { wch: Math.min(Math.max(maxLen + 2, 8), 50) };
        });
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, 'Data');

        // Trigger download using Blob + Anchor (reliable filename)
        this._downloadWorkbook(wb, filename);
    }

    /**
     * Export multiple tables to separate sheets
     * @param {Array<Object>} tables — array of { name, headers, rows }
     * @param {string} filename
     * @param {Object|null} templateConfig
     */
    exportMultiple(tables, filename = 'extracted_data', templateConfig = null) {
        const wb = XLSX.utils.book_new();

        tables.forEach((table, idx) => {
            let aoa = [];
            if (templateConfig && templateConfig.type === 'procurement') {
                const titleRow = [];
                titleRow[4] = templateConfig.title;
                aoa.push(titleRow);
                
                const subTitleRow = [];
                subTitleRow[4] = templateConfig.subtitle;
                aoa.push(subTitleRow);
                
                aoa.push(table.headers);
                aoa.push(...table.rows);
            } else {
                aoa = [table.headers, ...table.rows];
            }
            const ws = XLSX.utils.aoa_to_sheet(aoa);

            // Auto-size
            const colWidths = table.headers.map((h, i) => {
                let maxLen = h ? h.length : 10;
                table.rows.forEach(row => {
                    const cellLen = (row[i] || '').toString().length;
                    if (cellLen > maxLen) maxLen = cellLen;
                });
                return { wch: Math.min(Math.max(maxLen + 2, 8), 50) };
            });
            ws['!cols'] = colWidths;

            const sheetName = table.name || `Table ${idx + 1}`;
            // SheetJS sheet name max 31 chars
            XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        });

        this._downloadWorkbook(wb, filename);
    }

    /**
     * Download workbook as .xlsx using Blob + Anchor tag
     * This ensures the filename and extension are always correct
     */
    _downloadWorkbook(wb, filename) {
        // Generate Excel binary data
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        
        // Create Blob with correct MIME type
        const blob = new Blob([wbout], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        // Create download link
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${filename}.xlsx`;
        a.style.display = 'none';
        
        // Trigger download
        document.body.appendChild(a);
        a.click();
        
        // Cleanup
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 200);
    }
}
