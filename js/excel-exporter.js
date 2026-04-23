/**
 * ExcelExporter — ส่งออกข้อมูลเป็นไฟล์ .xlsx โดยใช้ SheetJS
 */
class ExcelExporter {
    /**
     * Export data to Excel file and trigger download
     * @param {Object} data — { headers: string[], rows: string[][] }
     * @param {string} filename — ชื่อไฟล์ (ไม่ต้องมีนามสกุล)
     */
    export(data, filename = 'extracted_data') {
        const wb = XLSX.utils.book_new();

        // Build array of arrays: [headers, ...rows]
        const aoa = [data.headers, ...data.rows];

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

        // Style header row (SheetJS free version has limited styling)
        // But we can set bold via cell properties if available

        XLSX.utils.book_append_sheet(wb, ws, 'Data');

        // Trigger download
        XLSX.writeFile(wb, `${filename}.xlsx`);
    }

    /**
     * Export multiple tables to separate sheets
     * @param {Array<Object>} tables — array of { name, headers, rows }
     * @param {string} filename
     */
    exportMultiple(tables, filename = 'extracted_data') {
        const wb = XLSX.utils.book_new();

        tables.forEach((table, idx) => {
            const aoa = [table.headers, ...table.rows];
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

        XLSX.writeFile(wb, `${filename}.xlsx`);
    }
}
