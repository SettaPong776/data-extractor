/**
 * ColumnMapper — UI สำหรับจัดเรียง/ตั้งชื่อ/ซ่อนคอลัมน์
 * รองรับ Drag & Drop เพื่อจัดลำดับ
 */
class ColumnMapper {
    constructor(containerId) {
        this.container = document.getElementById(containerId);
        this.mapping = []; // { sourceIndex, sourceName, targetName, enabled, order }
        this.onChangeCallback = null;
        this._dragItem = null;
    }

    /**
     * Initialize mapping from extracted table headers
     * @param {Array<string>} headers
     */
    setHeaders(headers) {
        this.mapping = headers.map((h, i) => ({
            sourceIndex: i,
            sourceName: h || `คอลัมน์ ${i + 1}`,
            targetName: h || `คอลัมน์ ${i + 1}`,
            enabled: true,
            order: i
        }));
        this.render();
    }

    /**
     * Set callback for when mapping changes
     */
    onChange(callback) {
        this.onChangeCallback = callback;
    }

    /**
     * Get current mapping (only enabled, in order)
     * @returns {Array} sorted enabled mappings
     */
    getMapping() {
        return [...this.mapping]
            .sort((a, b) => a.order - b.order)
            .filter(m => m.enabled);
    }

    /**
     * Get all mappings (including disabled)
     */
    getAllMapping() {
        return [...this.mapping].sort((a, b) => a.order - b.order);
    }

    /**
     * Apply mapping to data rows
     * @param {Array<Array>} rows — original data rows
     * @returns {Object} { headers, rows }
     */
    applyMapping(rows) {
        const mapping = this.getMapping();
        const newHeaders = mapping.map(m => m.targetName);
        const newRows = rows.map(row =>
            mapping.map(m => row[m.sourceIndex] || '')
        );
        return { headers: newHeaders, rows: newRows };
    }

    /**
     * Render the mapper UI
     */
    render() {
        this.container.innerHTML = '';

        const sorted = [...this.mapping].sort((a, b) => a.order - b.order);

        sorted.forEach((col, visualIdx) => {
            const item = document.createElement('div');
            item.className = 'mapper-item' + (col.enabled ? '' : ' disabled');
            item.draggable = true;
            item.dataset.sourceIndex = col.sourceIndex;

            item.innerHTML = `
                <span class="mapper-drag-handle material-icons-round">drag_indicator</span>
                <input type="checkbox" class="mapper-checkbox"
                    ${col.enabled ? 'checked' : ''}
                    data-source="${col.sourceIndex}"
                    title="เปิด/ปิดคอลัมน์นี้">
                <span class="mapper-source">${col.sourceName}</span>
                <span class="mapper-arrow material-icons-round">arrow_forward</span>
                <input type="text" class="mapper-target-input"
                    value="${col.targetName}"
                    data-source="${col.sourceIndex}"
                    placeholder="ชื่อคอลัมน์ใน Excel">
            `;

            // Checkbox toggle
            const checkbox = item.querySelector('.mapper-checkbox');
            checkbox.addEventListener('change', () => {
                const m = this.mapping.find(x => x.sourceIndex === col.sourceIndex);
                if (m) m.enabled = checkbox.checked;
                item.classList.toggle('disabled', !checkbox.checked);
                this._emitChange();
            });

            // Target name input
            const input = item.querySelector('.mapper-target-input');
            input.addEventListener('input', () => {
                const m = this.mapping.find(x => x.sourceIndex === col.sourceIndex);
                if (m) m.targetName = input.value;
                this._emitChange();
            });

            // Drag events
            item.addEventListener('dragstart', (e) => {
                this._dragItem = item;
                item.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', col.sourceIndex);
            });

            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
                this.container.querySelectorAll('.mapper-item').forEach(el => {
                    el.classList.remove('drag-over-item');
                });
                this._dragItem = null;
            });

            item.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                if (this._dragItem && this._dragItem !== item) {
                    item.classList.add('drag-over-item');
                }
            });

            item.addEventListener('dragleave', () => {
                item.classList.remove('drag-over-item');
            });

            item.addEventListener('drop', (e) => {
                e.preventDefault();
                item.classList.remove('drag-over-item');

                if (!this._dragItem || this._dragItem === item) return;

                const dragSourceIdx = parseInt(this._dragItem.dataset.sourceIndex);
                const dropSourceIdx = parseInt(item.dataset.sourceIndex);

                // Swap orders
                const dragMapping = this.mapping.find(m => m.sourceIndex === dragSourceIdx);
                const dropMapping = this.mapping.find(m => m.sourceIndex === dropSourceIdx);

                if (dragMapping && dropMapping) {
                    const tempOrder = dragMapping.order;
                    dragMapping.order = dropMapping.order;
                    dropMapping.order = tempOrder;
                    this.render();
                    this._emitChange();
                }
            });

            this.container.appendChild(item);
        });
    }

    _emitChange() {
        if (this.onChangeCallback) {
            this.onChangeCallback(this.getMapping());
        }
    }
}
