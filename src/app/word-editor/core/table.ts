/** Thao tác bảng có tính tới rowspan/colspan, dựa trên bản đồ lưới. */

type CellEl = HTMLTableCellElement;

export interface Grid {
  cells: (CellEl | null)[][];
  rows: HTMLTableRowElement[];
  cols: number;
}

export function buildGrid(table: HTMLTableElement): Grid {
  // Bỏ qua hàng đệm tạm của bộ chia trang.
  const rows = Array.from(table.rows).filter(r => !r.classList.contains('we-pb-spacer'));
  const cells: (CellEl | null)[][] = rows.map(() => []);
  rows.forEach((row, r) => {
    let c = 0;
    for (const cell of Array.from(row.cells)) {
      while (cells[r][c]) c++;
      for (let dr = 0; dr < Math.max(1, cell.rowSpan); dr++) {
        if (!cells[r + dr]) continue;
        for (let dc = 0; dc < Math.max(1, cell.colSpan); dc++) cells[r + dr][c + dc] = cell;
      }
      c += Math.max(1, cell.colSpan);
    }
  });
  const cols = Math.max(0, ...cells.map(r => r.length));
  for (const r of cells) for (let c = 0; c < cols; c++) r[c] ??= null;
  return { cells, rows, cols };
}

export function cellPos(g: Grid, cell: CellEl) {
  for (let r = 0; r < g.cells.length; r++) {
    const c = g.cells[r].indexOf(cell);
    if (c >= 0) return { row: r, col: c, lastRow: r + Math.max(1, cell.rowSpan) - 1, lastCol: c + Math.max(1, cell.colSpan) - 1 };
  }
  return { row: 0, col: 0, lastRow: 0, lastCol: 0 };
}

function newCell(like: CellEl | null): CellEl {
  const cell = (like ? like.cloneNode(false) : document.createElement('td')) as CellEl;
  cell.removeAttribute('rowspan');
  cell.removeAttribute('colspan');
  cell.removeAttribute('data-we-pb');
  cell.classList.remove('we-cell-sel');
  cell.innerHTML = '<p><br></p>';
  return cell;
}

/** Chèn ô vào hàng `r` sao cho nó nằm ở cột lưới `col`. */
function insertCellAt(g: Grid, r: number, col: number, cell: CellEl) {
  const row = g.rows[r];
  const after = Array.from(row.cells).find(c => cellPos(g, c).col > col);
  row.insertBefore(cell, after ?? null);
}

export function insertRow(table: HTMLTableElement, ref: CellEl, below: boolean) {
  const g = buildGrid(table);
  const p = cellPos(g, ref);
  const at = below ? p.lastRow : p.row; // chèn sau hàng này (below) hoặc trước hàng này
  const row = document.createElement('tr');
  const refRow = g.rows[at];
  for (const a of Array.from(refRow.attributes)) if (a.name !== 'data-we-pb') row.setAttribute(a.name, a.value);
  const grown = new Set<CellEl>();

  for (let c = 0; c < g.cols; c++) {
    const cell = g.cells[at][c];
    if (!cell) continue;
    const cp = cellPos(g, cell);
    const crosses = below ? cp.lastRow > at : cp.row < at;
    if (crosses) {
      if (!grown.has(cell)) {
        cell.rowSpan = Math.max(1, cell.rowSpan) + 1;
        grown.add(cell);
      }
      c = cp.lastCol;
      continue;
    }
    const nc = newCell(cell);
    if (cell.colSpan > 1) nc.colSpan = cell.colSpan;
    row.append(nc);
    c = cp.lastCol;
  }
  refRow.parentNode!.insertBefore(row, below ? refRow.nextSibling : refRow);
  return row;
}

export function insertColumn(table: HTMLTableElement, ref: CellEl, right: boolean) {
  const g = buildGrid(table);
  const p = cellPos(g, ref);
  const at = right ? p.lastCol : p.col;
  const done = new Set<CellEl>();

  for (let r = 0; r < g.rows.length; r++) {
    const cell = g.cells[r][at];
    if (!cell) {
      insertCellAt(g, r, at, newCell(null));
      continue;
    }
    if (done.has(cell)) continue;
    const cp = cellPos(g, cell);
    const crosses = right ? cp.lastCol > at : cp.col < at;
    if (crosses) {
      cell.colSpan = Math.max(1, cell.colSpan) + 1;
      done.add(cell);
      continue;
    }
    if (cp.row !== r) continue; // ô rowspan từ hàng trên: ô mới đã được thêm ở hàng của nó
    const nc = newCell(cell);
    if (cell.rowSpan > 1) nc.rowSpan = cell.rowSpan;
    cell.parentNode!.insertBefore(nc, right ? cell.nextSibling : cell);
    done.add(cell);
  }
  const colgroup = table.querySelector('colgroup');
  if (colgroup) colgroup.remove();
}

export function deleteRow(table: HTMLTableElement, ref: CellEl) {
  const g = buildGrid(table);
  const p = cellPos(g, ref);
  for (let r = p.lastRow; r >= p.row; r--) removeRowAt(table, r);
  if (!table.rows.length) table.remove();
}

function removeRowAt(table: HTMLTableElement, r: number) {
  const g = buildGrid(table);
  const seen = new Set<CellEl>();
  for (let c = 0; c < g.cols; c++) {
    const cell = g.cells[r][c];
    if (!cell || seen.has(cell)) continue;
    seen.add(cell);
    const cp = cellPos(g, cell);
    if (cell.rowSpan > 1) {
      if (cp.row === r && g.rows[r + 1]) {
        cell.rowSpan -= 1;
        insertCellAt(g, r + 1, cp.col, cell);
      } else if (cp.row < r) {
        cell.rowSpan -= 1;
      }
    }
  }
  g.rows[r].remove();
}

export function deleteColumn(table: HTMLTableElement, ref: CellEl) {
  const g = buildGrid(table);
  const p = cellPos(g, ref);
  for (let c = p.lastCol; c >= p.col; c--) {
    const grid = buildGrid(table);
    const seen = new Set<CellEl>();
    for (let r = 0; r < grid.rows.length; r++) {
      const cell = grid.cells[r][c];
      if (!cell || seen.has(cell)) continue;
      seen.add(cell);
      if (cell.colSpan > 1) cell.colSpan -= 1;
      else cell.remove();
    }
  }
  table.querySelector('colgroup')?.remove();
  for (const row of Array.from(table.rows)) if (!row.cells.length) row.remove();
  if (!table.rows.length || !table.rows[0].cells.length) table.remove();
}

/** Các ô trong hình chữ nhật bao quanh a và b (mở rộng để chứa trọn ô gộp). */
export function cellsInRect(table: HTMLTableElement, a: CellEl, b: CellEl): CellEl[] {
  const g = buildGrid(table);
  const pa = cellPos(g, a);
  const pb = cellPos(g, b);
  let r0 = Math.min(pa.row, pb.row), r1 = Math.max(pa.lastRow, pb.lastRow);
  let c0 = Math.min(pa.col, pb.col), c1 = Math.max(pa.lastCol, pb.lastCol);
  for (let changed = true; changed; ) {
    changed = false;
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const cell = g.cells[r][c];
        if (!cell) continue;
        const p = cellPos(g, cell);
        if (p.row < r0) (r0 = p.row), (changed = true);
        if (p.lastRow > r1) (r1 = p.lastRow), (changed = true);
        if (p.col < c0) (c0 = p.col), (changed = true);
        if (p.lastCol > c1) (c1 = p.lastCol), (changed = true);
      }
    }
  }
  const out = new Set<CellEl>();
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (g.cells[r][c]) out.add(g.cells[r][c]!);
  return Array.from(out);
}

export function mergeCells(table: HTMLTableElement, cells: CellEl[]): CellEl | null {
  if (cells.length < 2) return null;
  const g = buildGrid(table);
  const pos = cells.map(c => cellPos(g, c));
  const r0 = Math.min(...pos.map(p => p.row)), r1 = Math.max(...pos.map(p => p.lastRow));
  const c0 = Math.min(...pos.map(p => p.col)), c1 = Math.max(...pos.map(p => p.lastCol));
  const target = g.cells[r0][c0]!;
  for (const cell of cells) {
    if (cell === target) continue;
    const empty = !cell.textContent?.trim() && !cell.querySelector('img, table');
    if (!empty) target.append(...Array.from(cell.childNodes));
    cell.remove();
  }
  target.rowSpan = r1 - r0 + 1;
  target.colSpan = c1 - c0 + 1;
  target.style.removeProperty('width');
  // Hàng trống sau khi gộp (mọi ô đã bị gộp lên trên) phải xoá và giảm rowspan tương ứng.
  for (const row of Array.from(table.rows)) {
    if (row.cells.length) continue;
    const gg = buildGrid(table);
    const r = gg.rows.indexOf(row);
    const seen = new Set<CellEl>();
    for (const cell of gg.cells[r] ?? []) {
      if (cell && !seen.has(cell)) {
        seen.add(cell);
        cell.rowSpan = Math.max(1, cell.rowSpan - 1);
      }
    }
    row.remove();
  }
  return target;
}

export function splitCell(table: HTMLTableElement, cell: CellEl) {
  if (cell.rowSpan <= 1 && cell.colSpan <= 1) return;
  const g = buildGrid(table);
  const p = cellPos(g, cell);
  const rs = cell.rowSpan, cs = cell.colSpan;
  cell.rowSpan = 1;
  cell.colSpan = 1;
  for (let r = p.row; r < p.row + rs; r++) {
    for (let c = p.col + (r === p.row ? 1 : 0); c < p.col + cs; c++) {
      const grid = buildGrid(table);
      insertCellAt(grid, r, c - 0.5, newCell(cell));
    }
  }
}

/** Chia đều độ rộng cột, bảng rộng 100%. */
export function distributeColumns(table: HTMLTableElement) {
  const g = buildGrid(table);
  table.querySelector('colgroup')?.remove();
  table.style.width = '100%';
  table.style.tableLayout = 'fixed';
  for (const row of g.rows) {
    for (const cell of Array.from(row.cells)) cell.style.width = `${((100 / g.cols) * Math.max(1, cell.colSpan)).toFixed(2)}%`;
  }
}

/** Chuẩn hoá độ rộng cột thành <colgroup> px (dùng khi kéo đổi độ rộng cột). */
export function columnWidths(table: HTMLTableElement, scale: number): number[] {
  const g = buildGrid(table);
  const widths = new Array<number>(g.cols).fill(0);
  for (let c = 0; c < g.cols; c++) {
    for (let r = 0; r < g.rows.length; r++) {
      const cell = g.cells[r][c];
      if (cell && Math.max(1, cell.colSpan) === 1) {
        widths[c] = cell.getBoundingClientRect().width / scale;
        break;
      }
    }
  }
  const known = widths.filter(w => w > 0);
  const fallback = known.length ? known.reduce((a, b) => a + b, 0) / known.length : 100;
  return widths.map(w => w || fallback);
}

export function applyColumnWidths(table: HTMLTableElement, widths: number[]) {
  let colgroup = table.querySelector('colgroup');
  if (!colgroup) {
    colgroup = document.createElement('colgroup');
    table.prepend(colgroup);
  }
  colgroup.replaceChildren(
    ...widths.map(w => {
      const col = document.createElement('col');
      col.style.width = `${Math.round(w * 100) / 100}px`;
      return col;
    }),
  );
  for (const row of Array.from(table.rows)) for (const cell of Array.from(row.cells)) cell.style.removeProperty('width');
  table.style.tableLayout = 'fixed';
  table.style.width = `${Math.round(widths.reduce((a, b) => a + b, 0))}px`;
}

export function createTableHtml(rows: number, cols: number, contentWidth: number) {
  const w = Math.floor(contentWidth / cols);
  const cell = `<td style="border: 1px solid #000; width: ${w}px"><p><br></p></td>`;
  const row = `<tr>${cell.repeat(cols)}</tr>`;
  return `<table style="border-collapse: collapse; width: 100%; table-layout: fixed"><tbody>${row.repeat(rows)}</tbody></table><p><br></p>`;
}
