export const PAGE_BREAK_HTML = '<div class="we-page-break" contenteditable="false"></div>';

export const BLOCK_TAGS = new Set([
  'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'DIV', 'BLOCKQUOTE', 'PRE', 'TD', 'TH', 'UL', 'OL',
  'TABLE', 'TBODY', 'THEAD', 'TFOOT', 'TR', 'HR', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'FIGURE',
]);

/** Khối "lá": nơi chứa chữ, mà lệnh định dạng đoạn tác động lên. */
export const TEXT_BLOCKS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'DIV', 'TD', 'TH']);

export const isElement = (n: Node | null | undefined): n is HTMLElement => !!n && n.nodeType === 1;

export function closestBlock(node: Node | null, root: HTMLElement): HTMLElement | null {
  let n: Node | null = node;
  while (n && n !== root) {
    if (isElement(n) && TEXT_BLOCKS.has(n.tagName)) return n;
    n = n.parentNode;
  }
  return null;
}

export function closestTag<K extends keyof HTMLElementTagNameMap>(
  node: Node | null,
  tags: K[],
  root: HTMLElement,
): HTMLElementTagNameMap[K] | null {
  let n: Node | null = node;
  const set = new Set(tags.map(t => t.toUpperCase()));
  while (n && n !== root) {
    if (isElement(n) && set.has(n.tagName)) return n as HTMLElementTagNameMap[K];
    n = n.parentNode;
  }
  return null;
}

/** Các khối chữ mà vùng chọn chạm tới, theo thứ tự trong tài liệu. */
export function selectedBlocks(range: Range, root: HTMLElement): HTMLElement[] {
  const start = closestBlock(range.startContainer, root);
  if (range.collapsed) return start ? [start] : [];

  const out = new Set<HTMLElement>();
  if (start) out.add(start);
  const container = range.commonAncestorContainer;
  const walker = document.createTreeWalker(
    isElement(container) ? container : container.parentNode!,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
  );
  for (let n: Node | null = walker.currentNode; n; n = walker.nextNode()) {
    if (!range.intersectsNode(n)) continue;
    if (n.nodeType === 3 || (isElement(n) && (n.tagName === 'BR' || n.tagName === 'IMG'))) {
      const b = closestBlock(n, root);
      if (b) out.add(b);
    }
  }
  // Bỏ khối cha nếu đã có khối con (vd: td chứa p).
  const list = Array.from(out);
  return list.filter(b => !list.some(o => o !== b && b.contains(o)));
}

/** Các text node nằm trong vùng chọn; tách node ở hai đầu để chỉ lấy đúng phần được chọn. */
export function textNodesInRange(range: Range): Text[] {
  if (range.collapsed) return [];
  let { startContainer: sc, startOffset: so, endContainer: ec, endOffset: eo } = range;
  if (sc.nodeType === 3 && so > 0 && so < (sc as Text).length) {
    const tail = (sc as Text).splitText(so);
    if (ec === sc) {
      ec = tail;
      eo -= so;
    }
    sc = tail;
    so = 0;
  }
  if (ec.nodeType === 3 && eo > 0 && eo < (ec as Text).length) (ec as Text).splitText(eo);

  const r = document.createRange();
  r.setStart(sc, so);
  r.setEnd(ec, ec.nodeType === 3 ? Math.min(eo, (ec as Text).length) : eo);
  const root = r.commonAncestorContainer;
  const out: Text[] = [];
  if (root.nodeType === 3) return [root as Text];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (r.intersectsNode(n) && (n as Text).length) out.push(n as Text);
  }
  const sel = getSelection();
  if (out.length && sel) {
    const nr = document.createRange();
    nr.setStart(out[0], 0);
    nr.setEnd(out.at(-1)!, out.at(-1)!.length);
    sel.removeAllRanges();
    sel.addRange(nr);
  }
  return out;
}

// ---------- Lưu / khôi phục vùng chọn theo đường dẫn chỉ số (sống sót qua thay innerHTML) ----------

export interface SerializedSelection {
  start: number[];
  startOffset: number;
  end: number[];
  endOffset: number;
}

function pathOf(node: Node, root: Node): number[] | null {
  const path: number[] = [];
  for (let n: Node | null = node; n && n !== root; n = n.parentNode) {
    if (!n.parentNode) return null;
    path.unshift(Array.prototype.indexOf.call(n.parentNode.childNodes, n));
  }
  return path;
}

function nodeAt(path: number[], root: Node): Node | null {
  let n: Node | null = root;
  for (const i of path) {
    n = n?.childNodes[i] ?? null;
    if (!n) return null;
  }
  return n;
}

export function serializeSelection(root: HTMLElement): SerializedSelection | null {
  const sel = getSelection();
  if (!sel || !sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  if (!root.contains(r.startContainer) || !root.contains(r.endContainer)) return null;
  const start = pathOf(r.startContainer, root);
  const end = pathOf(r.endContainer, root);
  return start && end ? { start, startOffset: r.startOffset, end, endOffset: r.endOffset } : null;
}

export function restoreSelection(root: HTMLElement, s: SerializedSelection | null) {
  if (!s) return;
  const sn = nodeAt(s.start, root);
  const en = nodeAt(s.end, root);
  if (!sn || !en) return;
  const clamp = (n: Node, o: number) => Math.min(o, n.nodeType === 3 ? (n as Text).length : n.childNodes.length);
  const r = document.createRange();
  try {
    r.setStart(sn, clamp(sn, s.startOffset));
    r.setEnd(en, clamp(en, s.endOffset));
  } catch {
    return;
  }
  const sel = getSelection();
  sel?.removeAllRanges();
  sel?.addRange(r);
}

export function placeCaretAtStart(el: Node) {
  const r = document.createRange();
  r.setStart(el, 0);
  r.collapse(true);
  const sel = getSelection();
  sel?.removeAllRanges();
  sel?.addRange(r);
}

export function escapeHtml(s: string) {
  return s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
}

/** Đọc file thành data URL. */
export function readAsDataUrl(file: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as string);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

/** Kích thước tự nhiên của ảnh từ data URL / URL. */
export function imageSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 300, h: 200 });
    img.src = src;
  });
}
