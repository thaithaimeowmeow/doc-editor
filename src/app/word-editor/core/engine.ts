import { CONTENT_CSS } from './content-css';
import {
  PAGE_BREAK_HTML,
  closestBlock,
  closestTag,
  escapeHtml,
  imageSize,
  placeCaretAtStart,
  readAsDataUrl,
  selectedBlocks,
  textNodesInRange,
} from './dom';
import { Finder, FindOptions } from './find';
import { History } from './history';
import { Paginator } from './paginator';
import { buildPrintHtml, fillFields, printHtml } from './print';
import { cleanPastedHtml, sanitize, stripTransient } from './sanitize';
import * as T from './table';
import { DEFAULT_PAGE, PX_PER_PT, PageSetup, cssColorToHex, cssToPx } from './units';

export interface WordDocument {
  html: string;
  header: string;
  footer: string;
  firstHeader: string;
  firstFooter: string;
  differentFirstPage: boolean;
  page: PageSetup;
  title: string;
}

export type BlockStyle = 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'title' | 'subtitle' | 'quote' | 'pre';
export type HFKind = 'header' | 'footer';

export interface EditorState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  sub: boolean;
  sup: boolean;
  align: 'left' | 'center' | 'right' | 'justify';
  block: BlockStyle | 'li' | 'other';
  list: 'ul' | 'ol' | null;
  fontName: string;
  fontSize: number;
  lineHeight: string;
  inTable: boolean;
  cellsSelected: number;
  inLink: boolean;
  image: boolean;
  canUndo: boolean;
  canRedo: boolean;
  painter: boolean;
  /** Thụt lề đoạn hiện tại (px), cho thước kẻ. */
  indentLeft: number;
  indentFirst: number;
  indentRight: number;
}

export interface LayoutInfo {
  pages: number;
  currentPage: number;
  words: number;
  chars: number;
}

export type EngineRequest = 'link' | 'find' | 'replace' | 'save' | 'print' | 'open';

export interface EngineListener {
  state?(s: EditorState): void;
  layout?(l: LayoutInfo): void;
  change?(): void;
  hfMode?(mode: { kind: HFKind; first: boolean } | null): void;
  request?(r: EngineRequest): void;
}

const GAP = 16;
/** Lệnh vẫn chạy được ở chế độ chỉ xem. */
const READONLY_OK = new Set(['copy', 'selectAll']);
const ZOOM_MIN = 0.25;
const ZOOM_MAX = 4;
let instanceSeq = 0;

/**
 * Engine soạn thảo kiểu Word, không phụ thuộc framework: chỉ cần một phần tử host (vùng cuộn).
 * Vẽ trang, header/footer, xử lý lệnh, phím tắt, dán, bảng, ảnh, chia trang, in.
 */
export class WordEngine {
  readonly id = `we${++instanceSeq}`;
  readonly body: HTMLElement;

  private readonly sizer: HTMLElement;
  private readonly stage: HTMLElement;
  private readonly pagesEl: HTMLElement;
  private readonly overlay: HTMLElement;
  private readonly sheet: HTMLStyleElement;
  private readonly spacerSheet: HTMLStyleElement;
  private readonly paginator: Paginator;
  private readonly history: History;
  private readonly finder: Finder;
  private readonly observer: MutationObserver;
  private readonly resizeObserver: ResizeObserver;

  private doc: Omit<WordDocument, 'html'> = {
    header: '',
    footer: '',
    firstHeader: '',
    firstFooter: '',
    differentFirstPage: false,
    page: { ...DEFAULT_PAGE },
    title: 'Tài liệu',
  };
  private zoom = 1;
  private view: 'print' | 'web' = 'print';
  private pageCount = 1;
  private currentPage = 1;
  private savedRange: Range | null = null;
  private layoutQueued = false;
  private commitTimer = 0;
  private changeTimer = 0;
  private hf: { kind: HFKind; first: boolean; el: HTMLElement } | null = null;
  private image: HTMLImageElement | null = null;
  private cellAnchor: HTMLTableCellElement | null = null;
  private selectedCells: HTMLTableCellElement[] = [];
  private painter: Record<string, string> | null = null;
  private pendingStyle: { prop: string; value: string } | null = null;
  private internalDrag = false;
  private colResize: { table: HTMLTableElement; col: number; startX: number; widths: number[] } | null = null;
  private lastState = '';
  private readonly cleanup: (() => void)[] = [];
  private readonly = false;

  constructor(
    private readonly host: HTMLElement,
    private readonly listener: EngineListener = {},
  ) {
    injectContentCss();
    host.classList.add('we-canvas');
    host.innerHTML = '';

    this.sheet = document.createElement('style');
    this.sizer = el('div', 'we-sizer');
    this.stage = el('div', 'we-stage');
    this.pagesEl = el('div', 'we-pages');
    this.overlay = el('div', 'we-img-overlay');
    this.overlay.innerHTML = '<i data-h="nw"></i><i data-h="ne"></i><i data-h="sw"></i><i data-h="se"></i>';
    this.body = el('div', 'we-body we-doc');
    this.body.contentEditable = 'true';
    this.body.spellcheck = false;
    this.body.setAttribute('role', 'textbox');
    this.body.setAttribute('aria-multiline', 'true');
    this.body.setAttribute('aria-label', 'Nội dung tài liệu');
    this.body.dataset['weScope'] = this.id;
    this.body.innerHTML = '<p><br></p>';

    this.stage.append(this.pagesEl, this.body, this.overlay);
    this.sizer.append(this.stage);
    host.append(this.sheet, this.sizer);

    this.spacerSheet = document.createElement('style');
    host.append(this.spacerSheet);
    this.paginator = new Paginator(this.body, this.spacerSheet, `[data-we-scope="${this.id}"]`);
    this.history = new History(this.body, f => stripTransient(f));
    this.finder = new Finder(this.body, this.id);
    this.sheet.textContent =
      `::highlight(${this.id}-all) { background: #fff176; color: inherit; }` +
      `::highlight(${this.id}-cur) { background: #ff9800; color: #000; }`;

    this.observer = new MutationObserver(() => this.onMutate());
    this.observer.observe(this.body, {
      subtree: true,
      childList: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['style', 'src', 'colspan', 'rowspan', 'href', 'start', 'width', 'height'],
    });
    this.resizeObserver = new ResizeObserver(() => this.scheduleLayout());
    this.resizeObserver.observe(this.body);
    this.resizeObserver.observe(host);

    this.bindEvents();
    this.applyGeometry();
    this.history.reset();
    this.scheduleLayout();
  }

  destroy() {
    this.observer.disconnect();
    this.resizeObserver.disconnect();
    this.finder.clear();
    for (const fn of this.cleanup) fn();
    this.host.innerHTML = '';
  }

  // =============================================================================================
  // Tài liệu
  // =============================================================================================

  getDocument(): WordDocument {
    return { ...this.doc, page: { ...this.doc.page }, html: this.getHtml() };
  }

  setDocument(d: Partial<WordDocument>) {
    this.exitHeaderFooter();
    const { html, ...rest } = d;
    this.doc = { ...this.doc, ...rest, page: { ...this.doc.page, ...(rest.page ?? {}) } };
    for (const k of ['header', 'footer', 'firstHeader', 'firstFooter'] as const) this.doc[k] = sanitize(this.doc[k] ?? '');
    if (html !== undefined) this.setHtml(html, false);
    this.applyGeometry();
    this.renderPages(true);
    this.history.reset();
    this.scheduleLayout();
    this.emitChange();
  }

  getHtml(): string {
    const clone = this.body.cloneNode(true) as HTMLElement;
    stripTransient(clone);
    const html = clone.innerHTML;
    return html === '<p><br></p>' ? '' : html;
  }

  setHtml(html: string, resetHistory = true) {
    this.selectImage(null);
    this.clearCellSelection();
    this.body.innerHTML = sanitize(html) || '<p><br></p>';
    this.normalizeBody();
    this.observer.takeRecords();
    if (resetHistory) this.history.reset();
    this.scheduleLayout();
  }

  get page() {
    return { ...this.doc.page };
  }

  setPage(page: Partial<PageSetup>) {
    this.doc.page = { ...this.doc.page, ...page };
    this.applyGeometry();
    this.renderPages(true);
    this.scheduleLayout();
    this.emitChange();
  }

  get title() {
    return this.doc.title;
  }
  setTitle(t: string) {
    this.doc.title = t;
    this.emitChange();
  }

  get differentFirstPage() {
    return this.doc.differentFirstPage;
  }

  hasHeaderFooter(kind: HFKind) {
    return !!(this.doc[kind] || (this.doc.differentFirstPage && this.doc[kind === 'header' ? 'firstHeader' : 'firstFooter']));
  }

  // =============================================================================================
  // Hiển thị: zoom, chế độ xem, ký tự định dạng
  // =============================================================================================

  getZoom() {
    return this.zoom;
  }

  setZoom(z: number | 'width' | 'page') {
    const p = this.doc.page;
    const avail = this.host.clientWidth - 48;
    if (z === 'width') z = avail / p.width;
    else if (z === 'page') z = Math.min(avail / p.width, (this.host.clientHeight - 48) / p.height);
    this.zoom = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z * 100) / 100));
    this.applyGeometry();
    this.updateOverlay();
    this.emitLayout();
  }

  getView() {
    return this.view;
  }

  setView(v: 'print' | 'web') {
    this.view = v;
    this.host.classList.toggle('we-web', v === 'web');
    if (v === 'web') this.paginator.clear();
    this.applyGeometry();
    this.renderPages(true);
    this.scheduleLayout();
  }

  setShowMarks(on: boolean) {
    this.host.classList.toggle('we-marks', on);
  }

  /** Chỉ xem: khoá soạn thảo, vẫn cho chọn, sao chép, tìm, in. */
  setReadonly(on: boolean) {
    this.readonly = on;
    if (on) this.exitHeaderFooter();
    this.body.contentEditable = on ? 'false' : 'true';
    this.host.classList.toggle('we-readonly', on);
    this.selectImage(null);
  }

  isReadonly() {
    return this.readonly;
  }

  setGridlines(on: boolean) {
    this.host.classList.toggle('we-gridlines', on);
  }

  focus() {
    const target = this.hf?.el ?? this.body;
    const sel = getSelection();
    const inside = !!sel?.rangeCount && target.contains(sel.getRangeAt(0).startContainer);
    target.focus({ preventScroll: true });
    // Chỉ khôi phục vùng chọn đã lưu khi vùng chọn hiện tại nằm ngoài tài liệu (vd: vừa gõ vào ô cỡ chữ).
    if (!inside && this.savedRange && target.contains(this.savedRange.startContainer)) {
      const sel = getSelection();
      sel?.removeAllRanges();
      sel?.addRange(this.savedRange);
    }
  }

  // =============================================================================================
  // Lệnh
  // =============================================================================================

  /** Chạy một lệnh định dạng/chèn. `value` tuỳ lệnh. */
  exec(cmd: string, value?: unknown): void {
    if (this.readonly && !READONLY_OK.has(cmd)) return;
    if (cmd === 'undo') return this.undo();
    if (cmd === 'redo') return this.redo();
    this.focus();
    this.flushHistory();
    const v = value as string;
    const ed = this.editable();

    switch (cmd) {
      case 'bold':
      case 'italic':
      case 'underline':
      case 'strikeThrough':
      case 'subscript':
      case 'superscript':
      case 'cut':
      case 'copy':
      case 'selectAll':
      case 'unlink':
        this.native(cmd);
        break;
      case 'removeFormat':
        this.native('removeFormat');
        this.native('unlink');
        for (const b of this.blocks()) {
          b.style.removeProperty('background-color');
        }
        break;
      case 'fontName':
        this.native('fontName', v);
        break;
      case 'fontSize':
        this.inlineStyle('font-size', `${Number(v)}pt`);
        break;
      case 'growFont':
      case 'shrinkFont': {
        const sizes = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
        const cur = this.state().fontSize;
        const next = cmd === 'growFont' ? sizes.find(s => s > cur) ?? cur + 10 : [...sizes].reverse().find(s => s < cur) ?? Math.max(1, cur - 1);
        this.inlineStyle('font-size', `${next}pt`);
        break;
      }
      case 'foreColor':
        this.native('foreColor', v || '#000000');
        break;
      case 'hiliteColor':
        this.native('hiliteColor', v || 'transparent');
        break;
      case 'changeCase':
        this.changeCase(v as 'upper' | 'lower' | 'title' | 'sentence' | 'toggle');
        break;
      case 'formatPainter':
        this.painter = this.painter ? null : this.captureFormat();
        this.host.classList.toggle('we-painting', !!this.painter);
        break;
      case 'align': {
        const map: Record<string, string> = { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' };
        if (this.image && this.image.parentElement) {
          const b = closestBlock(this.image, ed);
          if (b) b.style.textAlign = v;
        } else this.native(map[v] ?? 'justifyLeft');
        break;
      }
      case 'insertUnorderedList':
      case 'insertOrderedList': {
        const want = cmd === 'insertOrderedList' ? 'OL' : 'UL';
        const list = this.closest(['ul', 'ol']);
        if (!(list && list.tagName === want && value)) this.native(cmd);
        const after = this.closest(['ul', 'ol']);
        if (after && value) after.style.listStyleType = v;
        break;
      }
      case 'indent':
      case 'outdent':
        this.indent(cmd === 'indent');
        break;
      case 'lineHeight':
        for (const b of this.blocks()) b.style.lineHeight = v;
        break;
      case 'spaceBefore':
      case 'spaceAfter':
        for (const b of this.blocks()) b.style.setProperty(cmd === 'spaceBefore' ? 'margin-top' : 'margin-bottom', `${Number(v)}pt`);
        break;
      case 'paragraphIndent': {
        const { left, first, right } = value as { left?: number; first?: number; right?: number };
        for (const b of this.blocks()) {
          if (left !== undefined) b.style.marginLeft = left ? `${round(left / PX_PER_PT)}pt` : '';
          if (first !== undefined) b.style.textIndent = first ? `${round(first / PX_PER_PT)}pt` : '';
          if (right !== undefined) b.style.marginRight = right ? `${round(right / PX_PER_PT)}pt` : '';
        }
        break;
      }
      case 'paragraphShading':
        for (const b of this.blocks()) b.style.backgroundColor = v === 'transparent' ? '' : v;
        break;
      case 'block':
        this.formatBlock(v as BlockStyle);
        break;
      case 'insertTable': {
        const { rows, cols } = value as { rows: number; cols: number };
        this.insertHtml(T.createTableHtml(rows, cols, this.contentWidth()));
        break;
      }
      case 'insertImage':
        void this.insertImage(v);
        return;
      case 'insertLink': {
        const { href, text } = value as { href: string; text?: string };
        const link = this.closest(['a']) as HTMLAnchorElement | null;
        if (link) {
          link.href = href;
          if (text) link.textContent = text;
        } else if (getSelection()?.isCollapsed ?? true) {
          this.insertHtml(`<a href="${escapeHtml(href)}">${escapeHtml(text || href)}</a>`);
        } else {
          this.native('createLink', href);
        }
        break;
      }
      case 'insertHr':
        this.insertHtml('<hr><p><br></p>');
        break;
      case 'insertPageBreak':
        if (!this.hf) this.insertHtml(PAGE_BREAK_HTML + '<p><br></p>');
        break;
      case 'insertSymbol':
        this.native('insertText', v);
        break;
      case 'insertText':
        this.native('insertText', v);
        break;
      case 'insertDate':
        this.native('insertText', new Date().toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }));
        break;
      case 'insertTab':
        this.insertHtml('<span style="white-space: pre">\t</span>');
        break;
      case 'insertField':
        this.insertHtml(`<span class="we-field" data-we-field="${v === 'pages' ? 'pages' : 'page'}" contenteditable="false">${v === 'pages' ? this.pageCount : this.currentPage}</span>&#8203;`);
        break;

      // ----- Bảng -----
      case 'rowAbove':
      case 'rowBelow':
      case 'colLeft':
      case 'colRight':
      case 'deleteRow':
      case 'deleteCol':
      case 'deleteTable':
      case 'mergeCells':
      case 'splitCell':
      case 'distributeCols':
        this.tableCommand(cmd);
        break;
      case 'cellShading':
        for (const c of this.cells()) c.style.backgroundColor = v === 'transparent' ? '' : v;
        break;
      case 'cellVAlign':
        for (const c of this.cells()) c.style.verticalAlign = v;
        break;
      case 'tableBorders':
        this.tableBorders(v as 'all' | 'none' | 'outside' | 'inside');
        break;
      case 'tableHeaderRow': {
        const table = this.closest(['table']) as HTMLTableElement | null;
        const first = table?.rows[0];
        if (first) for (const c of Array.from(first.cells)) c.style.fontWeight = c.style.fontWeight === 'bold' ? '' : 'bold';
        break;
      }

      // ----- Ảnh -----
      case 'imageWidth':
        if (this.image) {
          const w = (this.contentWidth() * Number(v)) / 100;
          const ratio = this.image.naturalHeight / (this.image.naturalWidth || 1);
          this.image.style.width = `${Math.round(w)}px`;
          this.image.style.height = `${Math.round(w * ratio)}px`;
          this.updateOverlay();
        }
        break;
      case 'imageAlt':
        if (this.image) this.image.alt = v;
        break;
      case 'deleteImage':
        if (this.image) {
          this.image.remove();
          this.selectImage(null);
        }
        break;

      // ----- Header / footer -----
      case 'editHeader':
      case 'editFooter':
        this.editHeaderFooter(cmd === 'editHeader' ? 'header' : 'footer', 0);
        return;
      case 'closeHeaderFooter':
        this.exitHeaderFooter();
        return;
      case 'removeHeaderFooter': {
        const kind = v as HFKind;
        this.doc[kind] = '';
        this.doc[kind === 'header' ? 'firstHeader' : 'firstFooter'] = '';
        this.exitHeaderFooter();
        this.renderPages(true);
        this.emitChange();
        return;
      }
      case 'pageNumber':
        this.insertPageNumber(value as { where: HFKind; align: string; format: string });
        return;
      case 'differentFirstPage':
        this.doc.differentFirstPage = !!value;
        if (this.hf) {
          const kind = this.hf.kind;
          this.exitHeaderFooter();
          this.editHeaderFooter(kind, 0);
        }
        this.renderPages(true);
        this.emitChange();
        return;
    }
    this.afterCommand();
  }

  undo() {
    if (this.hf) {
      document.execCommand('undo');
      return;
    }
    this.selectImage(null);
    if (this.history.undo()) this.afterRestore();
  }

  redo() {
    if (this.hf) {
      document.execCommand('redo');
      return;
    }
    this.selectImage(null);
    if (this.history.redo()) this.afterRestore();
  }

  /** Dán từ nút trên ribbon (cần quyền clipboard của trình duyệt). */
  async pasteFromClipboard(plain = false): Promise<boolean> {
    try {
      if (!plain && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const html = await (await item.getType('text/html')).text();
            this.focus();
            this.insertHtml(cleanPastedHtml(html));
            this.afterCommand();
            return true;
          }
          const img = item.types.find(t => t.startsWith('image/'));
          if (img) {
            await this.insertImage(await readAsDataUrl(await item.getType(img)));
            return true;
          }
        }
      }
      const text = await navigator.clipboard.readText();
      this.focus();
      this.native('insertText', text);
      this.afterCommand();
      return true;
    } catch {
      return false;
    }
  }

  async insertImage(src: string) {
    const { w, h } = await imageSize(src);
    const max = this.contentWidth();
    const width = Math.min(w, max);
    const height = Math.round((h * width) / (w || 1));
    this.focus();
    this.flushHistory();
    this.insertHtml(`<img src="${src}" style="width: ${Math.round(width)}px; height: ${height}px" alt="">`);
    this.afterCommand();
  }

  // ----- Tìm & thay -----

  find(query: string, opts: FindOptions) {
    const n = this.finder.search(query, opts);
    return { count: n, current: this.finder.current };
  }
  findNext(backward = false) {
    const r = this.finder.next(backward);
    if (r) this.scrollToRange(r);
    return this.finder.current;
  }
  replace(query: string, replacement: string, opts: FindOptions) {
    if (this.finder.current < 0) this.finder.next();
    this.flushHistory();
    this.finder.replaceCurrent(replacement);
    this.afterCommand();
    const res = this.find(query, opts);
    this.findNext();
    return res;
  }
  replaceAll(query: string, replacement: string, opts: FindOptions) {
    this.flushHistory();
    const n = this.finder.replaceAll(query, replacement, opts);
    this.afterCommand();
    return n;
  }
  clearFind() {
    this.finder.clear();
  }

  // ----- In -----

  async print() {
    this.exitHeaderFooter();
    const wasWeb = this.view === 'web';
    if (wasWeb) this.setView('print');
    this.layoutNow();
    const p = this.doc.page;
    const stride = p.height + GAP;
    const avail = p.height - p.marginTop - p.marginBottom;

    // Mỗi trang in chỉ chứa các khối nằm trên trang đó (không nhân cả tài liệu theo số trang:
    // tài liệu 50 trang nhiều ảnh sẽ vượt giới hạn độ dài chuỗi của trình duyệt).
    const origin = this.body.getBoundingClientRect().top;
    const blocks = Array.from(this.body.children) as HTMLElement[];
    const geo = blocks.map(b => {
      const r = b.getBoundingClientRect();
      const cs = getComputedStyle(b);
      return {
        top: (r.top - origin) / this.zoom - (parseFloat(cs.marginTop) || 0),
        bottom: (r.bottom - origin) / this.zoom + (parseFloat(cs.marginBottom) || 0),
      };
    });
    const htmlOf = (b: HTMLElement) => {
      const c = b.cloneNode(true) as HTMLElement;
      for (const e of [c, ...Array.from(c.querySelectorAll('.we-cell-sel, .we-img-sel'))]) e.classList.remove('we-cell-sel', 'we-img-sel');
      return c.outerHTML;
    };
    const slices = Array.from({ length: this.pageCount }, (_, i) => {
      const lo = i * stride;
      const hi = lo + avail;
      const idx = geo.flatMap((g, j) => (g.bottom > lo + 0.5 && g.top < hi - 0.5 ? [j] : []));
      return idx.length ? { html: idx.map(j => htmlOf(blocks[j])).join(''), offset: geo[idx[0]].top - lo } : { html: '', offset: 0 };
    });

    const html = buildPrintHtml({
      page: p,
      slices,
      spacerCss: this.spacerSheet.textContent ?? '',
      scope: this.id,
      headerFor: i => fillFields(this.hfSource('header', i), i + 1, this.pageCount),
      footerFor: i => fillFields(this.hfSource('footer', i), i + 1, this.pageCount),
      title: this.doc.title,
    });
    await printHtml(html);
    if (wasWeb) this.setView('web');
  }

  // =============================================================================================
  // Trạng thái
  // =============================================================================================

  state(): EditorState {
    const ed = this.editable();
    const sel = getSelection();
    const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
    const inside = !!node && ed.contains(node);
    const elx = inside ? (node!.nodeType === 1 ? (node as HTMLElement) : node!.parentElement!) : ed;
    const cs = getComputedStyle(elx);
    const block = inside ? closestBlock(node, ed) : null;
    const bcs = block ? getComputedStyle(block) : cs;
    const q = (c: string) => {
      try {
        return inside && document.queryCommandState(c);
      } catch {
        return false;
      }
    };
    const list = inside ? (closestTag(node, ['ul', 'ol'], ed) as HTMLElement | null) : null;
    const align = bcs.textAlign;
    const fontPx = parseFloat(cs.fontSize) || 16;
    return {
      bold: q('bold'),
      italic: q('italic'),
      underline: q('underline'),
      strike: q('strikeThrough'),
      sub: q('subscript'),
      sup: q('superscript'),
      align: align === 'center' ? 'center' : align === 'right' || align === 'end' ? 'right' : align === 'justify' ? 'justify' : 'left',
      block: blockStyleOf(block),
      list: list ? (list.tagName.toLowerCase() as 'ul' | 'ol') : null,
      fontName: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
      fontSize: Math.round(fontPx * 0.75 * 2) / 2,
      lineHeight: block?.style.lineHeight || '',
      inTable: inside && !!closestTag(node, ['td', 'th'], ed),
      cellsSelected: this.selectedCells.length,
      inLink: inside && !!closestTag(node, ['a'], ed),
      image: !!this.image,
      canUndo: this.hf ? true : this.history.canUndo,
      canRedo: this.hf ? true : this.history.canRedo,
      painter: !!this.painter,
      indentLeft: block ? parseFloat(bcs.marginLeft) || 0 : 0,
      indentFirst: block ? parseFloat(bcs.textIndent) || 0 : 0,
      indentRight: block ? parseFloat(bcs.marginRight) || 0 : 0,
    };
  }

  // =============================================================================================
  // Nội bộ: hình học, trang, header/footer
  // =============================================================================================

  private contentWidth() {
    const p = this.doc.page;
    return p.width - p.marginLeft - p.marginRight;
  }

  private applyGeometry() {
    const p = this.doc.page;
    const s = this.stage.style;
    const web = this.view === 'web';
    if (web) {
      const w = Math.max(320, this.host.clientWidth - 48) / this.zoom;
      s.width = `${w}px`;
      this.body.style.cssText = `left:0;top:0;width:${w}px;min-height:${Math.max(200, this.host.clientHeight - 48) / this.zoom}px`;
    } else {
      s.width = `${p.width}px`;
      this.body.style.cssText =
        `left:${p.marginLeft}px;top:${p.marginTop}px;width:${this.contentWidth()}px;` +
        `min-height:${p.height - p.marginTop - p.marginBottom}px`;
    }
    s.transform = this.zoom === 1 ? '' : `scale(${this.zoom})`;
    this.host.style.setProperty('--we-zoom', String(this.zoom));
    this.host.style.setProperty('--we-avail', `${p.height - p.marginTop - p.marginBottom}px`);
    this.updateSizer();
  }

  private updateSizer() {
    const p = this.doc.page;
    const web = this.view === 'web';
    const width = web ? parseFloat(this.stage.style.width) : p.width;
    const height = web ? this.body.offsetHeight : this.pageCount * (p.height + GAP) - GAP;
    this.stage.style.height = `${height}px`;
    this.sizer.style.width = `${width * this.zoom}px`;
    this.sizer.style.height = `${height * this.zoom}px`;
  }

  private hfSource(kind: HFKind, pageIndex: number) {
    if (pageIndex === 0 && this.doc.differentFirstPage) return this.doc[kind === 'header' ? 'firstHeader' : 'firstFooter'];
    return this.doc[kind];
  }

  /** Vẽ các trang nền + header/footer. `force` để vẽ lại nội dung header/footer. */
  private renderPages(force = false) {
    if (this.view === 'web') {
      this.pagesEl.replaceChildren();
      return;
    }
    const p = this.doc.page;
    const stride = p.height + GAP;
    while (this.pagesEl.children.length > this.pageCount) this.pagesEl.lastElementChild!.remove();
    while (this.pagesEl.children.length < this.pageCount) {
      const page = el('div', 'we-page');
      const header = el('div', 'we-hf we-doc we-header');
      const footer = el('div', 'we-hf we-doc we-footer');
      page.append(header, footer);
      this.pagesEl.append(page);
      force = true;
    }
    Array.from(this.pagesEl.children).forEach((node, i) => {
      const page = node as HTMLElement;
      page.style.cssText = `top:${i * stride}px;width:${p.width}px;height:${p.height}px`;
      const [header, footer] = Array.from(page.children) as HTMLElement[];
      header.style.cssText = `top:${p.header}px;left:${p.marginLeft}px;width:${this.contentWidth()}px;min-height:${Math.max(0, p.marginTop - p.header)}px`;
      footer.style.cssText = `bottom:${p.footer}px;left:${p.marginLeft}px;width:${this.contentWidth()}px;min-height:${Math.max(0, p.marginBottom - p.footer)}px`;
      for (const [elx, kind] of [[header, 'header'], [footer, 'footer']] as const) {
        if (this.hf?.el === elx) continue;
        const html = fillFields(this.hfSource(kind, i), i + 1, this.pageCount);
        if (force || elx.dataset['src'] !== html) {
          elx.innerHTML = html;
          elx.dataset['src'] = html;
        }
        elx.dataset['label'] = `${kind === 'header' ? 'Đầu trang' : 'Chân trang'}${i === 0 && this.doc.differentFirstPage ? ' trang đầu' : ''}`;
      }
    });
  }

  private scheduleLayout() {
    if (this.layoutQueued) return;
    this.layoutQueued = true;
    const run = () => {
      if (!this.layoutQueued) return;
      this.layoutQueued = false;
      this.layoutNow();
    };
    // rAF bị dừng khi tab/cửa sổ không hiển thị; setTimeout đảm bảo vẫn chạy.
    requestAnimationFrame(run);
    setTimeout(run, 120);
  }

  private layoutNow() {
    const p = this.doc.page;
    if (this.view === 'print') {
      const before = this.pageCount;
      this.observer.takeRecords();
      this.pageCount = this.paginator.run({ stride: p.height + GAP, avail: p.height - p.marginTop - p.marginBottom }, this.zoom);
      this.observer.takeRecords();
      this.renderPages(before !== this.pageCount && /data-we-field="pages"/.test(this.doc.header + this.doc.footer + this.doc.firstHeader + this.doc.firstFooter));
    } else {
      this.pageCount = 1;
    }
    this.updateSizer();
    this.updateOverlay();
    this.updateCurrentPage();
    this.emitLayout();
  }

  private updateCurrentPage() {
    const sel = getSelection();
    if (!sel?.rangeCount || !this.body.contains(sel.anchorNode)) return;
    const r = sel.getRangeAt(0).getBoundingClientRect();
    const top = r.top || (sel.anchorNode as HTMLElement).parentElement?.getBoundingClientRect().top || 0;
    const y = (top - this.body.getBoundingClientRect().top) / this.zoom;
    const stride = this.doc.page.height + GAP;
    this.currentPage = Math.min(this.pageCount, Math.max(1, Math.floor(y / stride) + 1));
  }

  private emitLayout() {
    let words = 0;
    let chars = 0;
    const walker = document.createTreeWalker(this.body, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const t = (n as Text).data;
      chars += t.replace(/\s/g, '').length;
      words += t.split(/\s+/).filter(Boolean).length;
    }
    this.listener.layout?.({ pages: this.pageCount, currentPage: this.currentPage, words, chars });
  }

  private emitState() {
    const s = this.state();
    const key = JSON.stringify(s);
    if (key === this.lastState) return;
    this.lastState = key;
    this.listener.state?.(s);
  }

  private emitChange() {
    clearTimeout(this.changeTimer);
    // Chỉ gom các thay đổi trong cùng một nhịp; debounce thật nằm ở phía component.
    this.changeTimer = window.setTimeout(() => this.listener.change?.(), 30);
  }

  // ----- Header / footer -----

  editHeaderFooter(kind: HFKind, pageIndex: number) {
    if (this.view === 'web') this.setView('print');
    this.exitHeaderFooter();
    const page = this.pagesEl.children[pageIndex] as HTMLElement | undefined;
    if (!page) return;
    const elx = page.querySelector(kind === 'header' ? '.we-header' : '.we-footer') as HTMLElement;
    const first = pageIndex === 0 && this.doc.differentFirstPage;
    this.hf = { kind, first, el: elx };
    if (!elx.innerHTML.trim()) elx.innerHTML = '<p><br></p>';
    for (const f of Array.from(elx.querySelectorAll<HTMLElement>('.we-field'))) f.contentEditable = 'false';
    elx.contentEditable = 'true';
    elx.classList.add('we-hf-editing');
    this.host.classList.add('we-hf-mode');
    this.body.contentEditable = 'false';
    elx.focus();
    const sel = getSelection();
    const r = document.createRange();
    r.selectNodeContents(elx.lastElementChild ?? elx);
    r.collapse(false);
    sel?.removeAllRanges();
    sel?.addRange(r);
    this.listener.hfMode?.({ kind, first });
    this.emitState();
  }

  exitHeaderFooter() {
    if (!this.hf) return;
    this.saveHeaderFooter();
    const elx = this.hf.el;
    elx.contentEditable = 'false';
    elx.classList.remove('we-hf-editing');
    this.host.classList.remove('we-hf-mode');
    this.body.contentEditable = 'true';
    this.hf = null;
    this.renderPages(true);
    this.listener.hfMode?.(null);
    this.emitChange();
  }

  private saveHeaderFooter() {
    if (!this.hf) return;
    const clone = this.hf.el.cloneNode(true) as HTMLElement;
    stripTransient(clone);
    let html = sanitize(clone.innerHTML);
    const probe = document.createElement('div');
    probe.innerHTML = html;
    if (!probe.textContent?.replace(/​/g, '').trim() && !probe.querySelector('img, .we-field, table')) html = '';
    const key = this.hf.first ? (this.hf.kind === 'header' ? 'firstHeader' : 'firstFooter') : this.hf.kind;
    this.doc[key] = html;
  }

  private insertPageNumber(o: { where: HFKind; align: string; format: string }) {
    const page = '<span class="we-field" data-we-field="page">1</span>';
    const pages = '<span class="we-field" data-we-field="pages">1</span>';
    const text =
      o.format === 'page-of' ? `Trang ${page} / ${pages}` : o.format === 'n-of' ? `${page} / ${pages}` : o.format === 'trang' ? `Trang ${page}` : page;
    const html = `<p style="text-align: ${o.align}">${text}</p>`;
    this.doc[o.where] = (this.doc[o.where] || '') + html;
    if (this.doc.differentFirstPage) {
      const k = o.where === 'header' ? 'firstHeader' : 'firstFooter';
      this.doc[k] = (this.doc[k] || '') + html;
    }
    this.exitHeaderFooter();
    this.renderPages(true);
    this.emitChange();
  }

  // =============================================================================================
  // Nội bộ: lệnh
  // =============================================================================================

  private editable(): HTMLElement {
    return this.hf?.el ?? this.body;
  }

  private native(cmd: string, value?: string) {
    document.execCommand('styleWithCSS', false, 'true');
    document.execCommand(cmd, false, value);
  }

  private insertHtml(html: string) {
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand('insertHTML', false, html);
  }

  private closest(tags: (keyof HTMLElementTagNameMap)[]): HTMLElement | null {
    const sel = getSelection();
    if (!sel?.rangeCount) return null;
    return closestTag(sel.getRangeAt(0).startContainer, tags, this.editable()) as HTMLElement | null;
  }

  private blocks(): HTMLElement[] {
    const sel = getSelection();
    if (!sel?.rangeCount) return [];
    return selectedBlocks(sel.getRangeAt(0), this.editable());
  }

  /** Áp một thuộc tính CSS inline lên vùng chọn (dùng mẹo fontSize=7 rồi thay bằng span). */
  private inlineStyle(prop: string, value: string) {
    const sel = getSelection();
    if (!sel?.rangeCount) return;
    if (sel.isCollapsed) {
      // Chưa chọn gì: áp cho chữ sắp gõ, giống Word.
      this.pendingStyle = { prop, value };
      document.execCommand('styleWithCSS', false, 'false');
      document.execCommand('fontSize', false, '7');
      return;
    }
    document.execCommand('styleWithCSS', false, 'false');
    document.execCommand('fontSize', false, '7');
    this.convertFontTags(prop, value);
  }

  private convertFontTags(prop: string, value: string) {
    const ed = this.editable();
    const fonts = Array.from(ed.querySelectorAll('font[size="7"]'));
    let first: Node | null = null;
    let last: Node | null = null;
    for (const f of fonts) {
      const span = document.createElement('span');
      span.style.setProperty(prop, value);
      for (const d of Array.from(f.querySelectorAll<HTMLElement>('[style]'))) d.style.removeProperty(prop);
      span.append(...Array.from(f.childNodes));
      f.replaceWith(span);
      first ??= span;
      last = span;
    }
    // Chrome đôi khi dùng span style="font-size: xxx-large" thay vì <font size=7>.
    for (const s of Array.from(ed.querySelectorAll<HTMLElement>('span[style*="xxx-large"]'))) {
      s.style.setProperty(prop, value);
      if (prop !== 'font-size') s.style.removeProperty('font-size');
      first ??= s;
      last = s;
    }
    if (first && last) {
      const r = document.createRange();
      r.setStartBefore(first);
      r.setEndAfter(last);
      const sel = getSelection();
      sel?.removeAllRanges();
      sel?.addRange(r);
    }
  }

  private changeCase(mode: 'upper' | 'lower' | 'title' | 'sentence' | 'toggle') {
    const sel = getSelection();
    if (!sel?.rangeCount || sel.isCollapsed) return;
    const nodes = textNodesInRange(sel.getRangeAt(0));
    let sentenceStart = true;
    for (const n of nodes) {
      const t = n.data;
      let out = t;
      if (mode === 'upper') out = t.toLocaleUpperCase('vi');
      else if (mode === 'lower') out = t.toLocaleLowerCase('vi');
      else if (mode === 'toggle') out = Array.from(t).map(c => (c === c.toLocaleUpperCase('vi') ? c.toLocaleLowerCase('vi') : c.toLocaleUpperCase('vi'))).join('');
      else if (mode === 'title') out = t.toLocaleLowerCase('vi').replace(/(^|[\s(\["'“‘-])(\p{L})/gu, (_, a, b: string) => a + b.toLocaleUpperCase('vi'));
      else {
        out = Array.from(t.toLocaleLowerCase('vi'))
          .map(c => {
            if (sentenceStart && /\p{L}/u.test(c)) {
              sentenceStart = false;
              return c.toLocaleUpperCase('vi');
            }
            if (/[.!?]/.test(c)) sentenceStart = true;
            return c;
          })
          .join('');
      }
      if (out !== t) n.data = out;
    }
    this.history.markDirty();
  }

  private captureFormat(): Record<string, string> {
    const sel = getSelection();
    const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
    const elx = node ? (node.nodeType === 1 ? (node as HTMLElement) : node.parentElement!) : this.body;
    const cs = getComputedStyle(elx);
    let deco = '';
    let bg = '';
    for (let e: HTMLElement | null = elx; e && e !== this.body; e = e.parentElement) {
      const c = getComputedStyle(e);
      if (!deco && c.textDecorationLine !== 'none') deco = c.textDecorationLine;
      if (!bg && cssColorToHex(c.backgroundColor) && !closestBlock(e, this.body)?.isSameNode(e)) bg = c.backgroundColor;
    }
    return {
      fontName: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
      fontSize: `${Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2}`,
      bold: String(Number(cs.fontWeight) >= 600),
      italic: String(cs.fontStyle === 'italic'),
      underline: String(deco.includes('underline')),
      strike: String(deco.includes('line-through')),
      color: cs.color,
      bg,
    };
  }

  private applyFormat(f: Record<string, string>) {
    this.flushHistory();
    this.native('removeFormat');
    this.native('fontName', f['fontName']);
    this.inlineStyle('font-size', `${f['fontSize']}pt`);
    this.native('foreColor', f['color']);
    if (f['bg']) this.native('hiliteColor', f['bg']);
    const toggles: [string, string][] = [['bold', 'bold'], ['italic', 'italic'], ['underline', 'underline'], ['strike', 'strikeThrough']];
    for (const [key, cmd] of toggles) if ((f[key] === 'true') !== document.queryCommandState(cmd)) this.native(cmd);
    this.afterCommand();
  }

  private indent(more: boolean) {
    const li = this.closest(['li']);
    if (li) {
      this.native(more ? 'indent' : 'outdent');
      return;
    }
    const step = 36; // pt = 1,27 cm, khoảng tab mặc định của Word
    for (const b of this.blocks()) {
      if (b.tagName === 'TD' || b.tagName === 'TH') continue;
      const cur = (cssToPx(b.style.marginLeft) ?? 0) / PX_PER_PT;
      const next = Math.max(0, Math.round((cur + (more ? step : -step)) / step) * step);
      b.style.marginLeft = next ? `${next}pt` : '';
    }
  }

  private formatBlock(style: BlockStyle) {
    const tag = style === 'title' || style === 'subtitle' ? 'p' : style === 'quote' ? 'blockquote' : style;
    this.native('formatBlock', `<${tag}>`);
    for (const b of this.blocks()) {
      b.classList.remove('we-title', 'we-subtitle');
      if (style === 'title') b.classList.add('we-title');
      if (style === 'subtitle') b.classList.add('we-subtitle');
      if (!b.classList.length) b.removeAttribute('class');
    }
  }

  private cells(): HTMLTableCellElement[] {
    if (this.selectedCells.length) return this.selectedCells;
    const c = this.closest(['td', 'th']) as HTMLTableCellElement | null;
    return c ? [c] : [];
  }

  private tableCommand(cmd: string) {
    this.paginator.clear(); // gỡ hàng đệm tạm trước khi sửa cấu trúc bảng
    const cell = (this.selectedCells[0] ?? this.closest(['td', 'th'])) as HTMLTableCellElement | null;
    const table = cell?.closest('table') as HTMLTableElement | null;
    if (!cell || !table) return;
    let focusCell: HTMLTableCellElement | null = cell;
    switch (cmd) {
      case 'rowAbove':
        focusCell = T.insertRow(table, cell, false).cells[0] ?? cell;
        break;
      case 'rowBelow':
        focusCell = T.insertRow(table, cell, true).cells[0] ?? cell;
        break;
      case 'colLeft':
        T.insertColumn(table, cell, false);
        break;
      case 'colRight':
        T.insertColumn(table, cell, true);
        break;
      case 'deleteRow':
        T.deleteRow(table, cell);
        focusCell = null;
        break;
      case 'deleteCol':
        T.deleteColumn(table, cell);
        focusCell = null;
        break;
      case 'deleteTable': {
        const next = table.nextElementSibling;
        table.remove();
        if (next) placeCaretAtStart(next);
        focusCell = null;
        break;
      }
      case 'mergeCells':
        focusCell = T.mergeCells(table, this.selectedCells.length > 1 ? this.selectedCells : []) ?? cell;
        break;
      case 'splitCell':
        T.splitCell(table, cell);
        break;
      case 'distributeCols':
        T.distributeColumns(table);
        break;
    }
    this.clearCellSelection();
    if (focusCell?.isConnected) placeCaretAtStart(focusCell.querySelector('p') ?? focusCell);
    else if (!table.isConnected || !table.rows.length) {
      if (table.isConnected) table.remove();
    }
  }

  private tableBorders(mode: 'all' | 'none' | 'outside' | 'inside') {
    const table = (this.selectedCells[0] ?? this.closest(['td', 'th']))?.closest('table') as HTMLTableElement | null;
    if (!table) return;
    const g = T.buildGrid(table);
    const line = '1px solid #000';
    for (const row of g.rows) {
      for (const c of Array.from(row.cells)) {
        const p = T.cellPos(g, c);
        const edge = { top: p.row === 0, bottom: p.lastRow === g.rows.length - 1, left: p.col === 0, right: p.lastCol === g.cols - 1 };
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const outer = edge[side];
          const on = mode === 'all' || (mode === 'outside' && outer) || (mode === 'inside' && !outer);
          c.style.setProperty(`border-${side}`, on ? line : 'none');
        }
      }
    }
  }

  private afterCommand() {
    this.normalizeBody();
    this.observer.takeRecords();
    this.history.markDirty();
    this.history.commit();
    this.scheduleLayout();
    this.emitChange();
    this.emitState();
    this.saveRange();
  }

  private afterRestore() {
    this.observer.takeRecords();
    this.scheduleLayout();
    this.emitChange();
    this.emitState();
  }

  private flushHistory() {
    clearTimeout(this.commitTimer);
    if (this.observer.takeRecords().length) this.history.markDirty();
    this.history.commit();
  }

  private onMutate() {
    if (this.hf) {
      this.saveHeaderFooter();
      this.renderPages();
      this.emitChange();
      return;
    }
    this.history.markDirty();
    clearTimeout(this.commitTimer);
    this.commitTimer = window.setTimeout(() => this.history.commit(), 600);
    if (this.pendingStyle && this.body.querySelector('font[size="7"]')) {
      const { prop, value } = this.pendingStyle;
      this.pendingStyle = null;
      const sel = getSelection();
      const r = sel?.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
      this.convertFontTags(prop, value);
      if (r) {
        sel?.removeAllRanges();
        sel?.addRange(r);
      }
    }
    this.scheduleLayout();
    this.emitChange();
  }

  /** Đảm bảo thân văn bản luôn có ít nhất một đoạn và không có chữ trần ở cấp cao nhất. */
  private normalizeBody() {
    if (!this.body.firstChild) this.body.innerHTML = '<p><br></p>';
    let inline: Node[] = [];
    const wrap = () => {
      if (!inline.length) return;
      if (inline.every(n => n.nodeType === 3 && !n.textContent!.trim())) {
        inline.forEach(n => n.parentNode?.removeChild(n));
      } else {
        const p = document.createElement('p');
        inline[0].parentNode!.insertBefore(p, inline[0]);
        p.append(...inline);
      }
      inline = [];
    };
    for (const n of Array.from(this.body.childNodes)) {
      const isBlock = n.nodeType === 1 && /^(P|H\d|UL|OL|TABLE|DIV|BLOCKQUOTE|PRE|HR)$/.test((n as Element).tagName);
      if (isBlock) wrap();
      else if (n.nodeType === 1 || n.nodeType === 3) inline.push(n);
    }
    wrap();
    if (!this.body.firstChild) this.body.innerHTML = '<p><br></p>';
    // Ngắt trang và trường số trang là một khối nguyên, không gõ chữ vào bên trong.
    for (const e of Array.from(this.body.querySelectorAll<HTMLElement>('.we-page-break, .we-field'))) {
      if (e.contentEditable !== 'false') e.contentEditable = 'false';
    }
  }

  // =============================================================================================
  // Nội bộ: sự kiện
  // =============================================================================================

  private bindEvents() {
    const on = <K extends keyof HTMLElementEventMap>(target: HTMLElement | Document, type: K, fn: (e: HTMLElementEventMap[K]) => void, opts?: AddEventListenerOptions) => {
      target.addEventListener(type, fn as EventListener, opts);
      this.cleanup.push(() => target.removeEventListener(type, fn as EventListener, opts));
    };

    on(document, 'selectionchange', () => {
      const sel = getSelection();
      if (!sel?.rangeCount) return;
      const ed = this.editable();
      if (!ed.contains(sel.anchorNode)) return;
      this.saveRange();
      this.updateCurrentPage();
      this.emitState();
      this.emitLayoutThrottled();
    });

    on(this.host, 'keydown', e => this.onKeyDown(e));
    on(this.host, 'beforeinput', e => {
      if (e.inputType === 'historyUndo') {
        e.preventDefault();
        this.undo();
      } else if (e.inputType === 'historyRedo') {
        e.preventDefault();
        this.redo();
      }
    });
    on(this.host, 'paste', e => this.onPaste(e));
    on(this.host, 'dragstart', () => (this.internalDrag = true));
    on(this.host, 'dragend', () => (this.internalDrag = false));
    on(this.host, 'drop', e => this.onDrop(e));
    on(this.host, 'dblclick', e => {
      const hf = (e.target as HTMLElement).closest('.we-hf') as HTMLElement | null;
      if (hf && !hf.isContentEditable && !this.readonly) {
        const page = Array.from(this.pagesEl.children).indexOf(hf.parentElement!);
        this.editHeaderFooter(hf.classList.contains('we-header') ? 'header' : 'footer', Math.max(0, page));
      }
    });
    on(this.host, 'mousedown', e => this.onMouseDown(e));
    on(this.host, 'mousemove', e => this.onMouseMove(e));
    on(document, 'mouseup', () => this.onMouseUp());
    on(this.host, 'click', e => {
      const a = (e.target as HTMLElement).closest('a');
      if (a && (e.ctrlKey || e.metaKey) && this.body.contains(a)) {
        e.preventDefault();
        window.open(a.href, '_blank', 'noopener');
      }
    });
    on(this.host, 'scroll', () => this.updateOverlay(), { passive: true });
    on(this.host, 'wheel', e => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    on(this.overlay, 'mousedown', e => this.startImageResize(e));
  }

  private layoutThrottle = 0;
  private emitLayoutThrottled() {
    clearTimeout(this.layoutThrottle);
    this.layoutThrottle = window.setTimeout(() => this.emitLayout(), 150);
  }

  private saveRange() {
    const sel = getSelection();
    if (sel?.rangeCount && this.editable().contains(sel.getRangeAt(0).startContainer)) this.savedRange = sel.getRangeAt(0).cloneRange();
  }

  private onKeyDown(e: KeyboardEvent) {
    if (this.readonly) {
      const k = e.key.toLowerCase();
      const mod = e.ctrlKey || e.metaKey;
      if (mod && (k === 'f' || k === 'p')) {
        e.preventDefault();
        this.listener.request?.(k === 'f' ? 'find' : 'print');
      }
      return;
    }
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    const run = (cmd: string, value?: unknown) => {
      e.preventDefault();
      this.exec(cmd, value);
    };

    if (e.key === 'Escape') {
      if (this.hf) {
        e.preventDefault();
        this.exitHeaderFooter();
        this.body.focus();
      }
      if (this.painter) {
        this.painter = null;
        this.host.classList.remove('we-painting');
        this.emitState();
      }
      this.selectImage(null);
      return;
    }

    if ((e.key === 'Delete' || e.key === 'Backspace') && this.image) {
      e.preventDefault();
      this.flushHistory();
      this.image.remove();
      this.selectImage(null);
      this.afterCommand();
      return;
    }

    if (e.key === 'Tab') {
      const cell = this.closest(['td', 'th']) as HTMLTableCellElement | null;
      if (cell && !this.closest(['li'])) {
        e.preventDefault();
        this.moveCell(cell, e.shiftKey);
        return;
      }
      if (this.closest(['li'])) return run(e.shiftKey ? 'outdent' : 'indent');
      return run('insertTab');
    }

    if (mod && e.key === 'Enter') return run('insertPageBreak');
    if (!mod) return;

    if (e.shiftKey && (e.key === '>' || e.key === '.')) return run('growFont');
    if (e.shiftKey && (e.key === '<' || e.key === ',')) return run('shrinkFont');
    if (e.key === ']') return run('growFont');
    if (e.key === '[') return run('shrinkFont');
    if (e.shiftKey && (e.key === '+' || e.key === '=')) return run('superscript');
    if (e.key === '=') return run('subscript');
    if (e.altKey && /^[1-4]$/.test(e.key)) return run('block', `h${e.key}`);
    if (e.shiftKey && k === 'l') return run('insertUnorderedList');
    if (e.shiftKey && k === 'n') return run('block', 'p');
    if (e.shiftKey && k === 'c') return run('formatPainter');
    if (e.shiftKey && k === 'a') return run('changeCase', 'upper');

    switch (k) {
      case 'b':
        return run('bold');
      case 'i':
        return run('italic');
      case 'u':
        return run('underline');
      case 'z':
        e.preventDefault();
        return e.shiftKey ? this.redo() : this.undo();
      case 'y':
        e.preventDefault();
        return this.redo();
      case 'e':
        return run('align', 'center');
      case 'l':
        return run('align', 'left');
      case 'r':
        return run('align', 'right');
      case 'j':
        return run('align', 'justify');
      case 'm':
        return run(e.shiftKey ? 'outdent' : 'indent');
      case ' ':
        return run('removeFormat');
      case 'k':
        e.preventDefault();
        return this.listener.request?.('link');
      case 'f':
        e.preventDefault();
        return this.listener.request?.('find');
      case 'h':
        e.preventDefault();
        return this.listener.request?.('replace');
      case 's':
        e.preventDefault();
        return this.listener.request?.('save');
      case 'o':
        e.preventDefault();
        return this.listener.request?.('open');
      case 'p':
        e.preventDefault();
        return this.listener.request?.('print');
      case '1':
        return run('lineHeight', '1');
      case '2':
        return run('lineHeight', '2');
      case '5':
        return run('lineHeight', '1.5');
    }
  }

  private moveCell(cell: HTMLTableCellElement, back: boolean) {
    const table = cell.closest('table') as HTMLTableElement;
    const all = Array.from(table.querySelectorAll<HTMLTableCellElement>('td, th')).filter(
      c => c.closest('table') === table && !c.closest('.we-pb-spacer'),
    );
    const i = all.indexOf(cell);
    let target = all[i + (back ? -1 : 1)];
    if (!target && !back) {
      this.flushHistory();
      target = T.insertRow(table, cell, true).cells[0];
      this.afterCommand();
    }
    if (!target) return;
    const r = document.createRange();
    r.selectNodeContents(target);
    const sel = getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
  }

  private onPaste(e: ClipboardEvent) {
    if (this.readonly) return;
    const dt = e.clipboardData;
    if (!dt) return;
    const files = Array.from(dt.files).filter(f => f.type.startsWith('image/'));
    const html = dt.getData('text/html');
    if (files.length && !html) {
      e.preventDefault();
      for (const f of files) void readAsDataUrl(f).then(src => this.insertImage(src));
      return;
    }
    if (html) {
      e.preventDefault();
      this.flushHistory();
      this.insertHtml(cleanPastedHtml(html));
      this.afterCommand();
      return;
    }
    const text = dt.getData('text/plain');
    if (text) {
      e.preventDefault();
      this.flushHistory();
      this.native('insertText', text);
      this.afterCommand();
    }
  }

  private onDrop(e: DragEvent) {
    if (this.readonly) {
      e.preventDefault();
      return;
    }
    const dt = e.dataTransfer;
    if (!dt || this.internalDrag) return;
    const images = Array.from(dt.files).filter(f => f.type.startsWith('image/'));
    const html = dt.getData('text/html');
    if (!images.length && !html) return;
    e.preventDefault();
    this.placeCaretAt(e.clientX, e.clientY);
    if (images.length) for (const f of images) void readAsDataUrl(f).then(src => this.insertImage(src));
    else {
      this.flushHistory();
      this.insertHtml(cleanPastedHtml(html));
      this.afterCommand();
    }
  }

  private placeCaretAt(x: number, y: number) {
    const d = document as Document & { caretPositionFromPoint?(x: number, y: number): { offsetNode: Node; offset: number } | null };
    const r = document.createRange();
    const pos = d.caretPositionFromPoint?.(x, y);
    if (pos) r.setStart(pos.offsetNode, pos.offset);
    else {
      const cr = document.caretRangeFromPoint?.(x, y);
      if (!cr) return;
      r.setStart(cr.startContainer, cr.startOffset);
    }
    r.collapse(true);
    const sel = getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
  }

  // ----- Chuột: ảnh, chọn ô, đổi độ rộng cột, format painter -----

  private onMouseDown(e: MouseEvent) {
    const target = e.target as HTMLElement;
    if (this.hf && !this.hf.el.contains(target) && this.body.contains(target)) {
      this.exitHeaderFooter();
      this.body.focus();
    }
    if (this.overlay.contains(target)) return;

    // Kéo đổi độ rộng cột.
    const cell = target.closest('td, th') as HTMLTableCellElement | null;
    if (cell && this.body.contains(cell) && this.nearRightEdge(cell, e.clientX)) {
      e.preventDefault();
      const table = cell.closest('table') as HTMLTableElement;
      const g = T.buildGrid(table);
      this.flushHistory();
      this.colResize = { table, col: T.cellPos(g, cell).lastCol, startX: e.clientX, widths: T.columnWidths(table, this.zoom) };
      return;
    }

    if (target.tagName === 'IMG' && this.editable().contains(target)) {
      this.selectImage(target as HTMLImageElement);
      const r = document.createRange();
      r.selectNode(target);
      getSelection()?.removeAllRanges();
      getSelection()?.addRange(r);
      e.preventDefault();
      return;
    }
    this.selectImage(null);
    this.clearCellSelection();
    this.cellAnchor = cell && this.body.contains(cell) ? cell : null;
  }

  private onMouseMove(e: MouseEvent) {
    if (this.colResize) {
      const { table, col, startX, widths } = this.colResize;
      const dx = (e.clientX - startX) / this.zoom;
      const next = [...widths];
      next[col] = Math.max(16, widths[col] + dx);
      if (col + 1 < widths.length) next[col + 1] = Math.max(16, widths[col + 1] - (next[col] - widths[col]));
      T.applyColumnWidths(table, next);
      return;
    }
    const target = e.target as HTMLElement;
    const cell = target.closest?.('td, th') as HTMLTableCellElement | null;
    this.body.classList.toggle('we-col-resize', !!cell && this.body.contains(cell) && !e.buttons && this.nearRightEdge(cell, e.clientX));

    if (e.buttons === 1 && this.cellAnchor && cell && cell !== this.cellAnchor) {
      const table = this.cellAnchor.closest('table');
      if (table && cell.closest('table') === table) {
        this.clearCellSelection();
        this.selectedCells = T.cellsInRect(table as HTMLTableElement, this.cellAnchor, cell);
        for (const c of this.selectedCells) c.classList.add('we-cell-sel');
        this.observer.takeRecords();
        this.emitState();
      }
    }
  }

  private onMouseUp() {
    if (this.colResize) {
      this.colResize = null;
      this.afterCommand();
      return;
    }
    this.cellAnchor = null;
    if (this.painter) {
      const sel = getSelection();
      if (sel && !sel.isCollapsed && this.editable().contains(sel.anchorNode)) {
        const f = this.painter;
        this.painter = null;
        this.host.classList.remove('we-painting');
        this.applyFormat(f);
      }
    }
  }

  private nearRightEdge(cell: HTMLElement, x: number) {
    const r = cell.getBoundingClientRect();
    return Math.abs(r.right - x) <= 4;
  }

  private clearCellSelection() {
    for (const c of this.selectedCells) {
      c.classList.remove('we-cell-sel');
      if (!c.classList.length) c.removeAttribute('class');
    }
    this.selectedCells = [];
    this.observer.takeRecords();
  }

  private selectImage(img: HTMLImageElement | null) {
    if (this.image === img) return;
    this.image?.classList.remove('we-img-sel');
    if (this.image && !this.image.classList.length) this.image.removeAttribute('class');
    this.image = img;
    img?.classList.add('we-img-sel');
    this.observer.takeRecords();
    this.updateOverlay();
    this.emitState();
  }

  private updateOverlay() {
    const img = this.image;
    if (!img || !img.isConnected) {
      this.overlay.style.display = 'none';
      if (img) this.image = null;
      return;
    }
    const s = this.stage.getBoundingClientRect();
    const r = img.getBoundingClientRect();
    const z = this.zoom;
    this.overlay.style.cssText = `display:block;left:${(r.left - s.left) / z}px;top:${(r.top - s.top) / z}px;width:${r.width / z}px;height:${r.height / z}px`;
  }

  private startImageResize(e: MouseEvent) {
    const handle = (e.target as HTMLElement).dataset['h'];
    const img = this.image;
    if (!handle || !img) return;
    e.preventDefault();
    e.stopPropagation();
    this.flushHistory();
    const startX = e.clientX;
    const startW = img.getBoundingClientRect().width / this.zoom;
    const ratio = img.getBoundingClientRect().height / img.getBoundingClientRect().width || 1;
    const dir = handle.includes('w') ? -1 : 1;
    const max = this.contentWidth();
    const move = (ev: MouseEvent) => {
      const w = Math.min(max, Math.max(16, startW + (dir * (ev.clientX - startX)) / this.zoom));
      img.style.width = `${Math.round(w)}px`;
      img.style.height = `${Math.round(w * ratio)}px`;
      this.updateOverlay();
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      this.afterCommand();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  private scrollToRange(r: Range) {
    const rect = r.getBoundingClientRect();
    const host = this.host.getBoundingClientRect();
    if (rect.top < host.top + 40 || rect.bottom > host.bottom - 40) {
      this.host.scrollTop += rect.top - host.top - host.height / 3;
    }
  }
}

// ---------------------------------------------------------------------------------------------------

function el(tag: string, cls: string) {
  const e = document.createElement(tag);
  e.className = cls;
  return e;
}

const round = (n: number) => Math.round(n * 10) / 10;

function blockStyleOf(b: HTMLElement | null): EditorState['block'] {
  if (!b) return 'p';
  if (b.classList.contains('we-title')) return 'title';
  if (b.classList.contains('we-subtitle')) return 'subtitle';
  switch (b.tagName) {
    case 'P':
    case 'DIV':
    case 'TD':
    case 'TH':
      return 'p';
    case 'H1':
    case 'H2':
    case 'H3':
    case 'H4':
      return b.tagName.toLowerCase() as BlockStyle;
    case 'BLOCKQUOTE':
      return 'quote';
    case 'PRE':
      return 'pre';
    case 'LI':
      return 'li';
    default:
      return 'other';
  }
}

let contentCssInjected = false;
function injectContentCss() {
  if (contentCssInjected) return;
  contentCssInjected = true;
  const style = document.createElement('style');
  style.dataset['weContent'] = '';
  style.textContent = CONTENT_CSS;
  document.head.append(style);
}
