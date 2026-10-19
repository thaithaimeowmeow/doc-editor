import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  ViewEncapsulation,
  afterNextRender,
  computed,
  effect,
  forwardRef,
  inject,
  input,
  output,
  signal,
  untracked,
  viewChild,
  WritableSignal,
} from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { ControlValueAccessor, NG_VALUE_ACCESSOR } from '@angular/forms';

import { DOCX_MIME, base64ToBlob, blobToBase64, downloadBlob, textToBase64 } from './core/base64';
import { readAsDataUrl } from './core/dom';
import { EditorState, HFKind, LayoutInfo, WordDocument, WordEngine } from './core/engine';
import { standaloneHtml } from './core/print';
import { MARGIN_PRESETS, PAPER_SIZES, PageSetup, cm, toCm } from './core/units';
import { WeIcon } from './we-icon';

/** Loại giá trị mà editor phát ra khi nội dung thay đổi (xem input `emit`). */
export type WordEmitKind = 'html' | 'htmlBase64' | 'docxBase64' | 'document';

/** Nguồn .docx truyền qua input `docx`: Blob/File, ArrayBuffer, Uint8Array hoặc chuỗi base64 (có/không tiền tố data:). */
export type DocxSource = Blob | ArrayBuffer | Uint8Array | string;

type Tab = 'file' | 'home' | 'insert' | 'layout' | 'view' | 'table' | 'image' | 'hf';
type Dialog = 'link' | 'find' | 'margins' | 'table' | 'alt' | 'imageUrl' | null;

const FONTS = [
  'Times New Roman', 'Arial', 'Calibri', 'Cambria', 'Tahoma', 'Verdana', 'Georgia', 'Garamond', 'Segoe UI',
  'Roboto', 'Courier New', 'Consolas', 'Trebuchet MS', 'Palatino Linotype', 'Book Antiqua', 'Comic Sans MS',
];
const SIZES = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const STANDARD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
const HIGHLIGHTS = ['FFFF00', '00FF00', '00FFFF', 'FF00FF', '0000FF', 'FF0000', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0', '000000'];
const SYMBOLS = '©®™§¶†‡•…–—‘’“”«»‹›°±×÷≠≈≤≥∞√∑∏∫∂∆∇πΩµαβγδθλσφψω€£¥¢₫←↑→↓↔⇒⇔✓✗★☆♥♦♣♠☎✉☐☑½¼¾¹²³'.split('');

/** Pha màu với trắng (t > 0) hoặc đen (t < 0), giống bảng màu chủ đề của Word. */
function mix(hex: string, t: number) {
  const n = parseInt(hex, 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(c => Math.round(t > 0 ? c + (255 - c) * t : c * (1 + t)));
  return ch.map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}
const THEME_GRID = [0.8, 0.6, 0.4, -0.25, -0.5].map((t, row) =>
  THEME.map((h, i) => (i === 0 ? mix(h, -[0.05, 0.15, 0.25, 0.35, 0.5][row]) : i === 1 ? mix(h, [0.5, 0.35, 0.25, 0.15, 0.05][row]) : mix(h, t))),
);

/**
 * Trình soạn thảo kiểu Microsoft Word: ribbon, trang A4, chia trang tự động, header/footer, số trang,
 * bảng, ảnh, tìm/thay, in, nhập/xuất .docx. Giá trị form (`formControl`/`ngModel`) là HTML thân văn bản.
 */
@Component({
  selector: 'app-word-editor',
  imports: [WeIcon, NgTemplateOutlet],
  templateUrl: './word-editor.html',
  styleUrl: './word-editor.css',
  encapsulation: ViewEncapsulation.None,
  changeDetection: ChangeDetectionStrategy.OnPush,
  providers: [{ provide: NG_VALUE_ACCESSOR, useExisting: forwardRef(() => WordEditor), multi: true }],
  host: { class: 'we-root', '(keydown)': 'onHostKey($event)' },
})
export class WordEditor implements ControlValueAccessor {
  /** 'download': Ctrl+S / nút Lưu tải file .docx. 'emit': chỉ phát sự kiện `save` để app tự lưu. */
  readonly saveMode = input<'download' | 'emit'>('download');
  readonly showRuler = input(true);

  // ---------- Dữ liệu vào ----------
  /** Nội dung HTML. Dùng hai chiều: `[(html)]="value"`. */
  readonly html = input<string>();
  /** Toàn bộ tài liệu (HTML + header/footer + khổ giấy + tiêu đề). */
  readonly document = input<Partial<WordDocument>>();
  /** Mở một file .docx (đổi giá trị để mở file khác). */
  readonly docx = input<DocxSource | null>();
  /** Chỉ xem, không cho sửa. */
  readonly readonly = input(false);

  // ---------- Dữ liệu ra (debounce) ----------
  /** Những output nào được phát. Tạo .docx tốn CPU nên chỉ bật khi cần. */
  readonly emit = input<WordEmitKind[]>(['html']);
  /**
   * 'auto': tự phát sau mỗi lần sửa (debounce). 'manual': không phát gì cho đến khi gọi `emitNow()`,
   * tránh tốn CPU khi gõ trên tài liệu lớn (tạo .docx 170–570 ms, lấy HTML 12 MB khoảng 45 ms).
   * Ở chế độ 'manual', `[(html)]` / formControl cũng chỉ cập nhật khi gọi `emitNow()`.
   */
  readonly emitMode = input<'auto' | 'manual'>('auto');
  /** ms chờ sau lần sửa cuối trước khi phát html / htmlBase64 / document. */
  readonly debounce = input(300);
  /** ms chờ sau lần sửa cuối trước khi tạo và phát docxBase64. */
  readonly docxDebounce = input(1500);

  readonly htmlChange = output<string>();
  readonly htmlBase64Change = output<string>();
  readonly docxBase64Change = output<string>();
  readonly documentChange = output<WordDocument>();

  readonly save = output<WordDocument>();
  readonly fileLoaded = output<string>();
  readonly failed = output<unknown>();

  protected readonly fonts = FONTS;
  protected readonly sizes = SIZES;
  protected readonly themeTop = THEME;
  protected readonly themeGrid = THEME_GRID;
  protected readonly standard = STANDARD;
  protected readonly highlights = HIGHLIGHTS;
  protected readonly symbols = SYMBOLS;
  protected readonly paperSizes = Object.entries(PAPER_SIZES);
  protected readonly marginPresets = Object.entries(MARGIN_PRESETS);
  protected readonly gridRows = Array.from({ length: 8 }, (_, i) => i);
  protected readonly gridCols = Array.from({ length: 10 }, (_, i) => i);
  protected readonly styles: { key: string; label: string; cls: string }[] = [
    { key: 'p', label: 'Bình thường', cls: 'st-normal' },
    { key: 'title', label: 'Tiêu đề', cls: 'st-title' },
    { key: 'subtitle', label: 'Phụ đề', cls: 'st-subtitle' },
    { key: 'h1', label: 'Đề mục 1', cls: 'st-h1' },
    { key: 'h2', label: 'Đề mục 2', cls: 'st-h2' },
    { key: 'h3', label: 'Đề mục 3', cls: 'st-h3' },
    { key: 'h4', label: 'Đề mục 4', cls: 'st-h4' },
    { key: 'quote', label: 'Trích dẫn', cls: 'st-quote' },
    { key: 'pre', label: 'Mã', cls: 'st-code' },
  ];

  protected readonly tab = signal<Tab>('home');
  protected readonly menu = signal<string | null>(null);
  protected readonly dialog = signal<Dialog>(null);
  protected readonly state = signal<EditorState | null>(null);
  protected readonly layout = signal<LayoutInfo>({ pages: 1, currentPage: 1, words: 0, chars: 0 });
  protected readonly hfMode = signal<{ kind: HFKind; first: boolean } | null>(null);
  protected readonly zoom = signal(1);
  protected readonly view = signal<'print' | 'web'>('print');
  protected readonly page = signal<PageSetup>(WordEditorDefaults.page());
  protected readonly title = signal('Tài liệu');
  protected readonly differentFirst = signal(false);
  protected readonly marks = signal(false);
  protected readonly gridlines = signal(true);
  protected readonly rulerOn = signal(true);
  protected readonly busy = signal<string | null>(null);
  protected readonly toast = signal<string | null>(null);
  protected readonly fontColor = signal('#C00000');
  protected readonly highlightColor = signal('#FFFF00');
  protected readonly shadingColor = signal('#D9E2F3');
  protected readonly gridHover = signal<[number, number]>([0, 0]);
  protected readonly rulerLeft = signal(0);

  // Form trong hộp thoại.
  protected readonly linkForm = signal({ text: '', href: '' });
  protected readonly findForm = signal({ query: '', replace: '', matchCase: false, wholeWord: false, showReplace: false });
  protected readonly findResult = signal<{ count: number; current: number } | null>(null);
  protected readonly marginForm = signal({ t: 2.54, b: 2.54, l: 2.54, r: 2.54, h: 1.25, f: 1.25 });
  protected readonly tableForm = signal({ rows: 3, cols: 3 });
  protected readonly altText = signal('');
  protected readonly imageUrl = signal('');

  protected readonly contextTable = computed(() => !!this.state()?.inTable);
  protected readonly contextImage = computed(() => !!this.state()?.image);
  protected readonly landscape = computed(() => this.page().width > this.page().height);
  protected readonly zoomPct = computed(() => Math.round(this.zoom() * 100));
  protected readonly paperKey = computed(() => {
    const p = this.page();
    const [w, h] = [Math.min(p.width, p.height), Math.max(p.width, p.height)];
    return Object.entries(PAPER_SIZES).find(([, s]) => Math.abs(s.w - w) < 2 && Math.abs(s.h - h) < 2)?.[0] ?? '';
  });
  protected readonly ruler = computed(() => {
    const p = this.page();
    const z = this.zoom();
    const s = this.state();
    const ticks: { x: number; major: boolean; label: string }[] = [];
    const step = cm(0.25);
    const origin = p.marginLeft;
    for (let x = origin % step; x <= p.width; x += step) {
      const fromMargin = Math.round((x - origin) / step);
      const major = fromMargin % 4 === 0;
      const label = major && fromMargin !== 0 ? String(Math.abs(fromMargin / 4)) : '';
      ticks.push({ x: x * z, major, label });
    }
    const contentLeft = p.marginLeft + (s?.indentLeft ?? 0);
    return {
      width: p.width * z,
      left: p.marginLeft * z,
      right: p.marginRight * z,
      ticks,
      first: (contentLeft + (s?.indentFirst ?? 0)) * z,
      hanging: contentLeft * z,
      indentRight: (p.width - p.marginRight - (s?.indentRight ?? 0)) * z,
    };
  });

  private readonly canvas = viewChild.required<ElementRef<HTMLElement>>('canvas');
  private readonly fileInput = viewChild.required<ElementRef<HTMLInputElement>>('fileInput');
  private readonly imageInput = viewChild.required<ElementRef<HTMLInputElement>>('imageInput');
  private engine?: WordEngine;
  private readonly host = inject<ElementRef<HTMLElement>>(ElementRef).nativeElement;
  private pending: Partial<WordDocument> | null = null;
  private disabled = false;
  private onChange: (html: string) => void = () => {};
  private onTouched: () => void = () => {};
  private toastTimer = 0;
  private emitTimer = 0;
  private docxTimer = 0;
  private docxSeq = 0;
  /** HTML vừa phát ra / vừa nhận vào, để `[(html)]` không tự ghi đè ngược lại (mất con trỏ, mất undo). */
  private lastHtml: string | undefined;
  private resolveReady!: () => void;
  private readonly ready = new Promise<void>(r => (this.resolveReady = r));

  constructor() {
    const destroyRef = inject(DestroyRef);
    const host = this.host;

    effect(() => {
      const html = this.html();
      if (html !== undefined) untracked(() => this.applyHtmlInput(html));
    });
    effect(() => {
      const doc = this.document();
      if (doc) untracked(() => this.setDocument(doc));
    });
    effect(() => {
      const src = this.docx();
      if (src) untracked(() => void this.loadFile(toDocxBlob(src), src instanceof File ? src.name : 'document.docx'));
    });
    effect(() => {
      const ro = this.readonly();
      untracked(() => this.applyEditable(ro));
    });

    afterNextRender(() => {
      const canvas = this.canvas().nativeElement;
      this.rulerOn.set(this.showRuler());
      this.engine = new WordEngine(canvas, {
        state: s => this.state.set(s),
        layout: l => {
          this.layout.set(l);
          this.updateRuler();
        },
        change: () => this.scheduleEmit(),
        hfMode: m => {
          this.hfMode.set(m);
          if (m) this.tab.set('hf');
          else if (this.tab() === 'hf') this.tab.set('home');
        },
        request: r => this.onRequest(r),
      });
      this.engine.setGridlines(true);
      this.engine.setDocument({ page: WordEditorDefaults.page(), ...this.pending });
      this.syncFromEngine();
      this.applyEditable(this.readonly());
      this.resolveReady();

      const ro = new ResizeObserver(() => this.updateRuler());
      ro.observe(canvas);
      canvas.addEventListener('scroll', () => this.updateRuler(), { passive: true });

      // Đăng ký trong afterNextRender để không chạm `document` khi render phía server (SSR).
      const closeMenu = (e: MouseEvent) => {
        if (this.menu() && !(e.target as HTMLElement).closest('.we-menu, .we-has-menu')) this.menu.set(null);
      };
      document.addEventListener('mousedown', closeMenu, true);
      host.addEventListener('focusout', () => this.onTouched());
      destroyRef.onDestroy(() => {
        ro.disconnect();
        document.removeEventListener('mousedown', closeMenu, true);
      });
    });

    destroyRef.onDestroy(() => {
      clearTimeout(this.emitTimer);
      clearTimeout(this.docxTimer);
      this.engine?.destroy();
    });
  }

  // =============================================================================================
  // API công khai
  // =============================================================================================

  getHtml(): string {
    return this.engine?.getHtml() ?? this.pending?.html ?? '';
  }

  setHtml(html: string) {
    if (this.engine) this.engine.setHtml(html);
    else this.pending = { ...this.pending, html };
  }

  getDocument(): WordDocument | null {
    return this.engine?.getDocument() ?? null;
  }

  setDocument(doc: Partial<WordDocument>) {
    if (this.engine) {
      this.engine.setDocument(doc);
      this.syncFromEngine();
    } else this.pending = { ...this.pending, ...doc };
  }

  getHtmlBase64(): Promise<string> {
    return textToBase64(this.getHtml());
  }

  async getDocxBlob(): Promise<Blob> {
    await this.ready;
    const doc = this.engine!.getDocument();
    const { exportDocx } = await import('./core/docx-export');
    return exportDocx(doc);
  }

  async getDocxBase64(): Promise<string> {
    return blobToBase64(await this.getDocxBlob());
  }

  async downloadDocx() {
    await this.run('Đang tạo file .docx…', async () => downloadBlob(await this.getDocxBlob(), `${this.fileName()}.docx`));
  }

  downloadHtml() {
    const doc = this.engine!.getDocument();
    downloadBlob(new Blob([standaloneHtml(doc.html, doc.page, doc.title)], { type: 'text/html;charset=utf-8' }), `${this.fileName()}.html`);
  }

  print() {
    void this.run('Đang chuẩn bị bản in…', () => this.engine!.print());
  }

  openFilePicker() {
    this.fileInput().nativeElement.click();
  }

  /** Mở .docx, .html hoặc .txt. */
  async loadFile(file: File | Blob, name = file instanceof File ? file.name : 'document.docx') {
    await this.ready;
    if (this.tab() === 'file') this.tab.set('home'); // đóng màn hình Tệp để thấy ngay vùng soạn thảo
    await this.run('Đang mở tài liệu…', async () => {
      const lower = name.toLowerCase();
      const title = name.replace(/\.(docx|html?|txt)$/i, '');
      if (lower.endsWith('.docx') || file.type.includes('wordprocessingml')) {
        const { importDocx } = await import('./core/docx-import');
        const d = await importDocx(file);
        this.engine!.setDocument({ ...d, title });
      } else {
        const text = await file.text();
        const html = lower.endsWith('.txt')
          ? text.split(/\r?\n/).map(l => `<p>${l.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!) || '<br>'}</p>`).join('')
          : (new DOMParser().parseFromString(text, 'text/html').body.innerHTML);
        this.engine!.setDocument({ html, header: '', footer: '', firstHeader: '', firstFooter: '', differentFirstPage: false, title });
      }
      this.syncFromEngine();
      await nextFrame(); // chờ màn hình Tệp gỡ khỏi DOM rồi mới focus
      this.engine!.focusStart();
      this.fileLoaded.emit(name);
    });
  }

  newDocument() {
    if (this.layout().words && !confirm('Tạo tài liệu mới? Nội dung chưa lưu sẽ mất.')) return;
    this.engine!.setDocument({
      html: '',
      header: '',
      footer: '',
      firstHeader: '',
      firstFooter: '',
      differentFirstPage: false,
      title: 'Tài liệu',
      page: WordEditorDefaults.page(),
    });
    this.syncFromEngine();
    this.tab.set('home');
  }

  // ---------- ControlValueAccessor ----------

  writeValue(v: string | null) {
    this.applyHtmlInput(v ?? '');
  }
  registerOnChange(fn: (html: string) => void) {
    this.onChange = fn;
  }
  registerOnTouched(fn: () => void) {
    this.onTouched = fn;
  }
  setDisabledState(d: boolean) {
    this.disabled = d;
    this.applyEditable(this.readonly());
  }

  // =============================================================================================
  // Template handlers
  // =============================================================================================

  protected exec(cmd: string, value?: unknown) {
    this.menu.set(null);
    this.engine?.exec(cmd, value);
  }

  protected toggleMenu(name: string, e?: Event) {
    e?.stopPropagation();
    this.menu.update(m => (m === name ? null : name));
  }

  protected selectTab(t: Tab) {
    this.tab.set(t);
    this.menu.set(null);
  }

  /** Giữ vùng chọn trong tài liệu khi bấm nút ribbon (trừ ô nhập liệu). */
  protected keepFocus(e: MouseEvent) {
    const t = e.target as HTMLElement;
    if (!t.closest('input, select, textarea, label')) e.preventDefault();
  }

  protected applyFont(name: string) {
    if (name.trim()) this.exec('fontName', name.trim());
  }

  protected applySize(v: string | number) {
    const n = parseFloat(String(v).replace(',', '.'));
    if (n > 0 && n <= 1638) this.exec('fontSize', n);
  }

  protected setFontColor(hex: string) {
    this.fontColor.set(hex);
    this.exec('foreColor', hex);
  }
  protected setHighlight(hex: string) {
    if (hex !== 'transparent') this.highlightColor.set(hex);
    this.exec('hiliteColor', hex);
  }
  protected setShading(hex: string, cell = false) {
    if (hex !== 'transparent') this.shadingColor.set(hex);
    this.exec(cell ? 'cellShading' : 'paragraphShading', hex);
  }

  protected insertTable(rows: number, cols: number) {
    this.menu.set(null);
    this.dialog.set(null);
    this.exec('insertTable', { rows, cols });
  }

  protected pickImage() {
    this.imageInput().nativeElement.click();
  }

  protected async onImagePicked(e: Event) {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) await this.engine!.insertImage(await readAsDataUrl(file));
  }

  protected onFilePicked(e: Event) {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) void this.loadFile(file);
  }

  protected async insertImageUrl() {
    const url = this.imageUrl().trim();
    this.dialog.set(null);
    if (url) await this.engine!.insertImage(url);
  }

  protected openLinkDialog() {
    const sel = getSelection();
    const a = sel?.anchorNode?.parentElement?.closest('a');
    this.linkForm.set({ text: a?.textContent ?? sel?.toString() ?? '', href: a?.getAttribute('href') ?? 'https://' });
    this.dialog.set('link');
  }

  protected applyLink() {
    const { href, text } = this.linkForm();
    this.dialog.set(null);
    if (href.trim() && href.trim() !== 'https://') this.exec('insertLink', { href: href.trim(), text });
  }

  protected openFind(replace = false) {
    const sel = getSelection()?.toString() ?? '';
    this.findForm.update(f => ({ ...f, showReplace: replace, query: sel && sel.length < 100 ? sel : f.query }));
    this.dialog.set('find');
    queueMicrotask(() => {
      (document.querySelector('.we-find input') as HTMLInputElement | null)?.select();
      this.runFind();
    });
  }

  protected runFind() {
    const f = this.findForm();
    this.findResult.set(this.engine!.find(f.query, f));
  }

  protected findNext(back = false) {
    if (!this.findResult()) this.runFind();
    const cur = this.engine!.findNext(back);
    this.findResult.update(r => (r ? { ...r, current: cur } : r));
  }

  protected replaceOne() {
    const f = this.findForm();
    this.findResult.set(this.engine!.replace(f.query, f.replace, f));
  }

  protected replaceAll() {
    const f = this.findForm();
    const n = this.engine!.replaceAll(f.query, f.replace, f);
    this.findResult.set({ count: 0, current: -1 });
    this.showToast(`Đã thay ${n} chỗ.`);
  }

  protected closeFind() {
    this.engine?.clearFind();
    this.findResult.set(null);
    this.dialog.set(null);
    this.engine?.focus();
  }

  protected openMargins() {
    const p = this.page();
    this.marginForm.set({ t: toCm(p.marginTop), b: toCm(p.marginBottom), l: toCm(p.marginLeft), r: toCm(p.marginRight), h: toCm(p.header), f: toCm(p.footer) });
    this.menu.set(null);
    this.dialog.set('margins');
  }

  protected applyMargins() {
    const m = this.marginForm();
    this.setPage({ marginTop: cm(m.t), marginBottom: cm(m.b), marginLeft: cm(m.l), marginRight: cm(m.r), header: cm(m.h), footer: cm(m.f) });
    this.dialog.set(null);
  }

  protected setMarginPreset(key: string) {
    const m = MARGIN_PRESETS[key];
    this.setPage({ marginTop: cm(m.t), marginRight: cm(m.r), marginBottom: cm(m.b), marginLeft: cm(m.l) });
  }

  protected setPaper(key: string) {
    const s = PAPER_SIZES[key];
    this.setPage(this.landscape() ? { width: s.h, height: s.w } : { width: s.w, height: s.h });
  }

  protected setOrientation(landscape: boolean) {
    const p = this.page();
    if (landscape !== this.landscape()) this.setPage({ width: p.height, height: p.width });
  }

  protected setPage(p: Partial<PageSetup>) {
    this.menu.set(null);
    this.engine!.setPage(p);
    this.page.set(this.engine!.page);
  }

  protected setZoom(z: number | 'width' | 'page') {
    this.engine!.setZoom(z);
    this.zoom.set(this.engine!.getZoom());
    this.updateRuler();
  }

  protected setView(v: 'print' | 'web') {
    this.engine!.setView(v);
    this.view.set(v);
  }

  protected toggleMarks() {
    this.marks.update(v => !v);
    this.engine!.setShowMarks(this.marks());
  }

  protected toggleGridlines() {
    this.gridlines.update(v => !v);
    this.engine!.setGridlines(this.gridlines());
  }

  protected setDifferentFirst(on: boolean) {
    this.differentFirst.set(on);
    this.exec('differentFirstPage', on);
  }

  protected pageNumber(where: HFKind, align: string, format: string) {
    this.exec('pageNumber', { where, align, format });
  }

  protected setTitle(t: string) {
    this.title.set(t);
    this.engine?.setTitle(t);
  }

  protected openAlt() {
    const img = document.querySelector('.we-img-sel') as HTMLImageElement | null;
    this.altText.set(img?.alt ?? '');
    this.dialog.set('alt');
  }

  protected applyAlt() {
    this.dialog.set(null);
    this.exec('imageAlt', this.altText());
  }

  protected async paste(plain = false) {
    this.menu.set(null);
    const ok = await this.engine!.pasteFromClipboard(plain);
    if (!ok) this.showToast('Trình duyệt chặn đọc clipboard. Hãy dùng Ctrl+V.');
  }

  protected copyCut(cmd: 'copy' | 'cut') {
    this.exec(cmd);
  }

  protected onSave() {
    this.menu.set(null);
    if (this.tab() === 'file') this.tab.set('home');
    const doc = this.engine!.getDocument();
    this.save.emit(doc);
    if (this.saveMode() === 'download') void this.downloadDocx();
  }

  protected onHostKey(e: KeyboardEvent) {
    if (e.key === 'Escape') {
      if (this.menu()) this.menu.set(null);
      else if (this.dialog() === 'find') this.closeFind();
      else if (this.dialog()) this.dialog.set(null);
      else if (this.tab() === 'file') this.tab.set('home');
    }
    if (e.key === 'F7' || ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'n' && !e.shiftKey)) {
      // Ctrl+N mở tài liệu mới như Word (trình duyệt có thể đã chặn phím này).
      e.preventDefault();
      this.newDocument();
    }
  }

  // ----- Thước kẻ -----

  protected startRulerDrag(e: MouseEvent, what: 'marginLeft' | 'marginRight' | 'first' | 'hanging' | 'indentRight') {
    e.preventDefault();
    e.stopPropagation();
    const z = this.zoom();
    const startX = e.clientX;
    const p0 = this.page();
    const s0 = this.state();
    const left0 = s0?.indentLeft ?? 0;
    const first0 = s0?.indentFirst ?? 0;
    const right0 = s0?.indentRight ?? 0;
    const snap = (px: number) => Math.round(px / cm(0.25)) * cm(0.25);
    const move = (ev: MouseEvent) => {
      const dx = snap((ev.clientX - startX) / z);
      const cw = p0.width - p0.marginLeft - p0.marginRight;
      switch (what) {
        case 'marginLeft':
          this.engine!.setPage({ marginLeft: Math.max(0, Math.min(p0.width - p0.marginRight - cm(1), p0.marginLeft + dx)) });
          break;
        case 'marginRight':
          this.engine!.setPage({ marginRight: Math.max(0, Math.min(p0.width - p0.marginLeft - cm(1), p0.marginRight - dx)) });
          break;
        case 'first':
          this.engine!.exec('paragraphIndent', { first: Math.max(-left0, Math.min(cw - left0, first0 + dx)) });
          break;
        case 'hanging':
          this.engine!.exec('paragraphIndent', { left: Math.max(0, left0 + dx), first: first0 - dx });
          break;
        case 'indentRight':
          this.engine!.exec('paragraphIndent', { right: Math.max(0, right0 - dx) });
          break;
      }
      this.page.set(this.engine!.page);
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  // =============================================================================================
  // Nội bộ
  // =============================================================================================

  private onRequest(r: string) {
    switch (r) {
      case 'link':
        return this.openLinkDialog();
      case 'find':
        return this.openFind(false);
      case 'replace':
        return this.openFind(true);
      case 'save':
        return this.onSave();
      case 'open':
        return this.openFilePicker();
      case 'print':
        return this.print();
    }
  }

  // ---------- Input / output ----------

  private applyHtmlInput(html: string) {
    if (html === this.lastHtml || (this.engine && html === this.engine.getHtml())) return;
    this.lastHtml = html;
    this.setHtml(html);
  }

  private applyEditable(readonly: boolean) {
    const locked = readonly || this.disabled;
    this.host.classList.toggle('we-readonly', locked);
    this.engine?.setReadonly(locked);
  }

  /** Gom các thay đổi liên tiếp, chỉ phát sau khi người dùng ngừng sửa `debounce` ms. */
  private scheduleEmit() {
    if (this.emitMode() === 'manual') return;
    clearTimeout(this.emitTimer);
    this.emitTimer = window.setTimeout(() => void this.flushEmit(), this.debounce());
    if (this.emit().includes('docxBase64')) {
      clearTimeout(this.docxTimer);
      this.docxTimer = window.setTimeout(() => void this.emitDocx(), this.docxDebounce());
    }
  }

  /** Phát ngay các output trong `emit` (hoặc `kinds`), kể cả khi `emitMode` là 'manual'. Hiện thanh tiến trình khi đang tạo. */
  async emitNow(kinds: WordEmitKind[] = this.emit()) {
    await this.ready;
    clearTimeout(this.emitTimer);
    clearTimeout(this.docxTimer);
    await this.run('Đang xuất dữ liệu…', async () => {
      await this.flushEmit(kinds, true);
      if (kinds.includes('docxBase64')) this.docxBase64Change.emit(await this.getDocxBase64());
    });
  }

  private async flushEmit(kinds: WordEmitKind[] = this.emit(), force = false) {
    if (!this.engine) return;
    const html = this.engine.getHtml();
    const changed = force || html !== this.lastHtml;
    this.lastHtml = html;
    if (changed) this.onChange(html);
    if (kinds.includes('html') && changed) this.htmlChange.emit(html);
    if (kinds.includes('document')) this.documentChange.emit(this.engine.getDocument());
    if (kinds.includes('htmlBase64') && changed) this.htmlBase64Change.emit(await textToBase64(html));
  }

  private async emitDocx() {
    const seq = ++this.docxSeq;
    try {
      const b64 = await this.getDocxBase64();
      if (seq === this.docxSeq) this.docxBase64Change.emit(b64); // bỏ kết quả cũ nếu đã có lần sửa mới hơn
    } catch (err) {
      this.failed.emit(err);
    }
  }

  private syncFromEngine() {
    const e = this.engine!;
    this.page.set(e.page);
    this.title.set(e.title);
    this.differentFirst.set(e.differentFirstPage);
    this.zoom.set(e.getZoom());
    this.view.set(e.getView());
  }

  private updateRuler() {
    const canvas = this.canvas().nativeElement;
    const stage = canvas.querySelector('.we-stage') as HTMLElement | null;
    if (!stage) return;
    this.rulerLeft.set(stage.getBoundingClientRect().left - canvas.getBoundingClientRect().left);
  }

  private fileName() {
    return (this.title() || 'tai-lieu').replace(/[\\/:*?"<>|]+/g, '_');
  }

  private async run<T>(message: string, task: () => Promise<T>): Promise<T | undefined> {
    this.busy.set(message);
    try {
      await nextFrame(); // để thanh "đang xử lý" kịp vẽ trước khi main thread bị chiếm
      return await task();
    } catch (err) {
      console.error(err);
      this.failed.emit(err);
      this.showToast('Có lỗi xảy ra. Kiểm tra lại file hoặc thử lại.');
      return undefined;
    } finally {
      this.busy.set(null);
    }
  }

  protected showToast(msg: string) {
    this.toast.set(msg);
    clearTimeout(this.toastTimer);
    this.toastTimer = window.setTimeout(() => this.toast.set(null), 4000);
  }

  protected hex(c: string) {
    return `#${c}`;
  }

  /** Cập nhật một trường của form trong hộp thoại (template không dùng được arrow function). */
  protected patch<T extends object, K extends keyof T>(s: WritableSignal<T>, key: K, value: T[K]) {
    s.update(v => ({ ...v, [key]: value }));
  }
}

/** Mặc định tách riêng để app có thể đọc khi cần. */
export const WordEditorDefaults = {
  page: (): PageSetup => ({
    width: PAPER_SIZES['A4'].w,
    height: PAPER_SIZES['A4'].h,
    marginTop: cm(2),
    marginRight: cm(2),
    marginBottom: cm(2),
    marginLeft: cm(3),
    header: cm(1.25),
    footer: cm(1.25),
  }),
};

function toDocxBlob(src: DocxSource): Blob {
  if (src instanceof Blob) return src;
  if (typeof src === 'string') return base64ToBlob(src.replace(/^data:[^,]*,/, ''), DOCX_MIME);
  return new Blob([src as BlobPart], { type: DOCX_MIME });
}

/** Chờ trình duyệt vẽ một khung hình (có hạn chờ vì rAF bị dừng khi tab không hiển thị). */
function nextFrame(): Promise<void> {
  return new Promise(resolve => {
    const done = () => resolve();
    requestAnimationFrame(() => setTimeout(done, 0));
    setTimeout(done, 80);
  });
}
