import { PAGE_BREAK_HTML } from './dom';
import { ListItem, buildLists } from './lists';
import { DEFAULT_PAGE, PageSetup, naturalLineFactor, twipToPx } from './units';

export interface HeaderFooter {
  header: string;
  footer: string;
  firstHeader: string;
  firstFooter: string;
  differentFirstPage: boolean;
}

export interface ImportedDocx extends HeaderFooter {
  html: string;
  page: PageSetup;
}

export { PAGE_BREAK_HTML } from './dom';

/** Thuộc tính kế thừa: chỉ ghi khi khác phần tử cha (khối cấp cao nhất so với mặc định của editor). */
const INHERITED = [
  'font-family', 'font-size', 'font-weight', 'font-style', 'color', 'text-align', 'line-height',
  'text-indent', 'letter-spacing', 'white-space', 'border-collapse',
] as const;

/** Thuộc tính không kế thừa: chỉ ghi khi khác giá trị mặc định. */
const OWN = [
  'background-color', 'text-decoration-line', 'vertical-align',
  'margin-top', 'margin-right', 'margin-bottom', 'margin-left',
  'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
  'border-top', 'border-right', 'border-bottom', 'border-left',
] as const;

const SIZED = new Set(['TABLE', 'TD', 'TH', 'COL', 'IMG']);
const DEFAULTS = new Set(['', 'none', 'auto', 'normal', 'start', 'baseline', '0px', 'rgba(0, 0, 0, 0)', 'transparent', 'separate']);

const isDefault = (prop: string, value: string) =>
  DEFAULTS.has(value) || (prop.startsWith('border-') && (value.includes('none') || value.startsWith('0px')));

const formatValue = (prop: string, value: string) =>
  prop === 'font-size' && value.endsWith('px') ? `${Math.round(parseFloat(value) * 0.75 * 10) / 10}pt` : value;

/**
 * Giãn dòng "multiple" của Word (1,15; 1,5…) được trình duyệt tính ra px. Nếu px đúng là bội số của
 * cỡ chữ thì ghi lại dạng tỉ lệ: co giãn theo cỡ chữ như Word và xuất ra .docx thành `lineRule="auto"`.
 */
function lineHeightValue(value: string, fontSize: string | undefined, fontFamily: string | undefined): string {
  const px = parseFloat(value);
  const fs = parseFloat(fontSize ?? '');
  if (!value.endsWith('px') || !fs) return value;
  const ratio = px / fs;
  const rounded = Math.round(ratio * 100) / 100;
  // Word: multiple × chiều cao dòng đơn của phông, CSS: multiple × cỡ chữ → nhân thêm hệ số dòng đơn của phông.
  return ratio >= 0.5 && ratio <= 5 && Math.abs(ratio - rounded) < 0.0005 ? String(Math.round(rounded * naturalLineFactor(fontFamily) * 1000) / 1000) : value;
}

/** Font mặc định phải trùng với `.we-doc` trong content-css.ts. */
const BASELINE_CSS = "font-family:'Times New Roman',Times,serif;font-size:12pt;line-height:normal;color:#000;text-align:start";

/**
 * Chuyển .docx thành dữ liệu cho editor.
 * - Thân văn bản: docx-preview dựng lại định dạng bằng CSS class; ta đọc computed style và ghi inline.
 * - Header/footer, số trang, khổ giấy: đọc trực tiếp XML (docx-preview chỉ render header trang đầu
 *   và bỏ mất trường PAGE/NUMPAGES).
 */
export async function importDocx(file: Blob): Promise<ImportedDocx> {
  const [{ renderAsync }, { default: JSZip }] = await Promise.all([import('docx-preview'), import('jszip')]);
  const zip = await JSZip.loadAsync(file);

  const box = document.createElement('div');
  box.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;' + BASELINE_CSS;
  document.body.append(box);

  try {
    await renderAsync(file, box, box, {
      inWrapper: false,
      breakPages: true, // mỗi đoạn ngắt trang thủ công thành một <section>, ta nối lại bằng PAGE_BREAK_HTML
      ignoreLastRenderedPageBreak: true,
      renderHeaders: false,
      renderFooters: false,
      renderComments: false,
      renderChanges: false,
      useBase64URL: true,
      experimental: true,
    });

    const sections = Array.from(box.querySelectorAll<HTMLElement>('section.docx'));
    const articles = sections.map(s => s.querySelector('article') ?? s);
    const docXml = await readXml(zip, 'word/document.xml');
    const tables = articles.flatMap(a => Array.from(a.querySelectorAll('table')));

    markKeep(articles, docXml, await readXml(zip, 'word/styles.xml'));
    const headings = readHeadings(articles);
    const lists = readLists(articles);
    inlineStyles(articles, box);
    applyTableXml(tables, docXml);
    applyHeadings(headings);
    buildLists(lists);
    unwrapInlineDivs(articles);
    hoistRunStyles(articles);
    fillEmptyBlocks(articles);
    applyKeep(articles);
    dropBreakRemainders(articles, docXml);

    const html = articles.map(a => a.innerHTML).join(PAGE_BREAK_HTML);
    const { page, hf } = await readSectionXml(zip, sections[0], docXml);
    return { html, page, ...hf };
  } finally {
    box.remove();
  }
}

// ---------- Inline computed style ----------

function readStyle(el: Element) {
  const cs = getComputedStyle(el);
  const values = new Map<string, string>();
  for (const p of [...INHERITED, ...OWN]) values.set(p, cs.getPropertyValue(p));
  if (SIZED.has(el.tagName)) {
    values.set('width', cs.width);
    if (el.tagName === 'IMG') values.set('height', cs.height);
  }
  // Viền: computed làm tròn 0,5pt thành 1px; giữ độ dày gốc (pt) nếu docx-preview có ghi inline.
  const inline = (el as HTMLElement).style;
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const authored = inline?.getPropertyValue(`border-${side}-width`);
    const v = values.get(`border-${side}`);
    if (authored && v && /\d(pt|cm|mm|in)$/.test(authored)) values.set(`border-${side}`, v.replace(/^[\d.]+px/, authored));
  }
  return values;
}

function inlineStyles(roots: HTMLElement[], baseline: HTMLElement) {
  const rootSet = new Set<Element>(roots);
  const elements = roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('*'))).filter(el => el instanceof HTMLElement);

  // Đọc hết rồi mới ghi: xoá class sẽ làm computed style của phần tử con thay đổi.
  const base = readStyle(baseline);
  const snapshot = new Map<Element, Map<string, string>>(elements.map(el => [el, readStyle(el)]));

  for (const el of elements) {
    const values = snapshot.get(el)!;
    const parent = el.parentElement && !rootSet.has(el.parentElement) ? snapshot.get(el.parentElement)! : base;
    const decl: string[] = [];
    for (const [prop, value] of values) {
      const inherited = (INHERITED as readonly string[]).includes(prop);
      const skip = inherited ? parent.get(prop) === value : isDefault(prop, value);
      if (!skip) decl.push(`${prop}: ${prop === 'line-height' ? lineHeightValue(value, values.get('font-size'), values.get('font-family')) : formatValue(prop, value)}`);
    }
    el.removeAttribute('class');
    if (decl.length) el.setAttribute('style', decl.join('; '));
    else el.removeAttribute('style');
  }

  for (const el of elements) {
    if (el.tagName === 'SPAN' && el.attributes.length === 0 && el.parentNode) el.replaceWith(...Array.from(el.childNodes));
  }
}

/**
 * docx-preview bọc ảnh (drawing) trong <div> nằm bên trong <p>. Khi HTML được phân tích lại, trình
 * duyệt tách thành <p></p><div><img></div><p></p>: ảnh mất căn lề của đoạn và sinh thêm 2 đoạn trống.
 * Gỡ lớp <div> để ảnh nằm thẳng trong đoạn như trong Word.
 */
function unwrapInlineDivs(roots: HTMLElement[]) {
  const sel = ['p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'span'].map(t => `${t} div`).join(', ');
  for (const div of roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>(sel)))) {
    if (!div.isConnected || div.querySelector('p, table, ul, ol, h1, h2, h3, h4, h5, h6, div')) continue;
    div.replaceWith(...Array.from(div.childNodes));
  }
}

/**
 * Nếu mọi chữ trong một đoạn có cùng cỡ chữ/phông (khai báo ở <span>), đưa lên chính đoạn đó.
 * Trình duyệt tính chiều cao dòng theo phông của khối chứa; Word tính theo các run và dấu đoạn văn,
 * và docx-preview không áp dụng định dạng dấu đoạn văn (pPr/rPr) lên <p>.
 */
function hoistRunStyles(roots: HTMLElement[]) {
  const blocks = roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('p, li, h1, h2, h3, h4, h5, h6')));
  for (const block of blocks) {
    for (const prop of ['font-size', 'font-family']) {
      const values = new Set<string>();
      const owners = new Set<HTMLElement>();
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (!n.textContent?.trim()) continue;
        let owner: HTMLElement | null = null;
        for (let e = n.parentElement; e && e !== block; e = e.parentElement) {
          if (/^(P|LI|H\d|TD|TH|TABLE|UL|OL)$/.test(e.tagName)) {
            owner = null;
            values.add('\u0000'); // chữ thuộc khối con: không xử lý
            break;
          }
          if (e.style.getPropertyValue(prop)) {
            owner = e;
            break;
          }
        }
        values.add(owner ? owner.style.getPropertyValue(prop) : block.style.getPropertyValue(prop) || '\u0001');
        if (owner) owners.add(owner);
      }
      // Đoạn trống: lấy định dạng từ <span> rỗng (run không có chữ) nếu có.
      if (!values.size) {
        for (const s of Array.from(block.querySelectorAll<HTMLElement>('span[style]'))) {
          const v = s.style.getPropertyValue(prop);
          if (v) {
            values.add(v);
            owners.add(s);
          }
        }
      }
      if (values.size !== 1 || !owners.size) continue;
      const [value] = values;
      if (value.startsWith('\u0000') || value === '\u0001') continue;
      block.style.setProperty(prop, value);
      for (const o of owners) {
        o.style.removeProperty(prop);
        if (!o.getAttribute('style')) o.removeAttribute('style');
      }
    }
  }
}

function fillEmptyBlocks(roots: HTMLElement[]) {
  for (const el of roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('p, li, h1, h2, h3, h4, h5, h6')))) {
    if (!el.textContent && !el.querySelector('img, br, table, svg')) el.innerHTML = '<br>';
  }
}

// ---------- Heading / Title ----------

interface HeadingInfo {
  tag: string;
  cls?: string;
  /** Giá trị computed cần ghi rõ, vì CSS mặc định của h1… trong editor khác với style trong file. */
  fixed: Record<string, string>;
}

function readHeadings(roots: HTMLElement[]) {
  const out = new Map<HTMLElement, HeadingInfo>();
  for (const p of roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('p[class]')))) {
    const c = p.className;
    const h = /(?:^|\s)\S*_heading(\d)(?:\s|$)/.exec(c)?.[1];
    const info: Pick<HeadingInfo, 'tag' | 'cls'> | null = h
      ? { tag: `h${Math.min(6, Number(h))}` }
      : /(?:^|\s)\S*_title(?:\s|$)/.test(c)
        ? { tag: 'p', cls: 'we-title' }
        : /(?:^|\s)\S*_subtitle(?:\s|$)/.test(c)
          ? { tag: 'p', cls: 'we-subtitle' }
          : null;
    if (!info) continue;
    const cs = getComputedStyle(p);
    out.set(p, {
      ...info,
      fixed: {
        'margin-top': cs.marginTop,
        'margin-bottom': cs.marginBottom,
        'font-size': formatValue('font-size', cs.fontSize),
        'font-weight': cs.fontWeight,
        'font-style': cs.fontStyle,
        color: cs.color,
        'letter-spacing': cs.letterSpacing,
      },
    });
  }
  return out;
}

function applyHeadings(headings: Map<HTMLElement, HeadingInfo>) {
  for (const [p, info] of headings) {
    if (!p.parentNode) continue;
    const el = info.tag === 'p' ? p : document.createElement(info.tag);
    if (el !== p) {
      for (const a of Array.from(p.attributes)) el.setAttribute(a.name, a.value);
      el.append(...Array.from(p.childNodes));
      p.replaceWith(el);
    }
    for (const [prop, value] of Object.entries(info.fixed)) if (!el.style.getPropertyValue(prop)) el.style.setProperty(prop, value);
    if (info.cls) el.classList.add(info.cls);
  }
}

// ---------- Danh sách ----------
// docx-preview vẽ số/bullet bằng ::before + counter trên <p class="docx-num-{id}-{level}">.

const BULLETS: Record<string, string> = { '○': 'circle', o: 'circle', '■': 'square', '▪': 'square', '§': 'square' };

function readLists(roots: HTMLElement[]): Map<HTMLElement, ListItem> {
  const items = new Map<HTMLElement, ListItem>();
  for (const p of roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('p[class*="-num-"]')))) {
    const m = /(?:^|\s)\S*-num-(\d+)-(\d+)(?:\s|$)/.exec(p.className);
    const content = getComputedStyle(p, '::before').content;
    if (!m || !content || content === 'none') continue;
    const counters = Array.from(content.matchAll(/counter\([^,)]+(?:,\s*([\w-]+))?\)/g));
    const bullet = /"([^"\\]*)/.exec(content)?.[1]?.trim() ?? '';
    items.set(p, {
      key: m[1],
      level: Number(m[2]),
      ordered: counters.length > 0,
      styleType: counters.length ? (counters.at(-1)![1] ?? 'decimal') : (BULLETS[bullet] ?? 'disc'),
    });
  }
  return items;
}

// ---------- Section: khổ giấy, lề, header/footer (đọc XML) ----------

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

type Zip = { file(path: string): { async(type: 'string' | 'base64'): Promise<string> } | null };

const wAttr = (el: Element | null | undefined, name: string) => el?.getAttributeNS(W, name) ?? el?.getAttribute(`w:${name}`) ?? null;
const child = (el: Element | null | undefined, name: string) =>
  el ? (Array.from(el.children).find(c => c.localName === name) ?? null) : null;
const children = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter(c => c.localName === name) : []);

async function readXml(zip: Zip, path: string) {
  const text = await zip.file(path)?.async('string');
  return text ? new DOMParser().parseFromString(text, 'application/xml') : null;
}

async function readRels(zip: Zip, path: string) {
  const dir = path.replace(/[^/]+$/, '');
  const name = path.slice(dir.length);
  const doc = await readXml(zip, `${dir}_rels/${name}.rels`);
  const map = new Map<string, string>();
  for (const r of Array.from(doc?.getElementsByTagName('Relationship') ?? [])) {
    const target = r.getAttribute('Target') ?? '';
    map.set(r.getAttribute('Id') ?? '', r.getAttribute('TargetMode') === 'External' ? target : resolvePath(dir, target));
  }
  return map;
}

function resolvePath(dir: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dir + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

async function readSectionXml(zip: Zip, firstSection: HTMLElement | undefined, doc: Document | null) {
  const page: PageSetup = { ...DEFAULT_PAGE };
  const hf: HeaderFooter = { header: '', footer: '', firstHeader: '', firstFooter: '', differentFirstPage: false };

  const sects = doc ? Array.from(doc.getElementsByTagNameNS(W, 'sectPr')) : [];
  const sect = sects.at(-1);
  if (!sect) {
    if (firstSection) Object.assign(page, measureSection(firstSection));
    return { page, hf };
  }

  const pgSz = child(sect, 'pgSz');
  const pgMar = child(sect, 'pgMar');
  const tw = (el: Element | null, n: string, fallback: number) => {
    const v = wAttr(el, n);
    return v !== null && !isNaN(Number(v)) ? twipToPx(Number(v)) : fallback;
  };
  page.width = tw(pgSz, 'w', page.width);
  page.height = tw(pgSz, 'h', page.height);
  page.marginTop = Math.abs(tw(pgMar, 'top', page.marginTop));
  page.marginRight = tw(pgMar, 'right', page.marginRight);
  page.marginBottom = Math.abs(tw(pgMar, 'bottom', page.marginBottom));
  page.marginLeft = tw(pgMar, 'left', page.marginLeft);
  page.header = tw(pgMar, 'header', page.header);
  page.footer = tw(pgMar, 'footer', page.footer);

  const titlePg = child(sect, 'titlePg');
  hf.differentFirstPage = !!titlePg && !['0', 'false'].includes(wAttr(titlePg, 'val') ?? '');

  const rels = await readRels(zip, 'word/document.xml');
  for (const kind of ['header', 'footer'] as const) {
    for (const ref of children(sect, `${kind}Reference`)) {
      const type = wAttr(ref, 'type') ?? 'default';
      const id = ref.getAttributeNS(R, 'id') ?? ref.getAttribute('r:id') ?? '';
      const path = rels.get(id);
      if (!path || (type !== 'default' && type !== 'first')) continue;
      const html = await partToHtml(zip, path);
      if (type === 'default') hf[kind] = html;
      else hf[kind === 'header' ? 'firstHeader' : 'firstFooter'] = html;
    }
  }
  return { page, hf };
}

function measureSection(section: HTMLElement): Partial<PageSetup> {
  const cs = getComputedStyle(section);
  return {
    width: parseFloat(cs.width) || DEFAULT_PAGE.width,
    height: parseFloat(cs.minHeight) || DEFAULT_PAGE.height,
    marginTop: parseFloat(cs.paddingTop) || DEFAULT_PAGE.marginTop,
    marginRight: parseFloat(cs.paddingRight) || DEFAULT_PAGE.marginRight,
    marginBottom: parseFloat(cs.paddingBottom) || DEFAULT_PAGE.marginBottom,
    marginLeft: parseFloat(cs.paddingLeft) || DEFAULT_PAGE.marginLeft,
  };
}

// ---------- Bảng: lấy lại thông số chính xác từ XML ----------

/**
 * docx-preview luôn render bảng với `table-layout: auto` nên độ rộng cột co giãn theo nội dung,
 * khác với Word (Word dùng đúng các cột trong `w:tblGrid`). Ở đây đọc lại tblGrid, căn bảng
 * và hàng tiêu đề lặp lại (`w:tblHeader`) từ XML. Bảng HTML và `w:tbl` được ghép theo thứ tự
 * trong tài liệu; nếu số lượng/số hàng không khớp thì giữ nguyên kết quả của docx-preview.
 */
function applyTableXml(tables: HTMLTableElement[], doc: Document | null) {
  if (!doc) return;
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  const xmlTables = body ? Array.from(body.getElementsByTagNameNS(W, 'tbl')) : [];
  if (xmlTables.length !== tables.length) return;

  tables.forEach((table, i) => {
    const x = xmlTables[i];
    const xmlRows = children(x, 'tr');
    const rows = Array.from(table.rows).filter(r => r.closest('table') === table);
    if (xmlRows.length !== rows.length) return;

    const grid = children(child(x, 'tblGrid'), 'gridCol').map(g => Number(wAttr(g, 'w')) || 0);
    if (grid.length && grid.every(w => w > 0)) {
      const widths = grid.map(w => Math.round(twipToPx(w) * 100) / 100);
      let colgroup = table.querySelector(':scope > colgroup');
      if (!colgroup) {
        colgroup = document.createElement('colgroup');
        table.prepend(colgroup);
      }
      colgroup.replaceChildren(
        ...widths.map(w => {
          const col = document.createElement('col');
          col.style.width = `${w}px`;
          return col;
        }),
      );
      table.style.width = `${Math.round(widths.reduce((a, b) => a + b, 0) * 100) / 100}px`;
      table.style.tableLayout = 'fixed';
      for (const r of rows) for (const c of Array.from(r.cells)) c.style.removeProperty('width');
    }
    table.style.removeProperty('padding-left');
    table.style.removeProperty('padding-right');

    const tblPr = child(x, 'tblPr');
    // Bảng tự co giãn (không có tblLayout=fixed): Word nới cột theo nội dung ngắn nhất; editor vẫn vẽ theo tblGrid
    // nhưng ghi nhớ để khi xuất không ép `fixed` (ép sẽ làm cột hẹp bị ngắt chữ, hàng cao hơn bản gốc).
    if (wAttr(child(tblPr, 'tblLayout'), 'type') !== 'fixed') table.dataset['weAutofit'] = '1';
    const jc = wAttr(child(tblPr, 'jc'), 'val');
    if (jc === 'center') {
      table.style.marginLeft = 'auto';
      table.style.marginRight = 'auto';
    } else if (jc === 'right' || jc === 'end') {
      table.style.marginLeft = 'auto';
    }
    const ind = Number(wAttr(child(tblPr, 'tblInd'), 'w'));
    if (ind && jc !== 'center') table.style.marginLeft = `${Math.round(twipToPx(ind) * 100) / 100}px`;

    // Hàng không được ngắt qua trang (w:cantSplit) → `break-inside: avoid`; xuất lại thành cantSplit.
    xmlRows.forEach((xr, k) => {
      if (isOn(child(child(xr, 'trPr'), 'cantSplit'))) rows[k].style.breakInside = 'avoid';
    });

    // Hàng tiêu đề lặp lại ở đầu mỗi trang → <thead>.
    let headerRows = 0;
    while (headerRows < xmlRows.length) {
      const h = child(child(xmlRows[headerRows], 'trPr'), 'tblHeader');
      if (!h || ['0', 'false'].includes(wAttr(h, 'val') ?? '')) break;
      headerRows++;
    }
    if (headerRows && headerRows < rows.length && !table.tHead) {
      // docx-preview có thể đặt <tr> thẳng trong <table> (không có <tbody>).
      const thead = document.createElement('thead');
      const anchor = rows[0].parentElement === table ? rows[0] : rows[0].parentElement!;
      table.insertBefore(thead, anchor);
      for (const r of rows.slice(0, headerRows)) thead.append(r);
    }
  });
}

// ---------- Ngắt trang ----------

/**
 * Đoạn Word chỉ chứa `<w:br w:type="page"/>` được docx-preview tách thành hai đoạn trống: một ở cuối trang trước,
 * một ở đầu trang sau. Word chỉ có một đoạn (dấu đoạn văn nằm cùng dòng với dấu ngắt trang, không chiếm dòng ở trang sau),
 * nên bỏ đoạn trống đầu trang sau; nếu không mọi trang sau ngắt trang đều bị đẩy xuống một dòng.
 * Chỉ làm khi số ngắt trang trong XML khớp số section đã tách (không thì không đủ căn cứ để ghép).
 */
function dropBreakRemainders(roots: HTMLElement[], doc: Document | null) {
  if (!doc || roots.length < 2) return;
  const breaks = Array.from(doc.getElementsByTagNameNS(W, 'br')).filter(b => wAttr(b, 'type') === 'page');
  if (breaks.length !== roots.length - 1) return;
  breaks.forEach((br, k) => {
    let p: Element | null = br;
    while (p && p.localName !== 'p') p = p.parentElement;
    if (!p) return;
    const hasContentAfter = Array.from(p.getElementsByTagNameNS(W, '*')).some(n => {
      if (!(br.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING)) return false;
      if (n.localName === 't') return !!n.textContent;
      return ['drawing', 'pict', 'tab', 'br', 'sym'].includes(n.localName);
    });
    if (hasContentAfter) return;
    const head = roots[k + 1].firstElementChild;
    if (head?.tagName === 'P' && !head.querySelector('img') && !head.textContent?.trim()) head.remove();
  });
}

// ---------- Giữ đoạn với đoạn sau / giữ nguyên đoạn (keepNext, keepLines) ----------

const isOn = (el: Element | null | undefined) => !!el && !['0', 'false', 'off'].includes(wAttr(el, 'val') ?? '1');

/** keepNext / keepLines hiệu lực của một đoạn: khai báo trực tiếp, nếu không thì theo kiểu (chuỗi basedOn). */
function keepOf(p: Element, styles: Map<string, Element>, defaultStyle: string | undefined) {
  const pPr = child(p, 'pPr');
  const out: { next?: boolean; lines?: boolean } = {};
  for (const [name, key] of [['keepNext', 'next'], ['keepLines', 'lines']] as const) {
    const direct = child(pPr, name);
    if (direct) {
      out[key] = isOn(direct);
      continue;
    }
    let id: string | undefined = wAttr(child(pPr, 'pStyle'), 'val') ?? defaultStyle;
    for (let depth = 0; id && depth < 12; depth++) {
      const st = styles.get(id);
      const flag = child(child(st, 'pPr'), name);
      if (flag) {
        out[key] = isOn(flag);
        break;
      }
      id = wAttr(child(st, 'basedOn'), 'val') ?? undefined;
    }
  }
  return out;
}

/** Cỡ chữ (half-point) của dấu đoạn văn: pPr/rPr, rồi kiểu đoạn (chuỗi basedOn), rồi docDefaults. */
function markSize(p: Element, styles: Map<string, Element>, defaultStyle: string | undefined, docDefault: number): number {
  const direct = Number(wAttr(child(child(child(p, 'pPr'), 'rPr'), 'sz'), 'val'));
  if (direct > 0) return direct;
  let id: string | undefined = wAttr(child(child(p, 'pPr'), 'pStyle'), 'val') ?? defaultStyle;
  for (let depth = 0; id && depth < 12; depth++) {
    const st = styles.get(id);
    const sz = Number(wAttr(child(child(st, 'rPr'), 'sz'), 'val'));
    if (sz > 0) return sz;
    id = wAttr(child(st, 'basedOn'), 'val') ?? undefined;
  }
  return docDefault;
}

/**
 * docx-preview bỏ qua keepNext/keepLines, nên đọc từ XML rồi gắn tạm `data-we-keep` lên khối cấp cao nhất
 * tương ứng (ghép dãy đoạn/bảng của XML với dãy khối đã dựng bằng LCS). Đến cuối `applyKeep` đổi thành CSS inline.
 */
function markKeep(roots: HTMLElement[], doc: Document | null, stylesDoc: Document | null) {
  const body = doc?.getElementsByTagNameNS(W, 'body')[0];
  if (!body) return;
  const styles = new Map<string, Element>();
  let defaultStyle: string | undefined;
  for (const st of Array.from(stylesDoc?.getElementsByTagNameNS(W, 'style') ?? [])) {
    const id = wAttr(st, 'styleId');
    if (!id) continue;
    styles.set(id, st);
    if (wAttr(st, 'type') === 'paragraph' && wAttr(st, 'default') === '1') defaultStyle = id;
  }
  const docDefaultSz = Number(wAttr(child(child(child(child(stylesDoc?.documentElement, 'docDefaults'), 'rPrDefault'), 'rPr'), 'sz'), 'val')) || 0;
  const xml = Array.from(body.children).filter(c => c.localName === 'p' || c.localName === 'tbl');
  const rendered = roots.flatMap(r => Array.from(r.children) as HTMLElement[]).filter(e => e.tagName === 'P' || e.tagName === 'TABLE');
  const norm = (s: string) => s.replace(/\s+/g, '').slice(0, 16);
  const xmlText = (x: Element) => norm(Array.from(x.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join(''));

  // Ghép hai dãy bằng LCS theo (loại khối, 16 ký tự đầu): docx-preview chèn/bớt vài khối (ảnh, ngắt trang) nên không ghép thẳng theo chỉ số.
  const xs = xml.map(x => `${x.localName}:${x.localName === 'p' ? xmlText(x) : ''}`);
  const rs = rendered.map(e => `${e.tagName === 'TABLE' ? 'tbl' : 'p'}:${e.tagName === 'P' ? norm(e.textContent ?? '') : ''}`);
  const n = xs.length;
  const m = rs.length;
  const dp = new Uint16Array((n + 1) * (m + 1));
  const at = (i: number, j: number) => i * (m + 1) + j;
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--) dp[at(i, j)] = xs[i] === rs[j] ? dp[at(i + 1, j + 1)] + 1 : Math.max(dp[at(i + 1, j)], dp[at(i, j + 1)]);
  const pairs: [Element, HTMLElement][] = [];
  for (let i = 0, j = 0; i < n && j < m; ) {
    if (xs[i] === rs[j]) pairs.push([xml[i++], rendered[j++]]);
    else if (dp[at(i + 1, j)] >= dp[at(i, j + 1)]) i++;
    else j++;
  }
  // Đoạn trống: chiều cao do cỡ chữ của dấu đoạn văn (pPr/rPr/sz, rồi kiểu, rồi docDefaults) quyết định, docx-preview không áp dụng.
  const markEmpty = (x: Element, el: HTMLElement) => {
    const empty = !xmlText(x) && !x.getElementsByTagNameNS(W, 'drawing').length && !x.getElementsByTagNameNS(W, 'pict').length;
    const sz = empty ? markSize(x, styles, defaultStyle, docDefaultSz) : 0;
    if (sz > 0 && !el.textContent?.trim() && !el.querySelector('img')) el.setAttribute('data-we-sz', String(sz / 2));
  };
  for (const [x, el] of pairs) {
    if (x.localName === 'tbl') {
      // Ô bảng: ghép hàng, ô, đoạn theo thứ tự khi số lượng khớp.
      const xr = children(x, 'tr');
      const rr = Array.from((el as HTMLTableElement).rows).filter(r => r.closest('table') === el);
      if (xr.length !== rr.length) continue;
      xr.forEach((row, i) => {
        const xc = children(row, 'tc');
        const rc = Array.from(rr[i].cells);
        if (xc.length !== rc.length) return;
        xc.forEach((cell, j) => {
          const xp = children(cell, 'p');
          const rp = Array.from(rc[j].children).filter(c => c.tagName === 'P') as HTMLElement[];
          if (xp.length === rp.length) xp.forEach((pp, k) => markEmpty(pp, rp[k]));
        });
      });
      continue;
    }
    const k = keepOf(x, styles, defaultStyle);
    // Không khai báo keepNext = tắt (Word); chỉ cần ghi rõ cho đề mục, vì editor và kiểu Heading khi xuất mặc định giữ với đoạn sau.
    const v = `${k.next ? 'n' : 'x'}${k.lines ? 'l' : ''}`;
    if (v) el.setAttribute('data-we-keep', v);
    markEmpty(x, el);
  }
}

/** `data-we-keep` → CSS inline: `break-after: avoid` (keepNext), `auto` (tắt rõ ràng), `break-inside: avoid` (keepLines). */
function applyKeep(roots: HTMLElement[]) {
  for (const el of roots.flatMap(r => Array.from(r.querySelectorAll<HTMLElement>('[data-we-keep], [data-we-sz]')))) {
    const v = el.getAttribute('data-we-keep') ?? '';
    el.removeAttribute('data-we-keep');
    const sz = el.getAttribute('data-we-sz');
    el.removeAttribute('data-we-sz');
    if (sz && !el.textContent?.trim() && !el.querySelector('img')) el.style.fontSize = `${sz}pt`;
    if (v.includes('n')) el.style.breakAfter = 'avoid';
    else if (v.includes('x') && /^H[1-6]$/.test(el.tagName)) el.style.breakAfter = 'auto';
    if (v.includes('l')) el.style.breakInside = 'avoid';
  }
}

// ---------- Header/footer XML → HTML (đủ cho nội dung header thông thường) ----------

async function partToHtml(zip: Zip, path: string): Promise<string> {
  const doc = await readXml(zip, path);
  if (!doc?.documentElement) return '';
  const rels = await readRels(zip, path);
  const ctx = { zip, rels };
  const parts: string[] = [];
  for (const el of Array.from(doc.documentElement.children)) parts.push(await blockToHtml(el, ctx));
  const html = parts.join('');
  const probe = document.createElement('div');
  probe.innerHTML = html;
  return probe.textContent?.trim() || probe.querySelector('img, .we-field') ? html : '';
}

interface PartCtx {
  zip: Zip;
  rels: Map<string, string>;
}

async function blockToHtml(el: Element, ctx: PartCtx): Promise<string> {
  if (el.localName === 'p') return paragraphToHtml(el, ctx);
  if (el.localName === 'tbl') {
    const rows: string[] = [];
    for (const tr of children(el, 'tr')) {
      const cells: string[] = [];
      for (const tc of children(tr, 'tc')) {
        const span = wAttr(child(child(tc, 'tcPr'), 'gridSpan'), 'val');
        const inner: string[] = [];
        for (const b of Array.from(tc.children)) inner.push(await blockToHtml(b, ctx));
        cells.push(`<td${span ? ` colspan="${span}"` : ''}>${inner.join('') || '<p><br></p>'}</td>`);
      }
      rows.push(`<tr>${cells.join('')}</tr>`);
    }
    return `<table style="width: 100%; border-collapse: collapse"><tbody>${rows.join('')}</tbody></table>`;
  }
  if (el.localName === 'sdt') {
    const out: string[] = [];
    for (const b of Array.from(child(el, 'sdtContent')?.children ?? [])) out.push(await blockToHtml(b, ctx));
    return out.join('');
  }
  return '';
}

const JC: Record<string, string> = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

async function paragraphToHtml(p: Element, ctx: PartCtx): Promise<string> {
  const jc = wAttr(child(child(p, 'pPr'), 'jc'), 'val');
  const style = jc && JC[jc] ? ` style="text-align: ${JC[jc]}"` : '';
  const out: string[] = [];

  // Trường phức hợp: fldChar begin → instrText → separate → kết quả → end.
  let field: { instr: string; result: string[] } | null = null;

  const runs = Array.from(p.children).flatMap(c =>
    c.localName === 'hyperlink' || c.localName === 'smartTag' ? Array.from(c.children) : [c],
  );
  for (const node of runs) {
    if (node.localName === 'fldSimple') {
      const instr = wAttr(node, 'instr') ?? '';
      out.push(fieldHtml(instr) ?? (await runsText(node, ctx)));
      continue;
    }
    if (node.localName !== 'r') continue;

    // Một <w:r> có thể chứa cả fldChar, instrText và chữ (vd: file do thư viện docx tạo),
    // nên duyệt từng phần tử con thay vì coi cả run là một khối.
    let pending: Element[] = [];
    const flush = async () => {
      if (!pending.length) return;
      const html = await runToHtml(node, ctx, pending);
      pending = [];
      if (field) field.result.push(html);
      else out.push(html);
    };
    for (const c of Array.from(node.children)) {
      if (c.localName === 'rPr') continue;
      if (c.localName === 'fldChar') {
        await flush();
        const t = wAttr(c, 'fldCharType');
        if (t === 'begin') field = { instr: '', result: [] };
        else if (t === 'end' && field) {
          out.push(fieldHtml(field.instr) ?? field.result.join(''));
          field = null;
        }
      } else if (c.localName === 'instrText') {
        await flush();
        if (field) field.instr += c.textContent ?? '';
      } else {
        pending.push(c);
      }
    }
    await flush();
  }
  return `<p${style}>${out.join('') || '<br>'}</p>`;
}

function fieldHtml(instr: string): string | null {
  const name = instr.trim().split(/\s+/)[0]?.toUpperCase();
  if (name === 'PAGE') return '<span class="we-field" data-we-field="page">1</span>';
  if (name === 'NUMPAGES' || name === 'SECTIONPAGES') return '<span class="we-field" data-we-field="pages">1</span>';
  return null;
}

async function runsText(el: Element, ctx: PartCtx) {
  const out: string[] = [];
  for (const r of children(el, 'r')) out.push(await runToHtml(r, ctx));
  return out.join('');
}

async function runToHtml(r: Element, ctx: PartCtx, only?: Element[]): Promise<string> {
  const rPr = child(r, 'rPr');
  const css: string[] = [];
  const on = (name: string) => {
    const e = child(rPr, name);
    return !!e && !['0', 'false', 'none'].includes(wAttr(e, 'val') ?? '');
  };
  if (on('b')) css.push('font-weight: bold');
  if (on('i')) css.push('font-style: italic');
  const deco = [on('u') && 'underline', (on('strike') || on('dstrike')) && 'line-through'].filter(Boolean);
  if (deco.length) css.push(`text-decoration-line: ${deco.join(' ')}`);
  const sz = wAttr(child(rPr, 'sz'), 'val');
  if (sz) css.push(`font-size: ${Number(sz) / 2}pt`);
  const color = wAttr(child(rPr, 'color'), 'val');
  if (color && color !== 'auto') css.push(`color: #${color}`);
  const font = wAttr(child(rPr, 'rFonts'), 'ascii') ?? wAttr(child(rPr, 'rFonts'), 'hAnsi');
  if (font) css.push(`font-family: '${font.replace(/'/g, '')}'`);
  const hl = wAttr(child(rPr, 'highlight'), 'val');
  if (hl && hl !== 'none') css.push(`background-color: ${hl.replace(/^dark/, 'dark').toLowerCase()}`);
  const shd = wAttr(child(rPr, 'shd'), 'fill');
  if (shd && shd !== 'auto') css.push(`background-color: #${shd}`);
  const va = wAttr(child(rPr, 'vertAlign'), 'val');

  const parts: string[] = [];
  for (const c of only ?? Array.from(r.children)) {
    switch (c.localName) {
      case 't':
        parts.push(escape(c.textContent ?? ''));
        break;
      case 'tab':
        parts.push('\u2003\u2003');
        break;
      case 'br':
        parts.push('<br>');
        break;
      case 'drawing':
        parts.push(await drawingToHtml(c, ctx));
        break;
    }
  }
  let html = parts.join('');
  if (!html) return '';
  if (va === 'superscript') html = `<sup>${html}</sup>`;
  if (va === 'subscript') html = `<sub>${html}</sub>`;
  return css.length ? `<span style="${css.join('; ')}">${html}</span>` : html;
}

async function drawingToHtml(drawing: Element, ctx: PartCtx): Promise<string> {
  const blip = drawing.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0];
  const id = blip?.getAttributeNS(R, 'embed') ?? blip?.getAttribute('r:embed');
  const path = id ? ctx.rels.get(id) : null;
  if (!path) return '';
  const data = await ctx.zip.file(path)?.async('base64');
  if (!data) return '';
  const ext = path.split('.').pop()!.toLowerCase();
  const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : `image/${ext}`;
  const extent = drawing.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
  const w = Math.round(Number(extent?.getAttribute('cx') ?? 0) / 9525);
  const h = Math.round(Number(extent?.getAttribute('cy') ?? 0) / 9525);
  const size = w && h ? ` style="width: ${w}px; height: ${h}px"` : '';
  return `<img src="data:${mime};base64,${data}"${size}>`;
}

const escape = (s: string) => s.replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!);
