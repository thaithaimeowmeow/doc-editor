import { DOCX_MIME } from './base64';
import { imageSize } from './dom';
import { buildGrid, cellPos } from './table';
import { PageSetup, cssColorToHex, cssToPx, pxToEmu, pxToTwip } from './units';

export interface ExportInput {
  html: string;
  header: string;
  footer: string;
  firstHeader: string;
  firstFooter: string;
  differentFirstPage: boolean;
  page: PageSetup;
  title: string;
}

/**
 * HTML → .docx, tự sinh WordprocessingML (không dùng thư viện chuyển đổi bên ngoài, chỉ jszip để nén).
 * Hỗ trợ: đoạn, heading/title/quote, định dạng ký tự, danh sách nhiều cấp, bảng (gộp ô, viền, nền),
 * ảnh, liên kết, ngắt trang, header/footer với trường số trang, khổ giấy & lề.
 */
export async function exportDocx(input: ExportInput): Promise<Blob> {
  const { default: JSZip } = await import('jszip');
  const pkg = new Package(input.page);

  const body = await parse(input.html);
  const documentPart = pkg.part('word/document.xml');
  const bodyXml = pkg.blocks(body, documentPart, rootCtx(input.page));

  const refs: string[] = [];
  const hf = async (kind: 'header' | 'footer', type: 'default' | 'first', html: string) => {
    const n = pkg.parts.filter(p => p.name.startsWith(`word/${kind}`)).length + 1;
    const part = pkg.part(`word/${kind}${n}.xml`);
    const xml = pkg.blocks(await parse(html), part, rootCtx(input.page)) || '<w:p/>';
    part.xml = `<w:${kind === 'header' ? 'hdr' : 'ftr'} ${NS}>${xml}</w:${kind === 'header' ? 'hdr' : 'ftr'}>`;
    const id = documentPart.rel(`${RT}/${kind}`, `${kind}${n}.xml`);
    refs.push(`<w:${kind}Reference w:type="${type}" r:id="${id}"/>`);
  };
  if (input.header.trim()) await hf('header', 'default', input.header);
  if (input.footer.trim()) await hf('footer', 'default', input.footer);
  if (input.differentFirstPage) {
    await hf('header', 'first', input.firstHeader);
    await hf('footer', 'first', input.firstFooter);
  }

  const p = input.page;
  const landscape = p.width > p.height;
  const sectPr =
    `<w:sectPr>${refs.join('')}` +
    `<w:pgSz w:w="${pxToTwip(p.width)}" w:h="${pxToTwip(p.height)}"${landscape ? ' w:orient="landscape"' : ''}/>` +
    `<w:pgMar w:top="${pxToTwip(p.marginTop)}" w:right="${pxToTwip(p.marginRight)}" w:bottom="${pxToTwip(p.marginBottom)}" ` +
    `w:left="${pxToTwip(p.marginLeft)}" w:header="${pxToTwip(p.header)}" w:footer="${pxToTwip(p.footer)}" w:gutter="0"/>` +
    `${input.differentFirstPage ? '<w:titlePg/>' : ''}</w:sectPr>`;
  documentPart.xml = `<w:document ${NS}><w:body>${bodyXml || '<w:p/>'}${sectPr}</w:body></w:document>`;

  documentPart.rel(`${RT}/styles`, 'styles.xml');
  documentPart.rel(`${RT}/settings`, 'settings.xml');
  if (pkg.abstractNums.length) documentPart.rel(`${RT}/numbering`, 'numbering.xml');

  const zip = new JSZip();
  const xml = (s: string) => XML_HEAD + s;
  zip.file('[Content_Types].xml', xml(pkg.contentTypes()));
  zip.file(
    '_rels/.rels',
    xml(
      `<Relationships xmlns="${PR}">` +
        `<Relationship Id="rId1" Type="${RT}/officeDocument" Target="word/document.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
        `<Relationship Id="rId3" Type="${RT}/extended-properties" Target="docProps/app.xml"/>` +
        `</Relationships>`,
    ),
  );
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  zip.file(
    'docProps/core.xml',
    xml(
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
        `<dc:title>${esc(input.title)}</dc:title><dc:creator>Word Editor</dc:creator>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
        `</cp:coreProperties>`,
    ),
  );
  zip.file('docProps/app.xml', xml(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Word Editor</Application></Properties>`));
  zip.file('word/styles.xml', xml(STYLES));
  zip.file('word/settings.xml', xml(SETTINGS));
  if (pkg.abstractNums.length) zip.file('word/numbering.xml', xml(`<w:numbering ${NS}>${pkg.abstractNums.join('')}${pkg.nums.join('')}</w:numbering>`));
  for (const part of pkg.parts) {
    zip.file(part.name, xml(part.xml));
    if (part.rels.length) {
      const dir = part.name.replace(/[^/]+$/, '');
      zip.file(
        `${dir}_rels/${part.name.slice(dir.length)}.rels`,
        xml(
          `<Relationships xmlns="${PR}">` +
            part.rels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') +
            `</Relationships>`,
        ),
      );
    }
  }
  // Ảnh PNG/JPEG/GIF đã nén sẵn: nén lại chỉ tốn CPU (≈1,6 s cho 14 MB ảnh, kích thước đổi ~3%), nên lưu nguyên.
  // Giải mã base64 bằng trình duyệt (native) thay vì bộ giải mã JS của jszip.
  const bytes = await Promise.all(pkg.media.map(m => base64ToBytes(m.data)));
  pkg.media.forEach((m, i) => zip.file(m.path, bytes[i], { compression: 'STORE' }));

  return zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------------------------------

/** Giải mã base64 → byte bằng bộ giải mã gốc của trình duyệt (nhanh hơn nhiều so với JS thuần). */
async function base64ToBytes(b64: string): Promise<Uint8Array> {
  const native = (Uint8Array as unknown as { fromBase64?: (s: string) => Uint8Array }).fromBase64;
  if (native) return native.call(Uint8Array, b64);
  return new Uint8Array(await (await fetch(`data:application/octet-stream;base64,${b64}`)).arrayBuffer());
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const esc = (s: string) =>
  s.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

/** Đọc HTML, chuẩn bị ảnh (tải ảnh ngoài, đổi định dạng Word không đọc được sang PNG, đo kích thước). */
async function parse(html: string): Promise<HTMLElement> {
  const body = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html').body;
  await Promise.all(Array.from(body.querySelectorAll('img')).map(prepareImage));
  return body;
}

async function prepareImage(img: HTMLImageElement) {
  let src = img.getAttribute('src') ?? '';
  try {
    if (!src.startsWith('data:')) {
      const blob = await (await fetch(src)).blob();
      src = await new Promise<string>(res => {
        const r = new FileReader();
        r.onload = () => res(r.result as string);
        r.readAsDataURL(blob);
      });
    }
    if (!/^data:image\/(png|jpe?g|gif|bmp)/i.test(src)) src = await toPng(src);
  } catch {
    img.remove(); // ảnh không tải được (CORS…): bỏ qua thay vì làm hỏng cả file
    return;
  }
  const natural = await imageSize(src);
  const w = cssToPx(img.style.width) ?? (Number(img.getAttribute('width')) || null);
  const h = cssToPx(img.style.height) ?? (Number(img.getAttribute('height')) || null);
  const width = w ?? (h ? (h * natural.w) / natural.h : natural.w);
  const height = h ?? (w ? (w * natural.h) / natural.w : natural.h);
  img.setAttribute('src', src);
  img.dataset['w'] = String(width);
  img.dataset['h'] = String(height);
}

async function toPng(src: string): Promise<string> {
  const img = new Image();
  img.src = src;
  await img.decode();
  const canvas = document.createElement('canvas');
  canvas.width = img.naturalWidth || 300;
  canvas.height = img.naturalHeight || 200;
  canvas.getContext('2d')!.drawImage(img, 0, 0);
  return canvas.toDataURL('image/png');
}

// ---------- Package ----------

interface Rel {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

class Part {
  xml = '';
  rels: Rel[] = [];
  constructor(readonly name: string) {}
  rel(type: string, target: string, external = false) {
    const found = this.rels.find(r => r.type === type && r.target === target);
    if (found) return found.id;
    const id = `rId${this.rels.length + 1}`;
    this.rels.push({ id, type, target, external });
    return id;
  }
}

/** Ngữ cảnh khi duyệt khối: thuộc tính ký tự/đoạn kế thừa từ phần tử cha. */
interface Ctx {
  run: RunProps;
  align?: string;
  lineHeight?: string;
  indentLeft: number;
  indentRight: number;
  contentWidth: number;
  pStyle?: string;
  list?: { numId: number; level: number };
  pre?: boolean;
}

interface RunProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  caps?: boolean;
  vert?: 'superscript' | 'subscript';
  color?: string;
  shd?: string;
  highlight?: string;
  font?: string;
  /** px */
  size?: number;
  rStyle?: string;
}

const rootCtx = (p: PageSetup): Ctx => ({
  run: {},
  indentLeft: 0,
  indentRight: 0,
  contentWidth: p.width - p.marginLeft - p.marginRight,
});

const HEADINGS: Record<string, string> = { H1: 'Heading1', H2: 'Heading2', H3: 'Heading3', H4: 'Heading4', H5: 'Heading5', H6: 'Heading6' };
const INLINE_BLOCKS = /^(P|H[1-6]|DIV|UL|OL|LI|TABLE|BLOCKQUOTE|PRE|HR|SECTION|ARTICLE|HEADER|FOOTER|FIGURE|FIGCAPTION|ADDRESS|DL|DT|DD)$/;
const NAMED_SIZES: Record<string, number> = { 'xx-small': 9, 'x-small': 10, small: 13, medium: 16, large: 18, 'x-large': 24, 'xx-large': 32, 'xxx-large': 48 };

class Package {
  parts: Part[] = [];
  media: { path: string; data: string; ext: string }[] = [];
  private mediaBySrc = new Map<string, string>();
  abstractNums: string[] = [];
  nums: string[] = [];
  private docPrId = 0;

  constructor(private readonly page: PageSetup) {}

  part(name: string) {
    const p = new Part(name);
    this.parts.push(p);
    return p;
  }

  contentTypes() {
    const exts = new Set(this.media.map(m => m.ext));
    const mime: Record<string, string> = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
    const WP = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
    return (
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      Array.from(exts).map(e => `<Default Extension="${e}" ContentType="${mime[e]}"/>`).join('') +
      `<Override PartName="/word/document.xml" ContentType="${WP}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WP}.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${WP}.settings+xml"/>` +
      (this.abstractNums.length ? `<Override PartName="/word/numbering.xml" ContentType="${WP}.numbering+xml"/>` : '') +
      this.parts
        .filter(p => p.name !== 'word/document.xml')
        .map(p => `<Override PartName="/${p.name}" ContentType="${WP}.${p.name.includes('header') ? 'header' : 'footer'}+xml"/>`)
        .join('') +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
      `</Types>`
    );
  }

  // ---------- Khối ----------

  blocks(container: Element, part: Part, ctx: Ctx): string {
    const out: string[] = [];
    let inline: Node[] = [];
    const flush = () => {
      if (inline.some(n => n.nodeType !== 3 || n.textContent!.trim())) out.push(this.paragraph(null, inline, part, ctx));
      inline = [];
    };
    // Ngắt trang: gắn pageBreakBefore vào đoạn kế tiếp. Một đoạn chỉ chứa <w:br type="page"/> sẽ để lại
    // một dòng trống ở đầu trang sau (cả trong Word).
    let breakNext = false;
    const push = (xml: string) => {
      if (breakNext && xml) {
        out.push(withPageBreakBefore(xml));
        breakNext = false;
      } else if (xml) out.push(xml);
    };
    for (const node of Array.from(container.childNodes)) {
      if (node.nodeType === 1 && (node as Element).classList.contains('we-page-break')) {
        if (inline.length) flush();
        if (breakNext) out.push(PAGE_BREAK_P);
        breakNext = true;
      } else if (node.nodeType === 1 && INLINE_BLOCKS.test((node as Element).tagName)) {
        const before = out.length;
        flush();
        if (out.length > before && breakNext) {
          out[before] = withPageBreakBefore(out[before]);
          breakNext = false;
        }
        push(this.block(node as HTMLElement, part, ctx));
      } else if (node.nodeType === 1 || node.nodeType === 3) {
        inline.push(node);
      }
    }
    const before = out.length;
    flush();
    if (breakNext) {
      if (out.length > before) out[before] = withPageBreakBefore(out[before]);
      else out.push(PAGE_BREAK_P);
    }
    return out.join('');
  }

  private block(el: HTMLElement, part: Part, ctx: Ctx): string {
    const tag = el.tagName;
    if (el.classList.contains('we-page-break')) return PAGE_BREAK_P;
    if (tag === 'HR') return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
    if (tag === 'TABLE') return this.table(el as HTMLTableElement, part, ctx);
    if (tag === 'UL' || tag === 'OL') return this.list(el, part, ctx);

    const inner = this.childCtx(el, ctx);
    if (tag === 'BLOCKQUOTE') return this.blocks(el, part, { ...inner, pStyle: inner.pStyle ?? 'Quote' });
    if (tag === 'PRE') return this.paragraph(el, Array.from(el.childNodes), part, { ...inner, pre: true, run: { ...inner.run, font: inner.run.font ?? 'Courier New', size: inner.run.size ?? (10 * 4) / 3 } });
    if (tag === 'LI') return this.listItem(el, part, ctx);

    const hasBlocks = Array.from(el.children).some(c => INLINE_BLOCKS.test(c.tagName));
    if (hasBlocks) return this.blocks(el, part, inner); // DIV/SECTION… chứa khối con
    return this.paragraph(el, Array.from(el.childNodes), part, ctx);
  }

  /** Ngữ cảnh con: gộp style của phần tử chứa (div, td, blockquote…) vào phần kế thừa. */
  private childCtx(el: HTMLElement, ctx: Ctx): Ctx {
    const s = el.style;
    const fontPx = ctx.run.size ?? 16;
    return {
      ...ctx,
      run: this.runProps(el, ctx.run),
      align: s.textAlign || ctx.align,
      lineHeight: s.lineHeight || ctx.lineHeight,
      indentLeft: ctx.indentLeft + (el.tagName === 'TD' || el.tagName === 'TH' ? 0 : cssToPx(s.marginLeft, fontPx) ?? 0),
      indentRight: ctx.indentRight + (el.tagName === 'TD' || el.tagName === 'TH' ? 0 : cssToPx(s.marginRight, fontPx) ?? 0),
    };
  }

  private paragraph(el: HTMLElement | null, nodes: Node[], part: Part, ctx: Ctx, numPr = ''): string {
    const s = el?.style;
    const fontPx = ctx.run.size ?? 16;
    const run = el ? this.runProps(el, ctx.run) : ctx.run;
    const ppr: string[] = [];

    let style = ctx.pStyle;
    if (el && HEADINGS[el.tagName]) style = HEADINGS[el.tagName];
    if (el?.classList.contains('we-title')) style = 'Title';
    if (el?.classList.contains('we-subtitle')) style = 'Subtitle';
    if (style) ppr.push(`<w:pStyle w:val="${style}"/>`);
    if (s && (s.pageBreakBefore === 'always' || s.breakBefore === 'page')) ppr.push('<w:pageBreakBefore/>');
    if (numPr) ppr.push(numPr);
    const bg = cssColorToHex(s?.backgroundColor);
    if (bg) ppr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`);

    const before = cssToPx(s?.marginTop, fontPx);
    const after = cssToPx(s?.marginBottom, fontPx);
    const line = this.lineSpacing(s?.lineHeight || ctx.lineHeight, fontPx);
    const spacing = [
      before !== null ? `w:before="${pxToTwip(Math.max(0, before))}"` : '',
      after !== null ? `w:after="${pxToTwip(Math.max(0, after))}"` : '',
      line,
    ].filter(Boolean);
    if (spacing.length) ppr.push(`<w:spacing ${spacing.join(' ')}/>`);

    const left = ctx.indentLeft + (cssToPx(s?.marginLeft, fontPx) ?? 0) + (cssToPx(s?.paddingLeft, fontPx) ?? 0);
    const right = ctx.indentRight + (cssToPx(s?.marginRight, fontPx) ?? 0);
    const indent = cssToPx(s?.textIndent, fontPx) ?? 0;
    if (!numPr && (left || right || indent)) {
      ppr.push(
        `<w:ind w:left="${pxToTwip(left)}" w:right="${pxToTwip(right)}"` +
          (indent > 0 ? ` w:firstLine="${pxToTwip(indent)}"` : indent < 0 ? ` w:hanging="${pxToTwip(-indent)}"` : '') +
          '/>',
      );
    }

    const align = s?.textAlign || ctx.align;
    const jc = align === 'center' ? 'center' : align === 'right' || align === 'end' ? 'right' : align === 'justify' ? 'both' : align === 'left' ? 'left' : '';
    if (jc) ppr.push(`<w:jc w:val="${jc}"/>`);
    // Định dạng của dấu đoạn văn: quyết định chiều cao dòng cuối/đoạn trống trong Word.
    const { rStyle: _ignored, ...markProps } = run;
    const markRpr = this.rPr(markProps);
    if (markRpr) ppr.push(markRpr);

    // <br> cuối khối chỉ giữ chỗ cho dòng trống trong trình duyệt, không phải một dòng mới.
    const content = [...nodes];
    while (content.length && content.at(-1)!.nodeType === 1 && (content.at(-1) as Element).tagName === 'BR') content.pop();

    const runs: string[] = [];
    this.inline(content, part, { ...ctx, run }, runs, { atStart: true, lastSpace: false });
    // Đoạn trống: thêm một run rỗng mang định dạng, để trình đọc không hiểu pPr/rPr vẫn giữ đúng cỡ dòng.
    if (!runs.length && markRpr) runs.push(`<w:r>${markRpr}<w:t></w:t></w:r>`);
    return `<w:p>${ppr.length ? `<w:pPr>${ppr.join('')}</w:pPr>` : ''}${runs.join('')}</w:p>`;
  }

  private lineSpacing(v: string | undefined, fontPx: number): string {
    if (!v || v === 'normal') return '';
    if (/^[\d.]+$/.test(v)) return `w:line="${Math.round(240 * parseFloat(v))}" w:lineRule="auto"`;
    if (v.endsWith('%')) return `w:line="${Math.round(2.4 * parseFloat(v))}" w:lineRule="auto"`;
    const px = cssToPx(v, fontPx);
    return px ? `w:line="${pxToTwip(px)}" w:lineRule="atLeast"` : '';
  }

  // ---------- Danh sách ----------

  private list(el: HTMLElement, part: Part, ctx: Ctx): string {
    const ordered = el.tagName === 'OL';
    const level = ctx.list ? ctx.list.level + 1 : 0;
    const start = Number(el.getAttribute('start')) || 1;
    const numId = this.numbering(ordered, el.style.listStyleType, start, level);
    const inner: Ctx = { ...this.childCtx(el, ctx), indentLeft: ctx.indentLeft, list: { numId, level } };
    const out: string[] = [];
    for (const c of Array.from(el.children) as HTMLElement[]) {
      if (c.tagName === 'LI') out.push(this.listItem(c, part, inner));
      else if (c.tagName === 'UL' || c.tagName === 'OL') out.push(this.list(c, part, inner)); // Chrome lồng <ul> trực tiếp trong <ul>
      else out.push(this.block(c, part, inner));
    }
    return out.join('');
  }

  private listItem(li: HTMLElement, part: Part, ctx: Ctx): string {
    const list = ctx.list ?? { numId: this.numbering(false, '', 1, 0), level: 0 };
    const numPr = `<w:numPr><w:ilvl w:val="${list.level}"/><w:numId w:val="${list.numId}"/></w:numPr>`;
    const out: string[] = [];
    let inline: Node[] = [];
    let first = true;
    const flush = () => {
      if (!inline.length && !first) return;
      out.push(this.paragraph(li, inline, part, { ...ctx, indentLeft: 0 }, first ? numPr : ''));
      first = false;
      inline = [];
    };
    for (const node of Array.from(li.childNodes)) {
      if (node.nodeType === 1 && INLINE_BLOCKS.test((node as Element).tagName)) {
        const tag = (node as Element).tagName;
        if (tag === 'P' && first) {
          inline.push(...Array.from(node.childNodes));
          flush();
          continue;
        }
        if (inline.length || first) flush();
        out.push(this.block(node as HTMLElement, part, { ...ctx, run: this.runProps(li, ctx.run) }));
      } else {
        inline.push(node);
      }
    }
    if (inline.length || first) flush();
    return out.join('');
  }

  private numbering(ordered: boolean, styleType: string, start: number, level: number): number {
    const id = this.abstractNums.length;
    const FMT: Record<string, string> = {
      decimal: 'decimal',
      'decimal-leading-zero': 'decimalZero',
      'lower-alpha': 'lowerLetter',
      'lower-latin': 'lowerLetter',
      'upper-alpha': 'upperLetter',
      'upper-latin': 'upperLetter',
      'lower-roman': 'lowerRoman',
      'upper-roman': 'upperRoman',
    };
    const BUL: Record<string, string> = { disc: '•', circle: '◦', square: '▪' };
    const orderedCycle = ['decimal', 'lowerLetter', 'lowerRoman'];
    const bulletCycle = ['•', '◦', '▪'];

    const lvls: string[] = [];
    for (let l = 0; l < 9; l++) {
      const ind = `<w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr>`;
      if (ordered) {
        const fmt = l === level && FMT[styleType] ? FMT[styleType] : orderedCycle[l % 3];
        lvls.push(
          `<w:lvl w:ilvl="${l}"><w:start w:val="${l === level ? start : 1}"/><w:numFmt w:val="${fmt}"/>` +
            `<w:lvlText w:val="%${l + 1}."/><w:lvlJc w:val="left"/>${ind}</w:lvl>`,
        );
      } else {
        const ch = l === level && BUL[styleType] ? BUL[styleType] : bulletCycle[l % 3];
        lvls.push(
          `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${ch}"/>` +
            `<w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:lvl>`,
        );
      }
    }
    this.abstractNums.push(`<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls.join('')}</w:abstractNum>`);
    const numId = this.nums.length + 1;
    this.nums.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${id}"/></w:num>`);
    return numId;
  }

  // ---------- Bảng ----------

  private table(table: HTMLTableElement, part: Part, ctx: Ctx): string {
    const g = buildGrid(table);
    if (!g.rows.length || !g.cols) return '';
    const avail = ctx.contentWidth - ctx.indentLeft - ctx.indentRight;
    const widths = this.columnWidths(table, g, avail);
    const total = widths.reduce((a, b) => a + b, 0);

    const tw = table.style.width;
    const tblW = tw.endsWith('%')
      ? `<w:tblW w:w="${Math.round(parseFloat(tw) * 50)}" w:type="pct"/>`
      : `<w:tblW w:w="${pxToTwip(cssToPx(tw) ?? total)}" w:type="dxa"/>`;
    const centered = table.style.marginLeft === 'auto' && table.style.marginRight === 'auto';
    const tblBorders = this.borders(table, 'tblBorders', true);
    // Lề ô: lấy từ padding của ô đầu tiên (Word dùng một giá trị cho cả bảng).
    const first = table.rows[0]?.cells[0];
    const pad = (side: string, fallback: number | null) => {
      const v = cssToPx(first?.style.getPropertyValue(`padding-${side}`));
      return v === null ? fallback : pxToTwip(v);
    };
    const mar = { top: pad('top', null), left: pad('left', 108), bottom: pad('bottom', null), right: pad('right', 108) };
    const cellMar = Object.entries(mar)
      .filter(([, v]) => v !== null)
      .map(([k, v]) => `<w:${k} w:w="${v}" w:type="dxa"/>`)
      .join('');
    const tblPr =
      `<w:tblPr>${tblW}${centered ? '<w:jc w:val="center"/>' : ''}` +
      (ctx.indentLeft ? `<w:tblInd w:w="${pxToTwip(ctx.indentLeft)}" w:type="dxa"/>` : '') +
      tblBorders +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar>${cellMar}</w:tblCellMar>` +
      `<w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>`;
    const grid = `<w:tblGrid>${widths.map(w => `<w:gridCol w:w="${pxToTwip(w)}"/>`).join('')}</w:tblGrid>`;

    const cellCtx = (cell: HTMLElement): Ctx => ({
      ...this.childCtx(cell, ctx),
      indentLeft: 0,
      indentRight: 0,
      list: undefined,
      pStyle: undefined,
      contentWidth: widths.slice(0, 1)[0],
    });

    const rows: string[] = [];
    for (let r = 0; r < g.rows.length; r++) {
      const tr = g.rows[r];
      const cells: string[] = [];
      for (let c = 0; c < g.cols; ) {
        const cell = g.cells[r][c];
        if (!cell) {
          cells.push(`<w:tc><w:tcPr><w:tcW w:w="${pxToTwip(widths[c])}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`);
          c++;
          continue;
        }
        const p = cellPos(g, cell);
        const span = p.lastCol - p.col + 1;
        const w = widths.slice(c, c + span).reduce((a, b) => a + b, 0);
        const tcPr: string[] = [`<w:tcW w:w="${pxToTwip(w)}" w:type="dxa"/>`];
        if (span > 1) tcPr.push(`<w:gridSpan w:val="${span}"/>`);
        if (p.lastRow > p.row) tcPr.push(p.row === r ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>');
        tcPr.push(this.borders(cell, 'tcBorders', false));
        const fill = cssColorToHex(cell.style.backgroundColor) ?? cssColorToHex(tr.style.backgroundColor);
        if (fill) tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`);
        const va = cell.style.verticalAlign || tr.style.verticalAlign;
        if (va === 'middle') tcPr.push('<w:vAlign w:val="center"/>');
        else if (va === 'bottom') tcPr.push('<w:vAlign w:val="bottom"/>');

        let content = p.row === r ? this.blocks(cell, part, { ...cellCtx(cell), contentWidth: w }) : '';
        if (!content.endsWith('</w:p>') && !content.endsWith('<w:p/>')) content += '<w:p/>';
        cells.push(`<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${content}</w:tc>`);
        c += span;
      }
      const header = tr.parentElement?.tagName === 'THEAD' ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
      rows.push(`<w:tr>${header}${cells.join('')}</w:tr>`);
    }
    return `<w:tbl>${tblPr}${grid}${rows.join('')}</w:tbl>`;
  }

  private columnWidths(table: HTMLTableElement, g: ReturnType<typeof buildGrid>, avail: number): number[] {
    const tablePx = table.style.width.endsWith('%') ? (avail * parseFloat(table.style.width)) / 100 : cssToPx(table.style.width) ?? avail;
    const widths = new Array<number>(g.cols).fill(0);
    const cols = Array.from(table.querySelectorAll(':scope > colgroup > col')) as HTMLElement[];
    cols.forEach((col, i) => {
      if (i < g.cols) widths[i] = cssToPx(col.style.width) ?? (Number(col.getAttribute('width')) || 0);
    });
    for (let c = 0; c < g.cols; c++) {
      if (widths[c]) continue;
      for (let r = 0; r < g.rows.length; r++) {
        const cell = g.cells[r][c];
        if (!cell || Math.max(1, cell.colSpan) !== 1) continue;
        const w = cell.style.width;
        const px = w.endsWith('%') ? (tablePx * parseFloat(w)) / 100 : cssToPx(w) ?? (Number(cell.getAttribute('width')) || 0);
        if (px) {
          widths[c] = px;
          break;
        }
      }
    }
    const known = widths.reduce((a, b) => a + b, 0);
    const missing = widths.filter(w => !w).length;
    const fill = missing ? Math.max(24, (tablePx - known) / missing) : 0;
    return widths.map(w => w || fill);
  }

  private borders(el: HTMLElement, tag: 'tblBorders' | 'tcBorders', table: boolean): string {
    const s = el.style;
    const side = (name: 'Top' | 'Right' | 'Bottom' | 'Left') => {
      const style = s.getPropertyValue(`border-${name.toLowerCase()}-style`);
      if (!style) return null;
      if (style === 'none' || style === 'hidden') return 'nil';
      const width = cssToPx(s.getPropertyValue(`border-${name.toLowerCase()}-width`)) ?? 1;
      const color = cssColorToHex(s.getPropertyValue(`border-${name.toLowerCase()}-color`)) ?? 'auto';
      const val = style === 'dashed' ? 'dashed' : style === 'dotted' ? 'dotted' : style === 'double' ? 'double' : 'single';
      return `w:val="${val}" w:sz="${Math.max(2, Math.min(96, Math.round(width * 6)))}" w:space="0" w:color="${color}"`;
    };
    const sides = { top: side('Top'), left: side('Left'), bottom: side('Bottom'), right: side('Right') };
    const parts: string[] = [];
    for (const [name, v] of Object.entries(sides)) {
      if (v === 'nil') parts.push(`<w:${name} w:val="nil"/>`);
      else if (v) parts.push(`<w:${name} ${v}/>`);
    }
    if (table && sides.top && sides.top !== 'nil') {
      parts.push(`<w:insideH ${sides.top}/>`, `<w:insideV ${sides.left && sides.left !== 'nil' ? sides.left : sides.top}/>`);
    }
    return parts.length ? `<w:${tag}>${parts.join('')}</w:${tag}>` : '';
  }

  // ---------- Inline / run ----------

  private runProps(el: HTMLElement, base: RunProps): RunProps {
    const rp: RunProps = { ...base };
    const s = el.style;
    switch (el.tagName) {
      case 'B':
      case 'STRONG':
      case 'TH':
        rp.b = true;
        break;
      case 'I':
      case 'EM':
      case 'CITE':
      case 'DFN':
        rp.i = true;
        break;
      case 'U':
      case 'INS':
        rp.u = true;
        break;
      case 'S':
      case 'STRIKE':
      case 'DEL':
        rp.strike = true;
        break;
      case 'SUP':
        rp.vert = 'superscript';
        break;
      case 'SUB':
        rp.vert = 'subscript';
        break;
      case 'MARK':
        rp.highlight = 'yellow';
        break;
      case 'CODE':
      case 'KBD':
      case 'SAMP':
        rp.font = 'Courier New';
        break;
      case 'FONT': {
        const face = el.getAttribute('face');
        const color = el.getAttribute('color');
        if (face) rp.font = face.split(',')[0].replace(/['"]/g, '').trim();
        if (color) rp.color = cssColorToHex(color) ?? rp.color;
        break;
      }
    }
    const fw = s.fontWeight;
    if (fw) rp.b = fw === 'bold' || fw === 'bolder' || Number(fw) >= 600;
    if (s.fontStyle) rp.i = s.fontStyle === 'italic' || s.fontStyle === 'oblique';
    const deco = s.textDecorationLine || s.textDecoration;
    if (deco) {
      if (deco.includes('none')) {
        rp.u = false;
        rp.strike = false;
      } else {
        if (deco.includes('underline')) rp.u = true;
        if (deco.includes('line-through')) rp.strike = true;
      }
    }
    if (s.textTransform) rp.caps = s.textTransform === 'uppercase';
    if (s.verticalAlign === 'super') rp.vert = 'superscript';
    else if (s.verticalAlign === 'sub') rp.vert = 'subscript';
    const color = cssColorToHex(s.color);
    if (color) rp.color = color;
    const bg = cssColorToHex(s.backgroundColor);
    if (bg && !TEXT_BLOCK_TAGS.has(el.tagName)) rp.shd = bg;
    if (s.fontFamily) rp.font = s.fontFamily.split(',')[0].replace(/['"]/g, '').trim();
    if (s.fontSize) {
      const parent = base.size ?? 16;
      const fs = s.fontSize;
      rp.size = NAMED_SIZES[fs] ?? (fs.endsWith('%') ? (parent * parseFloat(fs)) / 100 : (cssToPx(fs, parent) ?? rp.size));
    }
    return rp;
  }

  private inline(nodes: Node[], part: Part, ctx: Ctx, out: string[], ws: { atStart: boolean; lastSpace: boolean }) {
    for (const node of nodes) {
      if (node.nodeType === 3) {
        let text = node.textContent ?? '';
        if (!ctx.pre && ctx.run.font !== 'Courier New' && !this.preserve(node)) {
          text = text.replace(/[ \t\n\r\f]+/g, ' ');
          if (ws.atStart || ws.lastSpace) text = text.replace(/^ /, '');
        }
        if (!text) continue;
        out.push(this.textRun(text, ctx.run));
        ws.atStart = false;
        ws.lastSpace = text.endsWith(' ');
        continue;
      }
      if (node.nodeType !== 1) continue;
      const el = node as HTMLElement;
      switch (el.tagName) {
        case 'BR':
          out.push('<w:r><w:br/></w:r>');
          ws.lastSpace = false;
          continue;
        case 'IMG':
          out.push(this.image(el as HTMLImageElement, part));
          ws.atStart = false;
          continue;
        case 'A': {
          const href = el.getAttribute('href') ?? '';
          const inner: string[] = [];
          this.inline(Array.from(el.childNodes), part, { ...ctx, run: { ...this.runProps(el, ctx.run), rStyle: 'Hyperlink' } }, inner, ws);
          if (!href) out.push(inner.join(''));
          else if (href.startsWith('#')) out.push(`<w:hyperlink w:anchor="${esc(href.slice(1))}">${inner.join('')}</w:hyperlink>`);
          else out.push(`<w:hyperlink r:id="${part.rel(`${RT}/hyperlink`, href, true)}" w:history="1">${inner.join('')}</w:hyperlink>`);
          continue;
        }
        case 'SPAN':
          if (el.dataset['weField']) {
            const instr = el.dataset['weField'] === 'pages' ? 'NUMPAGES' : 'PAGE';
            const rp = this.runProps(el, ctx.run);
            out.push(`<w:fldSimple w:instr=" ${instr} \\* MERGEFORMAT ">${this.textRun('1', rp)}</w:fldSimple>`);
            ws.atStart = false;
            ws.lastSpace = false;
            continue;
          }
          break;
      }
      // Phần tử khối lồng trong inline (hiếm): coi như inline.
      this.inline(Array.from(el.childNodes), part, { ...ctx, run: this.runProps(el, ctx.run) }, out, ws);
    }
  }

  private preserve(node: Node) {
    for (let el = node.parentElement; el; el = el.parentElement) {
      const ws = el.style?.whiteSpace;
      if (ws) return ws.startsWith('pre') || ws === 'break-spaces';
    }
    return false;
  }

  private textRun(text: string, rp: RunProps): string {
    const parts = text.split(/(\t|\n)/).map(t =>
      t === '\t' ? '<w:tab/>' : t === '\n' ? '<w:br/>' : t ? `<w:t xml:space="preserve">${esc(t)}</w:t>` : '',
    );
    return `<w:r>${this.rPr(rp)}${parts.join('')}</w:r>`;
  }

  private rPr(rp: RunProps): string {
    const x: string[] = [];
    if (rp.rStyle) x.push(`<w:rStyle w:val="${rp.rStyle}"/>`);
    if (rp.font) {
      const f = esc(rp.font);
      x.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`);
    }
    if (rp.b !== undefined) x.push(rp.b ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>');
    if (rp.i !== undefined) x.push(rp.i ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>');
    if (rp.caps) x.push('<w:caps/>');
    if (rp.strike) x.push('<w:strike/>');
    if (rp.color) x.push(`<w:color w:val="${rp.color}"/>`);
    if (rp.size) {
      const half = Math.max(2, Math.round(rp.size * 0.75 * 2));
      x.push(`<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`);
    }
    if (rp.highlight) x.push(`<w:highlight w:val="${rp.highlight}"/>`);
    if (rp.u !== undefined) x.push(`<w:u w:val="${rp.u ? 'single' : 'none'}"/>`);
    if (rp.shd) x.push(`<w:shd w:val="clear" w:color="auto" w:fill="${rp.shd}"/>`);
    if (rp.vert) x.push(`<w:vertAlign w:val="${rp.vert}"/>`);
    return x.length ? `<w:rPr>${x.join('')}</w:rPr>` : '';
  }

  private image(img: HTMLImageElement, part: Part): string {
    const src = img.getAttribute('src') ?? '';
    const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(src);
    if (!m) return '';
    let path = this.mediaBySrc.get(src);
    if (!path) {
      const ext = m[1].toLowerCase() === 'jpg' ? 'jpeg' : m[1].toLowerCase();
      path = `media/image${this.media.length + 1}.${ext}`;
      this.media.push({ path: `word/${path}`, data: m[2], ext });
      this.mediaBySrc.set(src, path);
    }
    const rid = part.rel(`${RT}/image`, path);
    const maxW = this.page.width - this.page.marginLeft - this.page.marginRight;
    let w = Number(img.dataset['w']) || 200;
    let h = Number(img.dataset['h']) || 150;
    if (w > maxW) {
      h = (h * maxW) / w;
      w = maxW;
    }
    const cx = pxToEmu(w), cy = pxToEmu(h);
    const id = ++this.docPrId;
    const name = `Picture ${id}`;
    return (
      `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${id}" name="${name}" descr="${esc(img.getAttribute('alt') ?? '')}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    );
  }
}

const TEXT_BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'TD', 'TH', 'DIV', 'BLOCKQUOTE', 'PRE', 'TABLE', 'TR', 'UL', 'OL']);

const heading = (id: string, name: string, size: number, before: number, after: number, color: string, extra = '', level = 0) =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${level}"/></w:pPr>` +
  `<w:rPr><w:b/><w:bCs/>${extra}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

/** Phải khớp với content-css.ts để mở lại file trong editor trông giống nhau. */
const STYLES =
  `<w:styles ${NS}>` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>` +
  `<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="vi-VN" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
  `<w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  heading('Heading1', 'heading 1', 32, 240, 60, '2F5496', '', 0) +
  heading('Heading2', 'heading 2', 26, 200, 40, '2F5496', '', 1) +
  heading('Heading3', 'heading 3', 24, 160, 40, '1F3763', '', 2) +
  heading('Heading4', 'heading 4', 24, 160, 40, '2F5496', '<w:i/><w:iCs/>', 3) +
  heading('Heading5', 'heading 5', 24, 120, 0, '2F5496', '', 4) +
  heading('Heading6', 'heading 6', 24, 120, 0, '2F5496', '', 5) +
  `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="595959"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:spacing w:before="200" w:after="200"/><w:ind w:left="860" w:right="860"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>` +
  `</w:styles>`;

const SETTINGS =
  `<w:settings ${NS}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const PAGE_BREAK_P = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

/** Thêm <w:pageBreakBefore/> vào đoạn đầu tiên của khối XML (đúng vị trí theo schema: sau pStyle). Bảng thì chèn đoạn ngắt trang trước. */
function withPageBreakBefore(xml: string): string {
  const pPrEnd = xml.indexOf('</w:pPr>');
  if (xml.startsWith('<w:p><w:pPr>') && xml.slice(0, pPrEnd).includes('<w:pageBreakBefore/>')) return xml;
  if (!xml.startsWith('<w:p>') && !xml.startsWith('<w:p ')) return PAGE_BREAK_P + xml;
  if (xml.startsWith('<w:p><w:pPr>')) {
    const style = /^<w:p><w:pPr>(<w:pStyle [^>]*\/>)?/.exec(xml)![0];
    return style + '<w:pageBreakBefore/>' + xml.slice(style.length);
  }
  return '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>' + xml.slice('<w:p>'.length);
}
