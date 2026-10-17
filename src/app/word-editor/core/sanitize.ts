import DOMPurify from 'dompurify';

import { ListItem, buildLists, markerToStyle } from './lists';

const purifier = DOMPurify();

purifier.addHook('uponSanitizeAttribute', (_node, data) => {
  if (data.attrName === 'class') {
    // Chỉ giữ class của editor (we-title, we-page-break, we-field…), bỏ class rác từ Word/web.
    data.attrValue = data.attrValue
      .split(/\s+/)
      .filter(c => c.startsWith('we-') && !['we-cell-sel', 'we-img-sel', 'we-pb-spacer', 'we-pb-head'].includes(c))
      .join(' ');
    if (!data.attrValue) data.keepAttr = false;
  }
  if (data.attrName === 'data-we-pb' || data.attrName === 'contenteditable') data.keepAttr = false;
});

purifier.addHook('afterSanitizeAttributes', node => {
  if (node.tagName === 'A' && node.getAttribute('target') === '_blank') node.setAttribute('rel', 'noopener noreferrer');
});

/** Làm sạch HTML trước khi đưa vào editor hoặc trả ra ngoài. */
export function sanitize(html: string): string {
  // Ảnh nhúng base64 chiếm gần hết dung lượng tài liệu (14 MB / 14,7 MB trong file 50 trang) và
  // DOMPurify phải duyệt từng ký tự của chúng (170 ms). Thay bằng chỗ giữ chỗ rồi trả lại: ~37 ms.
  // An toàn vì chỗ giữ chỗ vẫn nằm trong thuộc tính do DOMPurify kiểm tra; thuộc tính nào bị loại
  // (vd: href="data:...") thì mất luôn chỗ giữ chỗ và không có gì để trả lại.
  const images: string[] = [];
  const light = html.length < LIGHT_THRESHOLD ? html : html.replace(IMAGE_DATA_URL, m => `${PLACEHOLDER}${images.push(m) - 1}@@`);
  const clean = purifier.sanitize(light, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['style', 'script', 'iframe', 'form', 'input', 'button', 'textarea', 'select', 'object', 'embed', 'video', 'audio', 'meta', 'link', 'base'],
    ADD_ATTR: ['target', 'start', 'colspan', 'rowspan'],
  }) as string;
  return images.length ? clean.replace(PLACEHOLDER_RE, (m, i) => images[Number(i)] ?? m) : clean;
}

const LIGHT_THRESHOLD = 200_000;
// Chỉ ảnh bitmap (không có svg: svg có thể chứa script).
const IMAGE_DATA_URL = /data:image\/(?:png|jpe?g|gif|webp|bmp);base64,[A-Za-z0-9+/=]{200,}/g;
const PLACEHOLDER = 'data:image/png;base64,@@';
const PLACEHOLDER_RE = /data:image\/png;base64,@@(\d+)@@/g;

const JUNK_CSS = /^(mso-|tab-stops|layout-grid|text-autospace|text-justify|font-variant-ligatures|font-variant-caps|orphans|widows|-webkit-|caret-color|word-spacing|text-size-adjust)/;

/** Dọn HTML dán vào từ Word / Google Docs / trang web, giữ định dạng, đổi danh sách Word thành <ol>/<ul>. */
export function cleanPastedHtml(html: string): string {
  const fragment = html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i)?.[1] ?? html;
  const doc = new DOMParser().parseFromString(fragment, 'text/html');
  const body = doc.body;

  // Danh sách của Word: <p style="mso-list:l0 level1 lfo1"><!--[if !supportLists]-->·<!--[endif]-->Nội dung</p>
  const items = new Map<HTMLElement, ListItem>();
  for (const p of Array.from(body.querySelectorAll<HTMLElement>('p[style*="mso-list"]'))) {
    const style = p.getAttribute('style') ?? '';
    const m = /mso-list:\s*l(\d+)\s+level(\d+)/i.exec(style);
    if (!m) continue;
    let marker = '';
    const walker = doc.createTreeWalker(p, NodeFilter.SHOW_COMMENT);
    const comments: Comment[] = [];
    for (let c = walker.nextNode(); c; c = walker.nextNode()) comments.push(c as Comment);
    const open = comments.find(c => /\[if !supportLists\]/i.test(c.data));
    const close = open && comments.find(c => c !== open && /\[endif\]/i.test(c.data) && (open.compareDocumentPosition(c) & Node.DOCUMENT_POSITION_FOLLOWING));
    if (open && close) {
      const r = doc.createRange();
      r.setStartAfter(open);
      r.setEndBefore(close);
      marker = r.toString();
      r.deleteContents();
    }
    items.set(p, { key: m[1], level: Number(m[2]) - 1, ...markerToStyle(marker) });
  }
  buildLists(items);

  for (const el of Array.from(body.querySelectorAll<HTMLElement>('*'))) {
    // Google Docs bọc toàn bộ trong <b style="font-weight:normal">.
    if (el.tagName === 'B' && /font-weight:\s*normal/i.test(el.getAttribute('style') ?? '')) {
      el.replaceWith(...Array.from(el.childNodes));
      continue;
    }
    if (el.tagName.includes(':')) {
      // o:p, w:sdt, v:shape…
      el.replaceWith(...Array.from(el.childNodes));
      continue;
    }
    const style = el.getAttribute('style');
    if (style) {
      const kept = style
        .split(';')
        .map(s => s.trim())
        .filter(s => s && !JUNK_CSS.test(s.split(':')[0].trim().toLowerCase()));
      if (kept.length) el.setAttribute('style', kept.join('; '));
      else el.removeAttribute('style');
    }
    el.removeAttribute('lang');
    el.removeAttribute('id');
  }
  return sanitize(body.innerHTML);
}

/** Bỏ thuộc tính tạm của editor khỏi bản sao trước khi lưu/xuất. */
export function stripTransient(root: ParentNode) {
  for (const el of Array.from(root.querySelectorAll('.we-pb-spacer'))) el.remove();
  for (const el of Array.from(root.querySelectorAll('[data-we-pb]'))) el.removeAttribute('data-we-pb');
  for (const el of Array.from(root.querySelectorAll('.we-cell-sel'))) {
    el.classList.remove('we-cell-sel');
    if (!el.classList.length) el.removeAttribute('class');
  }
  for (const el of Array.from(root.querySelectorAll('[contenteditable]'))) el.removeAttribute('contenteditable');
  for (const el of Array.from(root.querySelectorAll('.we-img-sel'))) {
    el.classList.remove('we-img-sel');
    if (!el.classList.length) el.removeAttribute('class');
  }
}
