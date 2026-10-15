import { CONTENT_CSS } from './content-css';
import { escapeHtml } from './dom';
import { PageSetup } from './units';

/** Điền số trang vào các trường PAGE/NUMPAGES. */
export function fillFields(html: string, page: number, total: number): string {
  if (!html.includes('data-we-field')) return html;
  const box = document.createElement('div');
  box.innerHTML = html;
  for (const f of Array.from(box.querySelectorAll<HTMLElement>('[data-we-field]'))) {
    f.textContent = String(f.dataset['weField'] === 'pages' ? total : page);
  }
  return box.innerHTML;
}

export interface PrintJob {
  page: PageSetup;
  /**
   * Mỗi trang: HTML các khối nằm trên trang (còn giữ data-we-pb / hàng đệm của bộ chia trang) và
   * vị trí (px) của mép trên khối đầu tiên so với đỉnh vùng nội dung của trang.
   */
  slices: { html: string; offset: number }[];
  spacerCss: string;
  scope: string;
  headerFor(page: number): string;
  footerFor(page: number): string;
  title: string;
}

/**
 * Dựng tài liệu in: mỗi trang là một "cửa sổ" cắt đúng phần thân văn bản của trang đó, nên bản in
 * khớp với màn hình (cùng điểm ngắt trang, có header/footer và số trang).
 */
export function buildPrintHtml(job: PrintJob): string {
  const p = job.page;
  const cw = p.width - p.marginLeft - p.marginRight;
  const avail = p.height - p.marginTop - p.marginBottom;
  const pages = job.slices.map(
    (s, i) =>
      `<div class="pg">` +
      `<div class="we-doc hf" style="top:${p.header}px">${job.headerFor(i)}</div>` +
      `<div class="win" style="top:${p.marginTop}px;height:${avail + 1}px">` +
      `<div class="we-doc we-body-print" data-we-scope="${job.scope}" style="top:${s.offset}px">${s.html}</div></div>` +
      `<div class="we-doc hf ft" style="bottom:${p.footer}px">${job.footerFor(i)}</div>` +
      `</div>`,
  );
  return `<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><title>${escapeHtml(job.title)}</title><style>
${CONTENT_CSS}
${job.spacerCss}
@page { size: ${p.width}px ${p.height}px; margin: 0; }
html, body { margin: 0; padding: 0; background: #fff; }
.pg { position: relative; width: ${p.width}px; height: ${p.height}px; overflow: hidden; break-after: page; }
.pg:last-child { break-after: auto; }
.hf { position: absolute; left: ${p.marginLeft}px; width: ${cw}px; }
.win { position: absolute; left: ${p.marginLeft}px; width: ${cw}px; overflow: hidden; }
.we-body-print { position: absolute; left: 0; width: ${cw}px; display: flow-root; }
.we-body-print tr.we-pb-spacer:not(.we-pb-head) > td { border: none !important; padding: 0 !important; background: transparent !important; }
</style></head><body>${pages.join('')}</body></html>`;
}

export async function printHtml(html: string) {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.append(iframe);
  await new Promise<void>(resolve => {
    iframe.onload = () => resolve();
    iframe.srcdoc = html;
  });
  const doc = iframe.contentDocument!;
  await Promise.all(Array.from(doc.images).map(img => img.decode().catch(() => {})));
  await doc.fonts?.ready;
  iframe.contentWindow!.focus();
  iframe.contentWindow!.print();
  setTimeout(() => iframe.remove(), 1000);
}

/** File .html độc lập (mở bằng trình duyệt / Word), có CSS khổ giấy để in. */
export function standaloneHtml(bodyHtml: string, page: PageSetup, title: string): string {
  const cw = page.width - page.marginLeft - page.marginRight;
  return `<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
${CONTENT_CSS}
@page { size: ${page.width}px ${page.height}px; margin: ${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px; }
body { margin: 0; background: #fff; }
.we-doc { max-width: ${cw}px; margin: 40px auto; padding: 0 16px; }
.we-doc .we-page-break { break-after: page; }
@media print { .we-doc { margin: 0; padding: 0; max-width: none; } }
</style>
</head>
<body>
<div class="we-doc">
${bodyHtml}
</div>
</body>
</html>`;
}
