/**
 * Chia trang kiểu Word trên một vùng contenteditable liên tục.
 *
 * Nội dung là một luồng duy nhất (để con trỏ, chọn vùng, undo hoạt động tự nhiên). Các trang chỉ là
 * hình nền phía sau. Khối nào cắt qua đáy vùng nội dung của một trang thì được đẩy xuống đầu trang
 * sau bằng `margin-top` (khoảng đẩy nằm trong một `<style>` riêng, gắn qua thuộc tính `data-we-pb`).
 * Hàng bảng không có margin, nên chèn một hàng đệm tạm (`tr.we-pb-spacer`) phía trước, kèm bản sao
 * hàng tiêu đề (`<thead>`) để lặp lại ở đầu trang mới như Word. Mọi thứ tạm đều bị gỡ khi lưu/xuất.
 */

export interface PageGeometry {
  /** Khoảng cách giữa đỉnh vùng nội dung của hai trang liên tiếp = chiều cao trang + khe giữa trang. */
  stride: number;
  /** Chiều cao vùng nội dung của một trang = chiều cao trang − lề trên − lề dưới. */
  avail: number;
}

const ATTR = 'data-we-pb';
const EPS = 0.5;
const DESCEND = new Set(['TABLE', 'TBODY', 'THEAD', 'TFOOT', 'UL', 'OL', 'BLOCKQUOTE', 'DIV', 'LI']);

export class Paginator {
  private rules: string[] = [];
  private seq = 0;

  constructor(
    private readonly body: HTMLElement,
    private readonly sheet: HTMLStyleElement,
    /** Selector phạm vi để rule không ảnh hưởng editor khác trên cùng trang. */
    private readonly scope: string,
  ) {}

  clear() {
    for (const el of Array.from(this.body.querySelectorAll(`[${ATTR}]`))) el.removeAttribute(ATTR);
    for (const el of Array.from(this.body.querySelectorAll(`.${SPACER}`))) el.remove();
    this.rules = [];
    this.seq = 0;
    this.sheet.textContent = '';
  }

  /** Chạy lại toàn bộ; trả về số trang. `scale` là hệ số zoom (transform) đang áp dụng. */
  run(g: PageGeometry, scale: number): number {
    this.clear();
    const origin = () => this.body.getBoundingClientRect().top;
    const rect = (el: Element) => {
      const r = el.getBoundingClientRect();
      const o = origin();
      return { top: (r.top - o) / scale, bottom: (r.bottom - o) / scale, height: r.height / scale };
    };

    let forcedTop = -1;

    const visit = (container: Element) => {
      for (const el of Array.from(container.children) as HTMLElement[]) {
        if (el.offsetParent === null && el.tagName !== 'TR') continue; // display:none
        if (el.classList.contains(SPACER)) continue;
        if (el.classList.contains('we-page-break')) {
          const r = rect(el);
          forcedTop = (Math.floor((r.top - EPS) / g.stride) + 1) * g.stride;
          continue;
        }

        let r = rect(el);
        if (forcedTop >= 0) {
          if (r.top < forcedTop - EPS) {
            this.push(el, forcedTop - r.top, rect);
            r = rect(el);
          }
          forcedTop = -1;
        }

        const page = Math.floor((r.top + EPS) / g.stride);
        const local = r.top - page * g.stride;
        const pageBottom = page * g.stride + g.avail;

        // Bắt đầu trong vùng lề dưới / khe giữa trang → đẩy xuống trang sau.
        if (local > g.avail - EPS) {
          this.push(el, (page + 1) * g.stride - r.top, rect);
          continue;
        }
        if (r.bottom <= pageBottom + EPS) continue;

        // Cắt qua đáy trang: bảng/danh sách thì ngắt giữa các hàng/mục như Word.
        if (DESCEND.has(el.tagName) && hasBlockChildren(el)) {
          visit(el);
          continue;
        }
        if (r.height <= g.avail && local > EPS) {
          this.push(el, (page + 1) * g.stride - r.top, rect);
        }
        // Khối cao hơn cả một trang (ảnh lớn, hàng bảng rất cao): để tràn, giống Word khi không thể ngắt.
      }
    };

    visit(this.body);
    this.sheet.textContent = this.rules.join('\n');

    const height = this.body.scrollHeight;
    return Math.max(1, Math.floor((height - EPS) / g.stride) + 1);
  }

  /** Đẩy phần tử xuống `delta` px; đo lại để bù cho margin collapse. */
  private push(el: HTMLElement, delta: number, rect: (el: Element) => { top: number }) {
    if (delta <= EPS) return;
    if (el.tagName === 'TR') return this.pushRow(el as HTMLTableRowElement, delta);
    const id = String(++this.seq);
    el.setAttribute(ATTR, id);
    const target = rect(el).top + delta;
    let amount = (parseFloat(getComputedStyle(el).marginTop) || 0) + delta;

    const sel = `${this.scope} [${ATTR}="${id}"]`;
    const rule = (v: number) => `${sel} { margin-top: ${v}px !important; }`;

    const index = this.rules.length;
    this.rules.push(rule(amount));
    for (let i = 0; i < 3; i++) {
      this.sheet.textContent = this.rules.join('\n');
      const diff = target - rect(el).top;
      if (Math.abs(diff) <= EPS) break;
      amount += diff;
      this.rules[index] = rule(amount);
    }
  }

  /** Hàng bảng: chèn hàng đệm cao `delta` rồi lặp lại hàng tiêu đề (nếu có) ngay trước hàng bị đẩy. */
  private pushRow(row: HTMLTableRowElement, delta: number) {
    const table = row.closest('table');
    if (!table) return;
    const cols = Math.max(1, ...Array.from(table.rows).map(r => Array.from(r.cells).reduce((n, c) => n + Math.max(1, c.colSpan), 0)));
    const gap = document.createElement('tr');
    gap.className = SPACER;
    gap.contentEditable = 'false';
    gap.setAttribute('aria-hidden', 'true');
    const td = document.createElement('td');
    td.colSpan = cols;
    td.style.cssText = `height:${delta}px;padding:0;border:none;background:transparent`;
    gap.append(td);
    row.before(gap);

    const head = table.tHead;
    if (head && row.parentElement !== head) {
      for (const h of Array.from(head.rows)) {
        const copy = h.cloneNode(true) as HTMLTableRowElement;
        copy.classList.add(SPACER, 'we-pb-head');
        copy.contentEditable = 'false';
        copy.setAttribute('aria-hidden', 'true');
        copy.removeAttribute(ATTR);
        row.before(copy);
      }
    }
  }
}

const SPACER = 'we-pb-spacer';

const CHILD_BLOCKS = /^(P|H\d|UL|OL|LI|TABLE|TBODY|THEAD|TFOOT|TR|DIV|BLOCKQUOTE|PRE|HR)$/;

function hasBlockChildren(el: HTMLElement) {
  return Array.from(el.children).some(c => CHILD_BLOCKS.test(c.tagName));
}
