/**
 * Chia trang kiểu Word trên một vùng contenteditable liên tục.
 *
 * Nội dung là một luồng duy nhất (để con trỏ, chọn vùng, undo hoạt động tự nhiên). Các trang chỉ là
 * hình nền phía sau. Khối nào cắt qua đáy vùng nội dung của một trang thì được đẩy xuống đầu trang
 * sau bằng `margin-top` (khoảng đẩy nằm trong một `<style>` riêng, gắn qua thuộc tính `data-we-pb`).
 * Hàng bảng không có margin, nên chèn một hàng đệm tạm (`tr.we-pb-spacer`) phía trước, kèm bản sao
 * hàng tiêu đề (`<thead>`) để lặp lại ở đầu trang mới như Word. Mọi thứ tạm đều bị gỡ khi lưu/xuất.
 *
 * Hiệu năng: đo toàn bộ trong MỘT lượt khi DOM chưa bị sửa, theo dõi độ dịch cộng dồn bằng số học
 * ("dịch ảo"), rồi ghi tất cả khoảng đẩy một lần và kiểm tra lại bằng vài lượt đọc theo lô. Cách cũ
 * xen kẽ đọc/ghi nên mỗi lần đẩy buộc trình duyệt dàn trang lại cả tài liệu (50 trang ≈ 550 ms).
 *
 * Tăng dần: khi biết khối cấp cao nhất bị sửa (`from`), giữ nguyên mọi khoảng đẩy đứng trước nó và chỉ
 * chia lại từ khối đó trở xuống. Nếu chỉ sửa một khối mà tổng chiều cao tài liệu không đổi (gõ thêm
 * một ký tự không xuống dòng) thì bố cục sau đó chắc chắn không đổi, bỏ qua hẳn.
 */

export interface PageGeometry {
  /** Khoảng cách giữa đỉnh vùng nội dung của hai trang liên tiếp = chiều cao trang + khe giữa trang. */
  stride: number;
  /** Chiều cao vùng nội dung của một trang = chiều cao trang − lề trên − lề dưới. */
  avail: number;
}

const ATTR = 'data-we-pb';
const SPACER = 'we-pb-spacer';
const EPS = 0.5;
const DESCEND = new Set(['TABLE', 'TBODY', 'THEAD', 'TFOOT', 'UL', 'OL', 'BLOCKQUOTE', 'DIV', 'LI']);
/** Phần tử không phải khối dàn trang (hộp của <colgroup> trùng cả bảng, margin vô tác dụng). */
const SKIP = new Set(['COLGROUP', 'COL', 'CAPTION', 'SCRIPT', 'STYLE']);
const CHILD_BLOCKS = /^(P|H\d|UL|OL|LI|TABLE|TBODY|THEAD|TFOOT|TR|DIV|BLOCKQUOTE|PRE|HR)$/;

/** Một lần đẩy đã lên kế hoạch, chưa ghi vào DOM. */
interface Push {
  el: HTMLElement;
  /** Vị trí đỉnh phần tử mong đợi sau khi đẩy (hệ toạ độ tài liệu, đã tính các lần đẩy trước). */
  expected: number;
  /** Khoảng đẩy hiện tại (px): margin-top mới, hoặc chiều cao hàng đệm. */
  amount: number;
  row?: { table: HTMLTableElement; head: HTMLTableRowElement[] };
  /** Ô của hàng đệm đã chèn (chỉ với hàng bảng). */
  gap?: HTMLElement;
  /** Mã thuộc tính `data-we-pb` (chỉ với đẩy bằng margin). */
  id?: string;
  /** Hàng đệm + bản sao hàng tiêu đề đã chèn, để gỡ khi chia lại. */
  nodes?: HTMLElement[];
}

export class Paginator {
  private rules = new Map<string, string>();
  private seq = 0;
  /** Các lần đẩy đang có hiệu lực, theo thứ tự tài liệu. */
  private pushes: Push[] = [];
  private valid = false;
  private geo = '';
  private height = 0;
  private pages = 1;

  constructor(
    private readonly body: HTMLElement,
    private readonly sheet: HTMLStyleElement,
    /** Selector phạm vi để rule không ảnh hưởng editor khác trên cùng trang. */
    private readonly scope: string,
  ) {}

  clear() {
    for (const el of Array.from(this.body.querySelectorAll(`[${ATTR}]`))) el.removeAttribute(ATTR);
    for (const el of Array.from(this.body.querySelectorAll(`.${SPACER}`))) el.remove();
    this.rules.clear();
    this.pushes = [];
    this.valid = false;
    this.seq = 0;
    this.sheet.textContent = '';
  }

  /** Chiều cao tài liệu đã đổi so với lần chia trang gần nhất (vd: ảnh vừa tải xong). */
  isStale(scale: number): boolean {
    return !this.valid || Math.abs(this.body.getBoundingClientRect().height / scale - this.height) > 0.5;
  }

  /**
   * Chia trang; trả về số trang. `scale` là hệ số zoom (transform) đang áp dụng.
   * `from`: khối cấp cao nhất sớm nhất bị sửa (phía trước nó không đổi); bỏ trống = chia lại toàn bộ.
   * `single`: chỉ đúng một khối bị sửa và không thêm/bớt khối nào, cho phép bỏ qua khi chiều cao không đổi.
   */
  run(g: PageGeometry, scale: number, from?: Element | null, single = false): number {
    const geo = `${g.stride}|${g.avail}`;
    const inc = !!from && this.valid && geo === this.geo && from.parentElement === this.body;
    if (inc && single && Math.abs(this.body.getBoundingClientRect().height / scale - this.height) < 0.01) return this.pages;

    let start: Element | null = null;
    if (inc) {
      // Lùi một khối (margin gộp với khối trước; ngắt trang cưỡng bức đứng trước khối bị sửa).
      start = from!.previousElementSibling ?? from!;
      // Chuỗi keepNext đứng trước, và ngắt trang cưỡng bức ngay trước điểm bắt đầu (nó quyết định vị trí khối kế tiếp).
      for (let prev = start.previousElementSibling; prev; prev = start.previousElementSibling) {
        if (!prev.classList.contains('we-page-break') && !keepsNext(prev as HTMLElement)) break;
        start = prev;
      }
      while (start.classList.contains('we-page-break') && start.previousElementSibling) start = start.previousElementSibling;
      this.drop(start);
    } else this.clear();
    this.geo = geo;

    const origin = this.body.getBoundingClientRect().top; // lượt dàn trang đầu tiên
    const pushes: Push[] = [];
    const headHeight = new Map<HTMLTableElement, number>();
    let shift = 0; // tổng chiều cao đã chèn thêm phía trước vị trí đang xét
    let forcedTop = -1;

    const rect = (el: Element) => {
      const r = el.getBoundingClientRect();
      return { top: (r.top - origin) / scale + shift, bottom: (r.bottom - origin) / scale + shift, height: r.height / scale };
    };

    const push = (el: HTMLElement, delta: number, top: number) => {
      if (delta <= EPS) return;
      if (el.tagName === 'TR') {
        const table = el.closest('table') as HTMLTableElement | null;
        if (!table) return;
        const head = table.tHead && el.parentElement !== table.tHead ? Array.from(table.tHead.rows) : [];
        let hh = headHeight.get(table);
        if (hh === undefined) {
          hh = head.reduce((n, h) => n + h.getBoundingClientRect().height / scale, 0);
          headHeight.set(table, hh);
        }
        pushes.push({ el, expected: top + delta + (head.length ? hh : 0), amount: delta, row: { table, head } });
        shift += delta + (head.length ? hh : 0);
        return;
      }
      // Margin của phần tử gộp với margin-bottom của phần tử đứng trước (lấy giá trị lớn hơn), nên
      // khe hiện tại là max(trước, mình); khe mong muốn = khe hiện tại + delta.
      const mt = parseFloat(getComputedStyle(el).marginTop) || 0;
      const prev = el.previousElementSibling;
      const prevMb = prev ? parseFloat(getComputedStyle(prev).marginBottom) || 0 : 0;
      pushes.push({ el, expected: top + delta, amount: Math.max(mt, prevMb) + delta });
      shift += delta;
    };

    const visit = (container: Element, startAt: Element | null = null) => {
      let list = Array.from(container.children) as HTMLElement[];
      if (startAt) list = list.slice(list.indexOf(startAt as HTMLElement));
      for (const el of list) {
        if (el.offsetParent === null && el.tagName !== 'TR') continue; // display:none
        if (el.classList.contains(SPACER) || SKIP.has(el.tagName)) continue;
        if (el.classList.contains('we-page-break')) {
          const r = rect(el);
          forcedTop = (Math.floor((r.top - EPS) / g.stride) + 1) * g.stride;
          continue;
        }

        let r = rect(el);
        if (forcedTop >= 0) {
          if (r.top < forcedTop - EPS) {
            push(el, forcedTop - r.top, r.top);
            r = rect(el);
          }
          forcedTop = -1;
        }

        const page = Math.floor((r.top + EPS) / g.stride);
        const local = r.top - page * g.stride;
        const pageBottom = page * g.stride + g.avail;

        // Bắt đầu trong vùng lề dưới / khe giữa trang → đẩy xuống trang sau.
        if (local > g.avail - EPS) {
          push(el, (page + 1) * g.stride - r.top, r.top);
          continue;
        }
        if (r.bottom <= pageBottom + EPS) {
          // Giữ với đoạn sau (keepNext, đề mục): cả chuỗi phải nằm trên một trang, không thì đẩy từ đầu chuỗi.
          if (local > EPS && keepsNext(el)) {
            const need = keepEnd(el, r.bottom, g.avail, rect);
            if (need > pageBottom + EPS && need - r.top <= g.avail) push(el, (page + 1) * g.stride - r.top, r.top);
          }
          continue;
        }

        // Cắt qua đáy trang: bảng/danh sách thì ngắt giữa các hàng/mục như Word.
        if (DESCEND.has(el.tagName) && hasBlockChildren(el)) {
          visit(el);
          continue;
        }
        if (r.height <= g.avail && local > EPS) push(el, (page + 1) * g.stride - r.top, r.top);
        // Khối cao hơn cả một trang (ảnh lớn, hàng bảng rất cao): để tràn, giống Word khi không thể ngắt.
      }
    };

    visit(this.body, start);
    this.apply(pushes);
    this.verify(pushes, scale);
    this.pushes.push(...pushes);

    const height = this.body.scrollHeight;
    this.height = this.body.getBoundingClientRect().height / scale;
    this.valid = true;
    return (this.pages = Math.max(1, Math.floor((height - EPS) / g.stride) + 1));
  }

  /** Gỡ mọi lần đẩy của `start` và các khối đứng sau nó (giữ những lần đẩy phía trước). */
  private drop(start: Element) {
    const after = (el: Element) => el === start || start.contains(el) || !!(start.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING);
    let n = this.pushes.length;
    while (n > 0 && after(this.pushes[n - 1].el)) {
      const p = this.pushes[--n];
      if (p.row) for (const node of p.nodes ?? []) node.remove();
      else {
        p.el.removeAttribute(ATTR);
        if (p.id) this.rules.delete(p.id);
      }
    }
    this.pushes.length = n;
    this.sheet.textContent = [...this.rules.values()].join('\n');
  }

  /** Ghi toàn bộ khoảng đẩy vào DOM/CSS trong một lần. */
  private apply(pushes: Push[]) {
    for (const p of pushes) {
      if (p.row) this.insertGap(p);
      else {
        const id = String(++this.seq);
        p.id = id;
        p.el.setAttribute(ATTR, id);
        this.rules.set(id, this.rule(id, p.amount));
      }
    }
    this.sheet.textContent = [...this.rules.values()].join('\n');
  }

  private rule(id: string, amount: number) {
    return `${this.scope} [${ATTR}="${id}"] { margin-top: ${amount}px !important; }`;
  }

  /** Hàng đệm cao `amount` rồi lặp lại hàng tiêu đề (nếu có) ngay trước hàng bị đẩy. */
  private insertGap(p: Push) {
    const { table, head } = p.row!;
    const row = p.el;
    const cols = Math.max(1, ...Array.from(table.rows).map(r => Array.from(r.cells).reduce((n, c) => n + Math.max(1, c.colSpan), 0)));
    const gap = document.createElement('tr');
    gap.className = SPACER;
    gap.contentEditable = 'false';
    gap.setAttribute('aria-hidden', 'true');
    const td = document.createElement('td');
    td.colSpan = cols;
    td.style.cssText = `height:${p.amount}px;padding:0;border:none;background:transparent`;
    gap.append(td);
    row.before(gap);
    p.gap = td;
    p.nodes = [gap];
    for (const h of head) {
      const copy = h.cloneNode(true) as HTMLTableRowElement;
      copy.classList.add(SPACER, 'we-pb-head');
      copy.contentEditable = 'false';
      copy.setAttribute('aria-hidden', 'true');
      copy.removeAttribute(ATTR);
      row.before(copy);
      p.nodes.push(copy);
    }
  }

  /**
   * Kiểm tra vị trí thực tế so với dự kiến (margin gộp có thể làm lệch), chỉnh lại và đọc lại.
   * Mỗi vòng chỉ tốn một lần dàn trang vì mọi lần đọc diễn ra sau khi đã ghi hết.
   */
  private verify(pushes: Push[], scale: number) {
    for (let round = 0; round < 8 && pushes.length; round++) {
      let drift = 0;
      let changed = false;
      // Đọc lại gốc toạ độ mỗi vòng: cuộn của vùng chứa có thể đã đổi từ lúc đo ban đầu (trình duyệt tự
      // neo cuộn khi nội dung phía trên đổi chiều cao), nếu dùng gốc cũ thì mọi vị trí lệch cùng một đoạn.
      const origin = this.body.getBoundingClientRect().top;
      for (let i = 0; i < pushes.length; i++) {
        const p = pushes[i];
        const actual = (p.el.getBoundingClientRect().top - origin) / scale;
        // Với hàng bảng, hàng gốc nằm sau các bản sao tiêu đề; đỉnh hàng gốc đã là vị trí cuối.
        const err = actual - p.expected;
        const fix = err - drift;
        drift = err;
        if (Math.abs(fix) <= 0.02) continue; // rất chặt: sai số từng lần đẩy sẽ cộng dồn xuống các trang sau
        p.amount = Math.max(0, p.amount - fix);
        if (p.gap) p.gap.style.height = `${p.amount}px`;
        else this.rules.set(p.id!, this.rule(p.id!, p.amount));
        changed = true;
      }
      if (!changed) return;
      this.sheet.textContent = [...this.rules.values()].join('\n');
    }
  }
}

/** Đoạn có keepNext: khai báo `break-after: avoid` (từ .docx), hoặc là đề mục (như kiểu Heading của Word) nếu không khai báo gì. */
function keepsNext(el: HTMLElement): boolean {
  const v = el.style.breakAfter;
  return v ? v.includes('avoid') : /^H[1-6]$/.test(el.tagName);
}

/**
 * Đáy của phần cần nằm cùng trang với `el` khi đi theo chuỗi keepNext: khối kế tiếp là đoạn thì cả đoạn,
 * là bảng thì tới hàng đầu tiên sau hàng tiêu đề, là danh sách thì mục đầu. Khối kế tiếp cao hơn một trang
 * thì không thể giữ, dừng chuỗi ở đó.
 */
function keepEnd(el: HTMLElement, bottom: number, avail: number, rect: (e: Element) => { top: number; bottom: number; height: number }): number {
  let end = bottom;
  let cur = el;
  for (let i = 0; i < 10 && keepsNext(cur); i++) {
    const next = cur.nextElementSibling as HTMLElement | null;
    if (!next || next.classList.contains('we-page-break') || next.classList.contains(SPACER) || SKIP.has(next.tagName)) break;
    let target: Element = next;
    if (next.tagName === 'TABLE') {
      const rows = (next as HTMLTableElement).rows;
      target = rows[Math.min(rows.length - 1, (next as HTMLTableElement).tHead?.rows.length ?? 0)] ?? next;
    } else if (next.tagName === 'UL' || next.tagName === 'OL') target = next.querySelector('li') ?? next;
    const r = rect(target);
    if (r.bottom - rect(el).top > avail) break;
    end = r.bottom;
    cur = next;
  }
  return end;
}

function hasBlockChildren(el: HTMLElement) {
  return Array.from(el.children).some(c => CHILD_BLOCKS.test(c.tagName));
}
