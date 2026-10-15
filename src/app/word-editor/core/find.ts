/** Tìm & thay thế. Tô sáng bằng CSS Custom Highlight API (không sửa DOM) nếu trình duyệt hỗ trợ. */

export interface FindOptions {
  matchCase: boolean;
  wholeWord: boolean;
}

interface HighlightRegistry {
  set(name: string, h: unknown): void;
  delete(name: string): void;
}
type HighlightCtor = new (...ranges: Range[]) => unknown;

const registry = (globalThis.CSS as unknown as { highlights?: HighlightRegistry } | undefined)?.highlights;
const HighlightClass = (globalThis as unknown as { Highlight?: HighlightCtor }).Highlight;

export class Finder {
  matches: Range[] = [];
  current = -1;

  constructor(
    private readonly root: HTMLElement,
    private readonly id: string,
  ) {}

  search(query: string, opts: FindOptions): number {
    this.matches = [];
    this.current = -1;
    if (query) {
      const { text, map } = this.flatten();
      const flags = opts.matchCase ? 'gu' : 'giu';
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(opts.wholeWord ? `(?<![\\p{L}\\p{N}_])${escaped}(?![\\p{L}\\p{N}_])` : escaped, flags);
      for (const m of text.matchAll(re)) {
        if (!m[0].length) continue;
        const r = document.createRange();
        const s = map(m.index!);
        const e = map(m.index! + m[0].length - 1);
        r.setStart(s.node, s.offset);
        r.setEnd(e.node, e.offset + 1);
        this.matches.push(r);
      }
    }
    this.paint();
    return this.matches.length;
  }

  /** Chọn kết quả tiếp theo (hoặc trước đó) tính từ vị trí con trỏ. */
  next(backward = false): Range | null {
    if (!this.matches.length) return null;
    if (this.current < 0) {
      const sel = getSelection();
      const caret = sel?.rangeCount ? sel.getRangeAt(0) : null;
      const idx = caret ? this.matches.findIndex(m => m.compareBoundaryPoints(Range.START_TO_START, caret) >= 0) : 0;
      this.current = backward ? (idx <= 0 ? this.matches.length - 1 : idx - 1) : Math.max(0, idx);
    } else {
      this.current = (this.current + (backward ? -1 : 1) + this.matches.length) % this.matches.length;
    }
    return this.select();
  }

  select(): Range | null {
    const r = this.matches[this.current];
    if (!r) return null;
    const sel = getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r.cloneRange());
    this.paint();
    return r;
  }

  /** Thay kết quả hiện tại. Trả về true nếu đã thay. */
  replaceCurrent(replacement: string): boolean {
    const r = this.matches[this.current];
    if (!r) return false;
    const sel = getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r.cloneRange());
    document.execCommand('insertText', false, replacement);
    return true;
  }

  replaceAll(query: string, replacement: string, opts: FindOptions): number {
    const count = this.search(query, opts);
    // Thay từ cuối lên đầu để các Range phía trước không bị lệch.
    for (let i = this.matches.length - 1; i >= 0; i--) {
      const r = this.matches[i];
      r.deleteContents();
      if (replacement) r.insertNode(document.createTextNode(replacement));
    }
    this.root.normalize();
    this.clear();
    return count;
  }

  clear() {
    this.matches = [];
    this.current = -1;
    registry?.delete(`${this.id}-all`);
    registry?.delete(`${this.id}-cur`);
  }

  private paint() {
    if (!registry || !HighlightClass) return;
    registry.set(`${this.id}-all`, new HighlightClass(...this.matches));
    const cur = this.matches[this.current];
    if (cur) registry.set(`${this.id}-cur`, new HighlightClass(cur));
    else registry.delete(`${this.id}-cur`);
  }

  /** Ghép text của tài liệu (mỗi khối cách nhau bằng \n) kèm hàm đổi vị trí → (node, offset). */
  private flatten() {
    const nodes: { node: Text; start: number }[] = [];
    let text = '';
    const walker = document.createTreeWalker(this.root, NodeFilter.SHOW_TEXT);
    let lastBlock: Node | null = null;
    for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) {
      const block = n.parentElement?.closest('p,li,td,th,h1,h2,h3,h4,h5,h6,div,pre,blockquote') ?? null;
      if (lastBlock && block !== lastBlock) text += '\n';
      lastBlock = block;
      nodes.push({ node: n, start: text.length });
      text += n.data;
    }
    const map = (i: number) => {
      let lo = 0, hi = nodes.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (nodes[mid].start <= i) lo = mid;
        else hi = mid - 1;
      }
      const n = nodes[lo];
      return { node: n.node, offset: Math.min(i - n.start, n.node.length - 1) };
    };
    return { text, map };
  }
}
