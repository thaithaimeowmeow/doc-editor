/**
 * Đổi các đoạn văn "giả danh sách" (Word vẽ số/bullet bằng ký tự hoặc CSS counter) thành <ol>/<ul> thật.
 * Dùng chung cho import .docx và dán từ Word.
 */
export interface ListItem {
  /** Mã danh sách: các đoạn cùng key + level nối tiếp nhau thuộc cùng một danh sách. */
  key: string;
  level: number;
  ordered: boolean;
  /** Giá trị CSS list-style-type. */
  styleType: string;
}

export function buildLists(items: Map<HTMLElement, ListItem>) {
  const parents = new Set(Array.from(items.keys(), p => p.parentElement!));
  const counts = new Map<string, number>(); // key-level → số mục đã gặp, để <ol start> nối tiếp như Word

  for (const parent of parents) {
    let stack: { list: HTMLElement; item: ListItem }[] = [];

    for (const child of Array.from(parent.children) as HTMLElement[]) {
      const item = items.get(child);
      if (!item) {
        stack = [];
        continue;
      }

      while (stack.length) {
        const top = stack.at(-1)!.item;
        if (top.level > item.level || (top.level === item.level && (top.key !== item.key || top.ordered !== item.ordered))) stack.pop();
        else break;
      }

      const countKey = `${item.key}-${item.level}`;
      const n = (counts.get(countKey) ?? 0) + 1;
      counts.set(countKey, n);
      for (const k of Array.from(counts.keys())) {
        const [key, level] = k.split('-');
        if (key === item.key && Number(level) > item.level) counts.delete(k); // mục cấp trên reset cấp dưới
      }

      let top = stack.at(-1);
      if (!top || top.item.level < item.level) {
        const list = document.createElement(item.ordered ? 'ol' : 'ul');
        if (item.styleType !== (item.ordered ? 'decimal' : 'disc')) list.style.listStyleType = item.styleType;
        if (item.ordered && n > 1) list.setAttribute('start', String(n));
        if (top) (top.list.lastElementChild ?? top.list).append(list);
        else child.before(list);
        top = { list, item };
        stack.push(top);
      }

      const li = document.createElement('li');
      const style = child.getAttribute('style');
      if (style) li.setAttribute('style', style);
      li.style.removeProperty('margin-left'); // thụt lề do <ol>/<ul> lo
      li.style.removeProperty('text-indent');
      if (!li.getAttribute('style')) li.removeAttribute('style');
      li.append(...Array.from(child.childNodes));
      child.remove();
      top.list.append(li);
    }
  }
}

/** Đoán kiểu danh sách từ ký tự đánh dấu ("1.", "a)", "•", "○"...). */
export function markerToStyle(marker: string): { ordered: boolean; styleType: string } {
  const m = marker.trim().replace(/ /g, '');
  const core = /^[([]?([0-9]+|[ivxlcdm]+|[IVXLCDM]+|[a-z]|[A-Z])[.)\]]/.exec(m)?.[1];
  if (core) {
    if (/^\d+$/.test(core)) return { ordered: true, styleType: 'decimal' };
    if (/^[ivxlcdm]+$/.test(core) && core.length > 1) return { ordered: true, styleType: 'lower-roman' };
    if (/^[IVXLCDM]+$/.test(core) && core.length > 1) return { ordered: true, styleType: 'upper-roman' };
    if (/^[a-z]$/.test(core)) return { ordered: true, styleType: 'lower-alpha' };
    return { ordered: true, styleType: 'upper-alpha' };
  }
  if (/^[o○◦]$/.test(m)) return { ordered: false, styleType: 'circle' };
  if (/^[§■▪]$/.test(m)) return { ordered: false, styleType: 'square' };
  return { ordered: false, styleType: 'disc' };
}
