import { SerializedSelection, restoreSelection, serializeSelection } from './dom';

interface Snapshot {
  /** Bản clone DOM thay vì chuỗi HTML: trình duyệt dùng chung bộ nhớ cho chuỗi base64 của ảnh. */
  content: DocumentFragment;
  selection: SerializedSelection | null;
}

/** Undo/redo tự quản lý, không dùng stack gốc của trình duyệt (vốn bị rối khi DOM đổi bằng code). */
export class History {
  private stack: Snapshot[] = [];
  private index = -1;
  private dirty = false;

  constructor(
    private readonly root: HTMLElement,
    private readonly clean: (fragment: DocumentFragment) => void,
    private readonly limit = 200,
  ) {}

  get canUndo() {
    return this.index > 0 || this.dirty;
  }
  get canRedo() {
    return !this.dirty && this.index < this.stack.length - 1;
  }

  markDirty() {
    this.dirty = true;
  }

  /** Lưu trạng thái hiện tại nếu có thay đổi. */
  commit(force = false) {
    if (!this.dirty && !force && this.stack.length) {
      // Không đổi nội dung, chỉ cập nhật vị trí con trỏ.
      const top = this.stack[this.index];
      if (top) top.selection = serializeSelection(this.root);
      return;
    }
    this.stack.length = this.index + 1;
    this.stack.push(this.snapshot());
    if (this.stack.length > this.limit) this.stack.shift();
    this.index = this.stack.length - 1;
    this.dirty = false;
  }

  reset() {
    this.stack = [];
    this.index = -1;
    this.dirty = false;
    this.commit(true);
  }

  undo(): boolean {
    if (this.dirty) this.commit();
    if (this.index <= 0) return false;
    this.index--;
    this.apply(this.stack[this.index]);
    return true;
  }

  redo(): boolean {
    if (this.dirty || this.index >= this.stack.length - 1) return false;
    this.index++;
    this.apply(this.stack[this.index]);
    return true;
  }

  private snapshot(): Snapshot {
    const content = document.createDocumentFragment();
    for (const n of Array.from(this.root.childNodes)) content.append(n.cloneNode(true));
    this.clean(content);
    return { content, selection: serializeSelection(this.root) };
  }

  private apply(s: Snapshot) {
    this.root.replaceChildren(s.content.cloneNode(true));
    restoreSelection(this.root, s.selection);
    this.dirty = false;
  }
}
