import { Component, computed, signal, viewChild } from '@angular/core';

import { WordDocument } from './word-editor/core/engine';
import { DocxSource, WordEditor, WordEmitKind } from './word-editor/word-editor';

type OutputKind = 'html' | 'htmlBase64' | 'docxBase64' | 'document';

/** Trang demo: truyền dữ liệu vào bằng input, nhận dữ liệu ra (đã debounce) bằng output. */
@Component({
  selector: 'app-root',
  imports: [WordEditor],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  private readonly editor = viewChild.required(WordEditor);
  protected readonly html = signal('<p>Xin chào! Mở file .docx bằng nút bên phải hoặc tab <b>Tệp</b>.</p>');
  protected readonly docx = signal<DocxSource | null>(null);
  protected readonly readonly = signal(false);
  protected readonly emit: WordEmitKind[] = ['html', 'htmlBase64', 'docxBase64', 'document'];

  protected readonly htmlBase64 = signal('');
  protected readonly docxBase64 = signal('');
  protected readonly doc = signal<WordDocument | null>(null);
  protected readonly kind = signal<OutputKind>('html');
  protected readonly counts = signal<Record<OutputKind, number>>({ html: 0, htmlBase64: 0, docxBase64: 0, document: 0 });
  protected readonly status = signal('');

  protected readonly tabs: { kind: OutputKind; label: string }[] = [
    { kind: 'html', label: 'html' },
    { kind: 'htmlBase64', label: 'htmlBase64' },
    { kind: 'docxBase64', label: 'docxBase64' },
    { kind: 'document', label: 'document' },
  ];

  protected readonly output = computed(() => {
    switch (this.kind()) {
      case 'html':
        return this.html();
      case 'htmlBase64':
        return this.htmlBase64();
      case 'docxBase64':
        return this.docxBase64();
      case 'document': {
        const d = this.doc();
        return d ? JSON.stringify({ ...d, html: `(${d.html.length} ký tự)` }, null, 2) : '';
      }
    }
  });

  protected count(kind: OutputKind) {
    this.counts.update(c => ({ ...c, [kind]: c[kind] + 1 }));
  }

  protected onHtmlBase64(v: string) {
    this.htmlBase64.set(v);
    this.count('htmlBase64');
  }
  protected onDocxBase64(v: string) {
    this.docxBase64.set(v);
    this.count('docxBase64');
  }
  protected onDocument(v: WordDocument) {
    this.doc.set(v);
    this.count('document');
  }

  protected pickDocx(e: Event) {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    input.value = '';
    if (file) this.docx.set(file); // truyền qua input [docx]
  }

  /** Chế độ manual: chỉ tạo html / base64 / .docx khi bấm nút, để gõ không bị giật. */
  protected fetchOutput() {
    return this.editor().emitNow();
  }

  protected async copy() {
    await navigator.clipboard.writeText(this.output());
    this.status.set('Đã copy');
    setTimeout(() => this.status.set(''), 1500);
  }

  protected size(text: string) {
    const bytes = new Blob([text]).size;
    return bytes < 1024 ? `${bytes} B` : bytes < 1024 ** 2 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1024 ** 2).toFixed(1)} MB`;
  }
}
