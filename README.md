# Word Editor (Angular)

Trình soạn thảo kiểu Microsoft Word, chạy hoàn toàn trên trình duyệt, không cần server.
Tự viết engine, chỉ phụ thuộc 3 thư viện miễn phí cho thương mại.

```bash
npm install
npm start          # http://localhost:4200
```

File mẫu để thử: `samples/bien-ban-mau.docx`, `samples/hop-dong-nhieu-trang.docx`.

## Tính năng

| Nhóm | Có |
|---|---|
| Giao diện | Ribbon kiểu Word (Tệp, Trang chủ, Chèn, Bố cục, Xem). Tab ngữ cảnh: Bảng, Ảnh, Đầu trang & Chân trang. Thanh truy cập nhanh, thước kẻ, thanh trạng thái, zoom |
| Trang | Trang A4 thật trên nền xám, **tự chia trang**, khổ giấy (A3/A4/A5/Letter/Legal), hướng dọc/ngang, lề (có mẫu văn bản hành chính VN), ngắt trang |
| Header/Footer | Nháy đúp để sửa, **số trang / tổng số trang**, trang đầu khác |
| Ký tự | Phông, cỡ chữ, tăng/giảm cỡ, đậm/nghiêng/gạch chân/gạch ngang, chỉ số trên/dưới, màu chữ (bảng màu Word), tô sáng, đổi hoa/thường, xoá định dạng, sao chép định dạng |
| Đoạn | Căn lề 4 kiểu, giãn dòng, khoảng cách trước/sau, thụt lề (nút + thước kẻ), tô nền đoạn, danh sách dấu đầu dòng/đánh số nhiều cấp (Tab/Shift+Tab), kiểu: Bình thường, Tiêu đề, Phụ đề, Đề mục 1–4, Trích dẫn, Mã |
| Bảng | Giữ đúng độ rộng cột của Word (tblGrid), hàng tiêu đề lặp lại đầu mỗi trang, lề ô, viền 0,5pt. Chèn bằng lưới, thêm/xoá hàng/cột, **gộp/tách ô** (kéo chuột chọn nhiều ô), kéo đổi độ rộng cột, chia đều cột, tô nền ô, viền, căn dọc, Tab sang ô kế tiếp |
| Ảnh | Chèn từ máy / URL / dán / kéo thả, kéo góc đổi kích thước, căn lề, văn bản thay thế |
| Khác | Liên kết (Ctrl+click để mở), ký hiệu, ngày, đường kẻ ngang, tìm & thay thế (khớp hoa/thường, cả từ), undo/redo riêng, ký tự định dạng ¶, đường lưới bảng, bố cục web |
| Dán | Từ Word (giữ định dạng, đổi danh sách Word thành list thật), Google Docs, web, ảnh |
| File | Mở .docx/.html/.txt, lưu .docx, tải .html, in / xuất PDF (bản in khớp màn hình, có header/footer) |
| Dữ liệu ra | HTML, base64 HTML, base64 .docx, Blob .docx, toàn bộ tài liệu (`WordDocument`) |

Phím tắt: Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+E/L/R/J, Ctrl+M / Ctrl+Shift+M, Ctrl+K, Ctrl+F/H, Ctrl+S, Ctrl+O, Ctrl+P,
Ctrl+Enter (ngắt trang), Ctrl+]/[ (cỡ chữ), Ctrl+= / Ctrl+Shift+= (chỉ số), Ctrl+Alt+1..4 (đề mục),
Ctrl+Shift+L (danh sách), Ctrl+Shift+C (sao chép định dạng), Ctrl+Space (xoá định dạng), Ctrl+1/2/5 (giãn dòng), Ctrl+cuộn (zoom).

## Bốc sang project khác

1. Copy nguyên thư mục `src/app/word-editor/`.
2. Cài dependency:

   ```bash
   npm i docx-preview jszip dompurify
   ```

3. `angular.json`:
   - `architect.build.options.allowedCommonJsDependencies`: thêm `"jszip"`.
   - `budgets` loại `anyComponentStyle`: nâng `maximumError` lên khoảng `40kB`, vì CSS ribbon khoảng 17 KB.
4. Dùng:

```html
<!-- Component cần một chiều cao xác định -->
<app-word-editor
  style="display:block; height: 100vh"
  [(html)]="html"
  [docx]="file"
  [readonly]="false"
  [emit]="['html', 'htmlBase64', 'docxBase64', 'document']"
  [debounce]="300"
  [docxDebounce]="1500"
  (htmlBase64Change)="htmlB64 = $event"
  (docxBase64Change)="docxB64 = $event"
  (documentChange)="doc = $event"
/>
```

### Input

| Input | Kiểu | Mô tả |
|---|---|---|
| `html` | `string` | Nội dung HTML. Dùng hai chiều `[(html)]`. Giá trị editor vừa phát ra sẽ không bị đẩy ngược vào lại, nên không mất con trỏ hay undo. |
| `document` | `Partial<WordDocument>` | Toàn bộ tài liệu: HTML, header/footer, khổ giấy, lề, tiêu đề |
| `docx` | `Blob \| File \| ArrayBuffer \| Uint8Array \| string` | Mở file .docx. `string` là base64, có hoặc không có tiền tố `data:` |
| `readonly` | `boolean` | Chỉ xem. Vẫn chọn, copy, tìm, in, zoom được |
| `emit` | `('html' \| 'htmlBase64' \| 'docxBase64' \| 'document')[]` | Những output được phát. Mặc định `['html']` |
| `debounce` | `number` (ms) | Chờ sau lần sửa cuối rồi mới phát html / htmlBase64 / document. Mặc định 300 |
| `docxDebounce` | `number` (ms) | Chờ rồi mới tạo và phát docxBase64. Mặc định 1500. Nếu đang tạo mà người dùng sửa tiếp thì bỏ kết quả cũ |
| `saveMode` | `'download' \| 'emit'` | Ctrl+S / nút Lưu: tải .docx, hoặc chỉ phát `save` |
| `showRuler` | `boolean` | Hiện thước kẻ |

### Output

Tất cả output nội dung đều đã debounce. `html` và `htmlBase64` chỉ phát khi HTML thực sự thay đổi.

| Output | Giá trị |
|---|---|
| `htmlChange` | `string`: HTML thân văn bản |
| `htmlBase64Change` | `string`: base64 UTF-8 của HTML |
| `docxBase64Change` | `string`: base64 của file .docx, không có tiền tố `data:` |
| `documentChange` | `WordDocument`: phát cả khi chỉ sửa header/footer, khổ giấy hoặc tiêu đề |
| `save` | `WordDocument` khi bấm Lưu / Ctrl+S |
| `fileLoaded` | Tên file vừa mở |
| `failed` | Lỗi khi mở/xuất |

Vẫn dùng được `[formControl]` / `[(ngModel)]` (giá trị là HTML, cũng đã debounce).

### Gọi trực tiếp (qua `viewChild`)

```ts
import { WordEditor } from './word-editor/word-editor';

const ed = this.editor(); // viewChild.required(WordEditor)

ed.getHtml();                 // HTML thân văn bản (đã sanitize)
ed.setHtml('<p>…</p>');
await ed.getHtmlBase64();     // base64 UTF-8 của HTML
await ed.getDocxBase64();     // base64 của .docx (không có tiền tố data:)
await ed.getDocxBlob();
await ed.downloadDocx();
ed.downloadHtml();
ed.print();
await ed.loadFile(file);      // .docx | .html | .txt
ed.openFilePicker();
ed.getDocument();             // { html, header, footer, firstHeader, firstFooter, differentFirstPage, page, title }
ed.setDocument({ ... });      // khôi phục đầy đủ, kể cả header/footer & khổ giấy
```


Giá trị `html` / `formControl` chỉ là HTML thân văn bản. Muốn lưu cả header/footer và khổ giấy thì
lưu `documentChange` / `getDocument()` rồi truyền lại qua `[document]`, hoặc lưu file .docx.

## Cấu trúc

```
word-editor/
  word-editor.ts/.html/.css   Component Angular: ribbon, hộp thoại, thước kẻ, thanh trạng thái
  we-icon.ts, icons.ts        Bộ icon SVG tự vẽ
  core/                       Không phụ thuộc Angular, dùng lại được cho React/Vue/JS thuần
    engine.ts                 Lệnh soạn thảo, phím tắt, dán, bảng, ảnh, header/footer, zoom
    paginator.ts              Chia trang tự động
    history.ts                Undo/redo
    table.ts                  Thao tác bảng (rowspan/colspan)
    find.ts                   Tìm & thay (CSS Highlight API)
    docx-import.ts            .docx → HTML (docx-preview + đọc XML header/footer)
    docx-export.ts            HTML → .docx (tự sinh OOXML, nén bằng jszip)
    print.ts                  In / HTML độc lập
    sanitize.ts, lists.ts     Làm sạch HTML, đổi danh sách Word
    content-css.ts            CSS nội dung dùng chung cho editor, bản in, HTML xuất ra
```

## Dung lượng (build production, gzip)

| Phần | Kích thước |
|---|---|
| Editor + DOMPurify (trong bundle chính) | khoảng 35 KB |
| Mở .docx: docx-preview + docx-import | khoảng 23 KB, tải khi cần |
| Lưu .docx: docx-export | khoảng 8 KB, tải khi cần |
| jszip (dùng chung cho mở và lưu) | khoảng 27 KB, tải khi cần |

Bundle chính của app demo (gồm cả Angular) là 91 KB.

## License (miễn phí cho thương mại)

| Package | License |
|---|---|
| docx-preview | Apache-2.0 |
| jszip | MIT hoặc GPL-3.0 (chọn MIT) |
| dompurify | MPL-2.0 hoặc Apache-2.0 (chọn Apache-2.0) |

Các phụ thuộc gián tiếp đều là MIT, ISC, 0BSD hoặc Zlib.

## Hiệu năng

Đo trên Chromium, máy phát triển, build dev. Số liệu là trung vị của nhiều lần chạy.

| Tài liệu | Thao tác | Trước | Sau |
|---|---|---|---|
| SRS 50 trang, 10,7 MB, 64 bảng, 33 ảnh | Chia trang lại (chạy sau mỗi lần gõ, đã gom theo khung hình) | 547–621 ms | **21–58 ms** |
| | Xuất .docx | 1,65 s | **0,57 s** |
| | Mở .docx | khoảng 1,2–1,7 s | khoảng 1,0–1,4 s (phần lớn là docx-preview) |
| | Sanitize 14,7 MB HTML | 170 ms | **37 ms** |
| | Độ trễ đồng bộ của một phím gõ | | 4,7 ms (trung vị) |
| Tài liệu tổng hợp 107 trang, 9 nghìn node | Chia trang lại / xuất .docx / một phím gõ | | 78 ms / 460 ms / 8 ms |

Đã làm:
- **Chia trang:** đo toàn bộ trong một lượt, theo dõi độ dịch cộng dồn bằng số học, ghi mọi khoảng đẩy một lần, rồi kiểm tra lại theo lô. Cách cũ xen kẽ đọc/ghi nên mỗi lần đẩy buộc trình duyệt dàn trang lại cả tài liệu (khoảng 50 lần). Vị trí khối giống thuật toán cũ (sai lệch tối đa 0,5 px, mới chính xác hơn) và không có khối nào vắt qua ranh giới trang.
- **Xuất .docx:** ảnh PNG/JPEG lưu nguyên (nén lại tốn 1,6 s mà chỉ nhỏ hơn 3%), giải mã base64 bằng bộ giải mã gốc của trình duyệt.
- **Sanitize:** thay ảnh base64 bằng chỗ giữ chỗ trong lúc DOMPurify chạy rồi trả lại.
- **Undo:** số snapshot tự giảm theo kích thước tài liệu (tài liệu 50 trang giữ khoảng 100 bước, khoảng 120 MB) thay vì cố định 200 bước (khoảng 230 MB).
- Bỏ qua `<colgroup>` khi chia trang (hộp của nó trùng cả bảng nên bị coi là một khối cần đẩy).

Chưa làm: chia trang tăng dần từ vị trí vừa sửa. Hiện mỗi lần gõ chia lại cả tài liệu, tỉ lệ thuận với số khối (tài liệu vài trăm trang sẽ thấy trễ khi gõ).

## Giới hạn đã biết

- **Chia trang theo đoạn:** một đoạn văn không bị cắt đôi qua hai trang như Word, mà được đẩy nguyên sang trang sau. Bảng ngắt theo hàng (có lặp hàng tiêu đề), danh sách ngắt theo mục. Một hàng bảng / đoạn / ảnh cao hơn cả trang sẽ tràn xuống phần lề.
- Không có: chú thích cuối trang, mục lục tự động, text box, ảnh nổi (wrap text), track changes, comment, cột báo, nhiều section với khổ giấy khác nhau.
- Lệnh định dạng dùng `document.execCommand`. API này đã bị đánh dấu deprecated nhưng mọi trình duyệt chính vẫn hỗ trợ.
- Tô sáng kết quả tìm kiếm cần CSS Custom Highlight API (Chrome 105+, Safari 17.2+, Firefox 140+). Trình duyệt cũ vẫn tìm và chọn được, chỉ không tô sáng tất cả kết quả.
- File .docx xuất ra đã kiểm tra XML hợp lệ và mở lại đúng bằng docx-preview, **chưa kiểm tra bằng Microsoft Word / LibreOffice**.
