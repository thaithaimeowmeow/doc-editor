/**
 * CSS của nội dung tài liệu (không phải giao diện editor). Dùng chung cho: vùng soạn thảo,
 * header/footer, bản in và file HTML xuất ra, để cả 4 nơi hiển thị giống nhau.
 * Mọi selector nằm dưới `.we-doc`.
 */
export const CONTENT_CSS = `
.we-doc {
  font-family: 'Times New Roman', Times, serif;
  font-size: 12pt;
  line-height: normal;
  color: #000;
  text-align: start;
  overflow-wrap: break-word;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.we-doc p, .we-doc li { margin: 0; }
.we-doc h1, .we-doc h2, .we-doc h3, .we-doc h4, .we-doc h5, .we-doc h6 { font-family: inherit; color: #2f5496; }
.we-doc h1 { font-size: 16pt; font-weight: bold; margin: 12pt 0 3pt; }
.we-doc h2 { font-size: 13pt; font-weight: bold; margin: 10pt 0 2pt; }
.we-doc h3 { font-size: 12pt; font-weight: bold; margin: 8pt 0 2pt; color: #1f3763; }
.we-doc h4 { font-size: 12pt; font-weight: bold; font-style: italic; margin: 8pt 0 2pt; }
.we-doc h5, .we-doc h6 { font-size: 12pt; font-weight: normal; margin: 6pt 0 0; }
.we-doc .we-title { font-size: 26pt; margin: 0 0 6pt; letter-spacing: -0.5pt; }
.we-doc .we-subtitle { font-size: 14pt; color: #595959; margin: 0 0 8pt; }
.we-doc blockquote { margin: 10pt 43pt; font-style: italic; color: #404040; }
.we-doc pre { margin: 0; font-family: 'Courier New', Courier, monospace; font-size: 10pt; white-space: pre-wrap; }
.we-doc ul, .we-doc ol { margin: 0; padding-left: 0.5in; }
.we-doc ul ul, .we-doc ol ol, .we-doc ul ol, .we-doc ol ul { padding-left: 0.5in; }
.we-doc a { color: #0563c1; text-decoration: underline; }
.we-doc img { max-width: 100%; vertical-align: bottom; }
.we-doc hr { border: 0; border-top: 1px solid #000; margin: 6pt 0; }
.we-doc table { border-collapse: collapse; }
.we-doc td, .we-doc th { vertical-align: top; padding: 0 5.4pt; }
.we-doc th { font-weight: bold; }
.we-doc sub, .we-doc sup { font-size: 75%; line-height: 0; }
.we-doc .we-page-break { display: block; height: 0; margin: 0; padding: 0; border: 0; overflow: hidden; }
.we-doc .we-field { white-space: nowrap; }
`;
