/**
 * Bộ icon 20×20, nét `currentColor`. Viết tay để không phụ thuộc thư viện icon.
 * Mỗi giá trị là phần bên trong thẻ <svg>.
 */
const L = (x1: number, y1: number, x2: number, y2: number) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"/>`;
const lines = (ys: number[], x1: number, x2: number) => ys.map(y => L(x1, y, x2, y)).join('');

export const ICONS: Record<string, string> = {
  // ---- Clipboard ----
  paste:
    '<rect x="4" y="3.5" width="12" height="14" rx="1.5"/><rect x="7" y="2" width="6" height="3" rx="1" fill="currentColor" fill-opacity=".15"/>' +
    lines([9, 12, 15], 7, 13),
  cut: '<circle cx="6" cy="14.5" r="2.5"/><circle cx="14" cy="14.5" r="2.5"/>' + L(7.5, 12.5, 14, 3) + L(12.5, 12.5, 6, 3),
  copy: '<rect x="6.5" y="6.5" width="10" height="11" rx="1.5"/><path d="M13.5 6.5V4a1.5 1.5 0 0 0-1.5-1.5H5A1.5 1.5 0 0 0 3.5 4v8.5A1.5 1.5 0 0 0 5 14h1.5"/>',
  painter: '<rect x="3" y="3" width="12" height="5" rx="1"/><path d="M15 5.5h2v4.5h-7.5v2.5"/><rect x="8" y="12.5" width="3" height="5.5" rx=".8" fill="currentColor"/>',

  // ---- Font ----
  grow: '<path d="M2.5 16 7 4l4.5 12M4.3 12h5.4"/><path d="M13 9l2.5-3 2.5 3"/>',
  shrink: '<path d="M2.5 16 7 4l4.5 12M4.3 12h5.4"/><path d="M13 6l2.5 3 2.5-3"/>',
  case: '<path d="M1.5 15 5.5 5l4 10M3 11.5h5"/><path d="M15 15v-1.2M15 9.5c-1.7 0-3 .9-3 2.8s1.3 2.7 3 2.7V9.5c0-1.3-.8-2-2-2"/>',
  clearFormat: '<path d="M4 4h10M9 4l-3 12"/><path d="m12 12 5 5m0-5-5 5"/>',
  fontColor: '<path d="M5 13 10 2l5 11M6.8 9.5h6.4"/><rect class="we-swatch" x="2" y="15.5" width="16" height="3" stroke="none"/>',
  highlight: '<path d="m6 12 6.5-8.5 3.5 3L9.5 15H6z"/><path d="M6 12l3.5 3"/><rect class="we-swatch" x="2" y="15.5" width="16" height="3" stroke="none"/>',

  // ---- Paragraph ----
  bullets: '<circle cx="4" cy="5" r="1.2" fill="currentColor"/><circle cx="4" cy="10" r="1.2" fill="currentColor"/><circle cx="4" cy="15" r="1.2" fill="currentColor"/>' + lines([5, 10, 15], 8, 17),
  numbering:
    '<path d="M3 3.5h1.2V7M3 7h2.4" stroke-width="1.2"/><path d="M2.8 11.3a1.2 1.2 0 0 1 2.3.4c0 .9-2.3 1.8-2.3 2.8h2.4" stroke-width="1.2"/>' +
    lines([5, 12.5], 8, 17) + L(8, 17.5, 17, 17.5),
  multilevel: '<circle cx="3.5" cy="4.5" r="1" fill="currentColor"/><circle cx="7.5" cy="10" r="1" fill="currentColor"/><circle cx="11.5" cy="15.5" r="1" fill="currentColor"/>' + L(6, 4.5, 17, 4.5) + L(10, 10, 17, 10) + L(14, 15.5, 17, 15.5),
  outdent: lines([3, 17], 3, 17) + lines([8, 12], 9, 17) + '<path d="M6.5 7.5 3.5 10l3 2.5"/>',
  indent: lines([3, 17], 3, 17) + lines([8, 12], 9, 17) + '<path d="m3.5 7.5 3 2.5-3 2.5"/>',
  alignLeft: L(3, 4, 17, 4) + L(3, 8, 12, 8) + L(3, 12, 17, 12) + L(3, 16, 12, 16),
  alignCenter: L(3, 4, 17, 4) + L(5.5, 8, 14.5, 8) + L(3, 12, 17, 12) + L(5.5, 16, 14.5, 16),
  alignRight: L(3, 4, 17, 4) + L(8, 8, 17, 8) + L(3, 12, 17, 12) + L(8, 16, 17, 16),
  alignJustify: lines([4, 8, 12, 16], 3, 17),
  lineSpacing: lines([4.5, 10, 15.5], 9, 17) + '<path d="M4.5 3v14M2.5 5l2-2 2 2M2.5 15l2 2 2-2"/>',
  shading: '<path d="m4 9 5-5 6 6-5 5z"/><path d="M15.5 12.5s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3z" fill="currentColor"/><rect class="we-swatch" x="2" y="17" width="12" height="2" stroke="none"/>',
  marks: '<path d="M11 17V3M14.5 17V3M15 3H9a3.5 3.5 0 0 0 0 7h2"/>',
  borders: '<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.5 1.5"/>' + L(3, 17, 17, 17),

  // ---- Editing ----
  find: '<circle cx="8.5" cy="8.5" r="5"/>' + L(12.2, 12.2, 17, 17),
  replace: '<path d="M3 7.5h9.5M10 4.5l3 3-3 3M17 12.5H7.5M10 9.5l-3 3 3 3"/>',
  select: '<path d="M5 3v13l3.5-3.5 2.5 5 2-1-2.5-5H15z"/>',

  // ---- Insert ----
  pageBreak: '<path d="M5 2v5h10V2M5 18v-5h10v5"/>' + '<path d="M2 10h3m2 0h2m2 0h2m2 0h3" />',
  blankPage: '<path d="M5 2h7l4 4v12H5z"/><path d="M12 2v4h4"/>',
  table: '<rect x="2.5" y="3.5" width="15" height="13" rx="1"/>' + L(2.5, 8, 17.5, 8) + L(2.5, 12.3, 17.5, 12.3) + L(7.5, 3.5, 7.5, 16.5) + L(12.5, 3.5, 12.5, 16.5),
  image: '<rect x="2.5" y="3.5" width="15" height="13" rx="1.5"/><circle cx="7" cy="8" r="1.5"/><path d="m3 15 4.5-4.5 3 3 2-2 4.5 4"/>',
  link: '<path d="M8.5 11.5a3.5 3.5 0 0 0 5 0l2.5-2.5a3.5 3.5 0 0 0-5-5l-1 1"/><path d="M11.5 8.5a3.5 3.5 0 0 0-5 0L4 11a3.5 3.5 0 0 0 5 5l1-1"/>',
  unlink: '<path d="M8.5 11.5a3.5 3.5 0 0 0 5 0l2.5-2.5a3.5 3.5 0 0 0-5-5l-1 1M11.5 8.5a3.5 3.5 0 0 0-5 0L4 11a3.5 3.5 0 0 0 5 5l1-1"/>' + L(3, 3, 17, 17),
  hr: L(2, 10, 18, 10) + '<path d="M4 6h12M4 14h12" stroke-opacity=".35"/>',
  symbol: '<path d="M4 16.5h4v-1.8A5.5 5.5 0 0 1 10 4a5.5 5.5 0 0 1 2 10.7v1.8h4"/>',
  header: '<rect x="3" y="2.5" width="14" height="15" rx="1"/><rect x="5" y="4.5" width="10" height="3" fill="currentColor" fill-opacity=".25"/>' + lines([10.5, 13.5], 5, 15),
  footer: '<rect x="3" y="2.5" width="14" height="15" rx="1"/><rect x="5" y="12.5" width="10" height="3" fill="currentColor" fill-opacity=".25"/>' + lines([5.5, 8.5], 5, 15),
  pageNumber: '<rect x="3" y="2.5" width="14" height="15" rx="1"/><text x="10" y="14.5" text-anchor="middle" font-size="8" font-weight="700" fill="currentColor" stroke="none" font-family="Arial">#</text>',
  date: '<rect x="2.5" y="4" width="15" height="13" rx="1.5"/>' + L(2.5, 8, 17.5, 8) + L(6.5, 2, 6.5, 5.5) + L(13.5, 2, 13.5, 5.5) + '<rect x="6" y="10.5" width="3" height="3" fill="currentColor" stroke="none"/>',
  textField: '<rect x="2.5" y="5" width="15" height="10" rx="1"/>' + L(6, 8, 6, 12),

  // ---- Layout ----
  margins: '<rect x="3.5" y="2" width="13" height="16" rx="1"/><rect x="6" y="5" width="8" height="10" stroke-dasharray="1.5 1.2"/>',
  orientation: '<rect x="2" y="6" width="10" height="12" rx="1"/><rect x="8" y="2" width="10" height="7" rx="1" fill="currentColor" fill-opacity=".15"/>',
  size: '<path d="M5 2h7l4 4v12H5z"/><path d="M12 2v4h4"/><path d="M8 10h5M8 13h5" stroke-opacity=".6"/>',
  spaceBefore: '<path d="M6 7V3M4 5l2-2 2 2"/>' + lines([10, 13, 16], 3, 17),
  spaceAfter: lines([4, 7, 10], 3, 17) + '<path d="M6 13v4M4 15l2 2 2-2"/>',

  // ---- View ----
  printLayout: '<rect x="4" y="2" width="12" height="16" rx="1"/>' + lines([5.5, 8.5, 11.5], 6.5, 13.5),
  webLayout: '<rect x="2" y="3" width="16" height="14" rx="1.5"/>' + L(2, 6.5, 18, 6.5) + '<circle cx="4.2" cy="4.8" r=".6" fill="currentColor"/>',
  zoomIn: '<circle cx="8.5" cy="8.5" r="5"/>' + L(12.2, 12.2, 17, 17) + L(6, 8.5, 11, 8.5) + L(8.5, 6, 8.5, 11),
  zoomOut: '<circle cx="8.5" cy="8.5" r="5"/>' + L(12.2, 12.2, 17, 17) + L(6, 8.5, 11, 8.5),
  zoom100: '<circle cx="8.5" cy="8.5" r="5"/>' + L(12.2, 12.2, 17, 17) + '<text x="8.5" y="10.6" text-anchor="middle" font-size="5.5" fill="currentColor" stroke="none" font-family="Arial">1:1</text>',
  pageWidth: '<rect x="4" y="2" width="12" height="16" rx="1"/><path d="M2 10h16M4.5 8l-2 2 2 2M15.5 8l2 2-2 2"/>',
  onePage: '<rect x="5" y="2.5" width="10" height="15" rx="1"/><path d="M3 5v10M17 5v10" stroke-opacity=".4"/>',
  gridlines: '<rect x="2.5" y="2.5" width="15" height="15" stroke-dasharray="2 1.5"/>' + '<path d="M2.5 10h15M10 2.5v15" stroke-dasharray="2 1.5"/>',
  ruler: '<rect x="2" y="6" width="16" height="8" rx="1"/>' + [5, 8, 11, 14].map(x => L(x, 6, x, x % 2 ? 9 : 10)).join(''),

  // ---- Table ----
  rowAbove: '<rect x="2.5" y="9" width="15" height="8.5" rx="1"/>' + L(10, 9, 10, 17.5) + '<path d="M10 2v5M7.5 4.5h5" />',
  rowBelow: '<rect x="2.5" y="2.5" width="15" height="8.5" rx="1"/>' + L(10, 2.5, 10, 11) + '<path d="M10 13v5M7.5 15.5h5" />',
  colLeft: '<rect x="9" y="2.5" width="8.5" height="15" rx="1"/>' + L(9, 10, 17.5, 10) + '<path d="M2 10h5M4.5 7.5v5" />',
  colRight: '<rect x="2.5" y="2.5" width="8.5" height="15" rx="1"/>' + L(2.5, 10, 11, 10) + '<path d="M13 10h5M15.5 7.5v5" />',
  deleteRow: '<rect x="2.5" y="6.5" width="15" height="7" rx="1"/><path d="m7.5 8 5 4m0-4-5 4" stroke="#c42b1c"/>',
  deleteCol: '<rect x="6.5" y="2.5" width="7" height="15" rx="1"/><path d="m8 7.5 4 5m0-5-4 5" stroke="#c42b1c"/>',
  deleteTable: '<rect x="2.5" y="3.5" width="15" height="13" rx="1"/>' + L(2.5, 8, 17.5, 8) + '<path d="m7 10 6 5m0-5-6 5" stroke="#c42b1c"/>',
  merge: '<rect x="2.5" y="4.5" width="15" height="11" rx="1"/><path d="M5 10h3.5M6.5 8l2 2-2 2M15 10h-3.5M13.5 8l-2 2 2 2"/>',
  split: '<rect x="2.5" y="4.5" width="15" height="11" rx="1"/>' + L(10, 4.5, 10, 15.5) + '<path d="M8 10H4.5M6.5 8l-2 2 2 2M12 10h3.5M13.5 8l2 2-2 2"/>',
  bordersAll: '<rect x="3" y="3" width="14" height="14"/>' + L(3, 10, 17, 10) + L(10, 3, 10, 17),
  bordersNone: '<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.2 1.6" stroke-opacity=".6"/>' + '<path d="M3 10h14M10 3v14" stroke-dasharray="1.2 1.6" stroke-opacity=".6"/>',
  bordersOutside: '<rect x="3" y="3" width="14" height="14"/>' + '<path d="M3 10h14M10 3v14" stroke-dasharray="1.2 1.6" stroke-opacity=".6"/>',
  bordersInside: '<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.2 1.6" stroke-opacity=".6"/>' + L(3, 10, 17, 10) + L(10, 3, 10, 17),
  vTop: '<rect x="3" y="3" width="14" height="14" rx="1" stroke-opacity=".5"/>' + lines([6, 9], 6, 14),
  vMiddle: '<rect x="3" y="3" width="14" height="14" rx="1" stroke-opacity=".5"/>' + lines([8.5, 11.5], 6, 14),
  vBottom: '<rect x="3" y="3" width="14" height="14" rx="1" stroke-opacity=".5"/>' + lines([11, 14], 6, 14),
  distribute: '<rect x="2.5" y="4.5" width="15" height="11" rx="1"/>' + L(7.5, 4.5, 7.5, 15.5) + L(12.5, 4.5, 12.5, 15.5) + '<path d="M3.5 18h13M3.5 17v2M16.5 17v2" />',
  headerRow: '<rect x="2.5" y="3.5" width="15" height="13" rx="1"/><rect x="2.5" y="3.5" width="15" height="4.5" fill="currentColor" fill-opacity=".3"/>' + L(2.5, 12.3, 17.5, 12.3),

  // ---- Image ----
  imgSmall: '<rect x="6" y="6" width="8" height="8" rx="1"/>',
  imgMedium: '<rect x="4" y="4" width="12" height="12" rx="1"/>',
  imgLarge: '<rect x="2" y="2" width="16" height="16" rx="1"/>',
  trash: '<path d="M3.5 5.5h13M8 5.5V3.5h4v2M5 5.5l1 12h8l1-12"/>' + L(8.5, 8.5, 8.5, 14.5) + L(11.5, 8.5, 11.5, 14.5),
  alt: '<rect x="2" y="4" width="16" height="12" rx="1.5"/><text x="10" y="12.6" text-anchor="middle" font-size="6" font-weight="700" fill="currentColor" stroke="none" font-family="Arial">ALT</text>',

  // ---- File / quick access ----
  newDoc: '<path d="M5 2h7l4 4v12H5z"/><path d="M12 2v4h4M10.5 9v6M7.5 12h6"/>',
  open: '<path d="M2.5 5.5V15a1 1 0 0 0 1 1h12l2-7H6l-2 7"/><path d="M2.5 5.5V4a1 1 0 0 1 1-1h4l1.5 2h6a1 1 0 0 1 1 1v3"/>',
  save: '<path d="M4 2.5h9.5L17 6v10.5a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-13a1 1 0 0 1 1-1z"/><path d="M6 2.5v4.5h7V2.5M6 17.5v-5h8v5"/>',
  download: '<path d="M10 3v10M6 9l4 4 4-4M3 15v2h14v-2"/>',
  html: '<path d="M7 6 3 10l4 4M13 6l4 4-4 4M11.5 4l-3 12"/>',
  print: '<path d="M5.5 7.5V2.5h9v5"/><rect x="2.5" y="7.5" width="15" height="7" rx="1.5"/><rect x="5.5" y="12" width="9" height="6"/>',
  pdf: '<path d="M5 2h7l4 4v12H5z"/><path d="M12 2v4h4"/><text x="10.5" y="15" text-anchor="middle" font-size="5" font-weight="700" fill="currentColor" stroke="none" font-family="Arial">PDF</text>',
  undo: '<path d="M7 4 3 8l4 4"/><path d="M3 8h9a5 5 0 0 1 0 10H8"/>',
  redo: '<path d="m13 4 4 4-4 4"/><path d="M17 8H8a5 5 0 0 0 0 10h4"/>',
  close: '<path d="m5 5 10 10M15 5 5 15"/>',
  chevron: '<path d="m6 8 4 4 4-4"/>',
  check: '<path d="m4 10.5 4 4 8-9"/>',
  info: '<circle cx="10" cy="10" r="7.5"/><path d="M10 9v5M10 6.2v.1"/>',
  more: '<circle cx="5" cy="10" r="1.2" fill="currentColor"/><circle cx="10" cy="10" r="1.2" fill="currentColor"/><circle cx="15" cy="10" r="1.2" fill="currentColor"/>',
};
