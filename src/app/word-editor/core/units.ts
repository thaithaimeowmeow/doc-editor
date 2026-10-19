/** Mọi kích thước bên trong engine dùng px CSS (96 dpi). */
export const PX_PER_MM = 96 / 25.4;
export const PX_PER_PT = 96 / 72;

export const mm = (v: number) => v * PX_PER_MM;
export const cm = (v: number) => v * 10 * PX_PER_MM;
export const toCm = (px: number) => Math.round((px / PX_PER_MM / 10) * 100) / 100;

/** 1 px = 15 twip (1440 twip/inch). */
export const pxToTwip = (px: number) => Math.round(px * 15);
export const twipToPx = (twip: number) => twip / 15;
/** EMU dùng cho ảnh trong OOXML: 914400 EMU/inch. */
export const pxToEmu = (px: number) => Math.round(px * 9525);

export interface PageSetup {
  width: number;
  height: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  /** Khoảng cách từ mép trên trang tới header. */
  header: number;
  /** Khoảng cách từ mép dưới trang tới footer. */
  footer: number;
}

export const PAPER_SIZES: Record<string, { label: string; w: number; h: number }> = {
  A4: { label: 'A4 (21 × 29,7 cm)', w: mm(210), h: mm(297) },
  A3: { label: 'A3 (29,7 × 42 cm)', w: mm(297), h: mm(420) },
  A5: { label: 'A5 (14,8 × 21 cm)', w: mm(148), h: mm(210) },
  Letter: { label: 'Letter (21,59 × 27,94 cm)', w: mm(215.9), h: mm(279.4) },
  Legal: { label: 'Legal (21,59 × 35,56 cm)', w: mm(215.9), h: mm(355.6) },
};

export const MARGIN_PRESETS: Record<string, { label: string; t: number; r: number; b: number; l: number }> = {
  normal: { label: 'Bình thường (2,54 cm)', t: 2.54, r: 2.54, b: 2.54, l: 2.54 },
  vn: { label: 'Văn bản hành chính (T2 D2 T3 P2 cm)', t: 2, r: 2, b: 2, l: 3 },
  narrow: { label: 'Hẹp (1,27 cm)', t: 1.27, r: 1.27, b: 1.27, l: 1.27 },
  moderate: { label: 'Vừa (2,54 / 1,91 cm)', t: 2.54, r: 1.91, b: 2.54, l: 1.91 },
  wide: { label: 'Rộng (2,54 / 5,08 cm)', t: 2.54, r: 5.08, b: 2.54, l: 5.08 },
};

export const DEFAULT_PAGE: PageSetup = {
  width: mm(210),
  height: mm(297),
  marginTop: cm(2.54),
  marginRight: cm(2.54),
  marginBottom: cm(2.54),
  marginLeft: cm(2.54),
  header: cm(1.25),
  footer: cm(1.25),
};

/** Đổi một giá trị CSS (px, pt, cm, mm, in, em) sang px. `em` tính theo `fontPx`. */
export function cssToPx(value: string | null | undefined, fontPx = 16): number | null {
  if (!value) return null;
  const m = /^(-?[\d.]+)(px|pt|cm|mm|in|em|rem|pc)?$/.exec(value.trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case 'pt':
      return n * PX_PER_PT;
    case 'cm':
      return cm(n);
    case 'mm':
      return mm(n);
    case 'in':
      return n * 96;
    case 'pc':
      return n * 16;
    case 'em':
    case 'rem':
      return n * fontPx;
    default:
      return n;
  }
}

/** rgb()/rgba()/#rgb/#rrggbb → "RRGGBB" (không có #), hoặc null nếu trong suốt/không đọc được. */
export function cssColorToHex(value: string | null | undefined): string | null {
  if (!value) return null;
  const v = value.trim().toLowerCase();
  if (v === 'transparent' || v === 'inherit' || v === 'initial' || v === 'currentcolor') return null;
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return m[1].split('').map(c => c + c).join('').toUpperCase();
  m = /^#([0-9a-f]{6})/.exec(v);
  if (m) return m[1].toUpperCase();
  m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?\s*\)$/.exec(v);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
    return [m[1], m[2], m[3]].map(n => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  if (typeof document !== 'undefined' && /^[a-z]+$/.test(v)) {
    const probe = document.createElement('span');
    probe.style.color = v;
    if (!probe.style.color) return null;
    document.body.append(probe);
    const rgb = getComputedStyle(probe).color;
    probe.remove();
    return cssColorToHex(rgb);
  }
  return null;
}

const naturalCache = new Map<string, number>();

/**
 * Chiều cao một dòng đơn của phông (tỉ lệ so với cỡ chữ, vd Times New Roman ≈ 1,149).
 * "Multiple" của Word (1,15; 1,5…) nhân với số này, còn CSS `line-height: 1.15` chỉ nhân với cỡ chữ,
 * nên tài liệu Word giãn dòng 1,15 hiển thị thấp hơn bản gốc ~15% nếu bỏ qua hệ số này.
 */
export function naturalLineFactor(fontFamily: string | null | undefined): number {
  const family = (fontFamily ?? '').trim() || "'Times New Roman'";
  let k = naturalCache.get(family);
  if (k === undefined) {
    const probe = document.createElement('span');
    probe.textContent = 'Hg';
    probe.style.cssText = `position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap;font-size:100px;line-height:normal;font-family:${family}`;
    document.body.append(probe);
    k = probe.getBoundingClientRect().height / 100;
    probe.remove();
    if (!(k > 0.8 && k < 2)) k = 1.149;
    naturalCache.set(family, k);
  }
  return k;
}
