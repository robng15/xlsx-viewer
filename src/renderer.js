import * as XLSX from 'xlsx';

// ─── Border Style Mapping ────────────────────────────────────────────────────
const BORDER_STYLES = {
  thin:             '1px solid',
  medium:           '2px solid',
  thick:            '3px solid',
  hair:             '1px solid',
  dashed:           '1px dashed',
  mediumDashed:     '2px dashed',
  dotted:           '1px dotted',
  double:           '3px double',
  dashDot:          '1px dashed',
  mediumDashDot:    '2px dashed',
  dashDotDot:       '1px dotted',
  mediumDashDotDot: '2px dotted',
  slantDashDot:     '1px dashed',
};

// ─── Color Helpers ────────────────────────────────────────────────────────────
/**
 * Convert a SheetJS color object { rgb, theme, indexed } to a CSS color string.
 * Returns null if the color cannot be resolved.
 */
function resolveColor(colorObj) {
  if (!colorObj) return null;

  // ARGB hex string (Excel stores alpha as first byte, e.g. "FF2563EB")
  if (colorObj.rgb) {
    const hex = colorObj.rgb.replace(/^FF/i, ''); // strip full-opacity alpha
    if (hex.length === 6) return `#${hex}`;
    if (hex.length === 8) {
      // has non-trivial alpha — convert to rgba
      const a = parseInt(colorObj.rgb.slice(0, 2), 16) / 255;
      return `rgba(${parseInt(hex.slice(0, 2), 16)},${parseInt(hex.slice(2, 4), 16)},${parseInt(hex.slice(4, 6), 16)},${a.toFixed(2)})`;
    }
  }

  // Indexed colors: Excel's legacy 56-color palette (partial mapping)
  if (colorObj.indexed !== undefined) {
    return INDEXED_COLORS[colorObj.indexed] ?? null;
  }

  // Theme colors require the workbook theme, which SheetJS CE doesn't expose.
  // We return null and let the browser default handle it.
  return null;
}

// Partial Excel indexed color palette (most common entries)
const INDEXED_COLORS = {
  0:  '#000000', 1:  '#FFFFFF', 2:  '#FF0000', 3:  '#00FF00',
  4:  '#0000FF', 5:  '#FFFF00', 6:  '#FF00FF', 7:  '#00FFFF',
  8:  '#000000', 9:  '#FFFFFF', 10: '#FF0000', 11: '#00FF00',
  12: '#0000FF', 13: '#FFFF00', 14: '#FF00FF', 15: '#00FFFF',
  16: '#800000', 17: '#008000', 18: '#000080', 19: '#808000',
  20: '#800080', 21: '#008080', 22: '#C0C0C0', 23: '#808080',
  24: '#9999FF', 25: '#993366', 26: '#FFFFCC', 27: '#CCFFFF',
  28: '#660066', 29: '#FF8080', 30: '#0066CC', 31: '#CCCCFF',
  32: '#000080', 33: '#FF00FF', 34: '#FFFF00', 35: '#00FFFF',
  36: '#800080', 37: '#800000', 38: '#008080', 39: '#0000FF',
  40: '#00CCFF', 41: '#CCFFFF', 42: '#CCFFCC', 43: '#FFFF99',
  44: '#99CCFF', 45: '#FF99CC', 46: '#CC99FF', 47: '#FFCC99',
  48: '#3366FF', 49: '#33CCCC', 50: '#99CC00', 51: '#FFCC00',
  52: '#FF9900', 53: '#FF6600', 54: '#666699', 55: '#969696',
  56: '#003366', 57: '#339966', 58: '#003300', 59: '#333300',
  60: '#993300', 61: '#993366', 62: '#333399', 63: '#333333',
  64: '#000000', // system foreground
};

// ─── Style → CSS ──────────────────────────────────────────────────────────────
/**
 * Convert a SheetJS cell style object to an inline CSS string.
 * Returns empty string for null/undefined styles.
 */
export function styleToCSS(style) {
  if (!style) return '';
  const parts = [];

  // Fill / background
  if (style.fill) {
    const { patternType, fgColor, bgColor } = style.fill;
    if (patternType && patternType !== 'none') {
      const bg = resolveColor(fgColor) ?? resolveColor(bgColor);
      if (bg) parts.push(`background-color:${bg}`);
    }
  }

  // Font
  if (style.font) {
    const f = style.font;
    if (f.bold)      parts.push('font-weight:bold');
    if (f.italic)    parts.push('font-style:italic');
    const deco = [];
    if (f.underline) deco.push('underline');
    if (f.strike)    deco.push('line-through');
    if (deco.length) parts.push(`text-decoration:${deco.join(' ')}`);
    const fc = resolveColor(f.color);
    if (fc) parts.push(`color:${fc}`);
    if (f.sz)   parts.push(`font-size:${f.sz}pt`);
    if (f.name) parts.push(`font-family:"${f.name}",sans-serif`);
  }

  // Alignment
  if (style.alignment) {
    const a = style.alignment;
    if (a.horizontal) parts.push(`text-align:${a.horizontal}`);
    if (a.vertical) {
      const vmap = { top: 'top', center: 'middle', bottom: 'bottom' };
      parts.push(`vertical-align:${vmap[a.vertical] ?? a.vertical}`);
    }
    if (a.wrapText) parts.push('white-space:pre-wrap;word-break:break-word');
    if (a.indent)   parts.push(`padding-left:${a.indent * 8}px`);
  }

  // Borders
  if (style.border) {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const b = style.border[side];
      if (b?.style && b.style !== 'none') {
        const bDecl  = BORDER_STYLES[b.style] ?? '1px solid';
        const bColor = resolveColor(b.color) ?? '#000';
        parts.push(`border-${side}:${bDecl} ${bColor}`);
      }
    }
  }

  return parts.join(';');
}

// ─── Cell Display Value ───────────────────────────────────────────────────────
/**
 * Return the display string for a cell.
 * Prefers the pre-formatted `w` value (which Excel stored), then falls back
 * to the raw `v` value. Boolean cells are shown as TRUE / FALSE.
 * Formula cells show the cached result — no re-evaluation occurs.
 */
function cellDisplay(cell) {
  if (!cell) return '';
  if (cell.w !== undefined) return cell.w;
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.t === 'e') return cell.w ?? String(cell.v ?? '');
  if (cell.v !== undefined) return String(cell.v);
  return '';
}

// ─── HTML Escape ──────────────────────────────────────────────────────────────
function esc(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── Column Width ─────────────────────────────────────────────────────────────
function colWidthPx(colInfo) {
  if (!colInfo || colInfo.hidden) return null;
  if (colInfo.wpx)  return colInfo.wpx;
  if (colInfo.wch)  return Math.round(colInfo.wch * 7 + 5);
  if (colInfo.width) return Math.round(colInfo.width * 7 + 5);
  return null;
}

// ─── Main Render ─────────────────────────────────────────────────────────────
/**
 * Render a SheetJS worksheet object to an HTML string.
 * @param {object} ws  - SheetJS worksheet
 * @returns {string}   - HTML string for the table
 */
export function renderWorksheet(ws) {
  if (!ws || !ws['!ref']) {
    return '<p class="empty-sheet">This sheet is empty.</p>';
  }

  const range  = XLSX.utils.decode_range(ws['!ref']);
  const merges = ws['!merges'] ?? [];
  const cols   = ws['!cols']   ?? [];
  const rows   = ws['!rows']   ?? [];

  // ── Build merge lookup ─────────────────────────────────────────────────────
  // mergeSpan: "r,c" → { rowspan, colspan }
  // skipCell:  Set of "r,c" strings that are covered by a merge (not rendered)
  const mergeSpan = new Map();
  const skipCell  = new Set();

  for (const m of merges) {
    const rs = m.e.r - m.s.r + 1;
    const cs = m.e.c - m.s.c + 1;
    mergeSpan.set(`${m.s.r},${m.s.c}`, { rowspan: rs, colspan: cs });
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r !== m.s.r || c !== m.s.c) skipCell.add(`${r},${c}`);
      }
    }
  }

  // ── Build HTML ─────────────────────────────────────────────────────────────
  const parts = ['<table class="sheet-table">'];

  // colgroup — controls column widths
  parts.push('<colgroup>');
  parts.push('<col style="width:48px;min-width:48px">'); // row-header column
  for (let c = range.s.c; c <= range.e.c; c++) {
    const w = colWidthPx(cols[c]);
    if (w) {
      parts.push(`<col style="width:${w}px;min-width:${w}px">`);
    } else {
      parts.push('<col style="min-width:80px">');
    }
  }
  parts.push('</colgroup>');

  // thead — column letter headers (A, B, C …)
  parts.push('<thead><tr>');
  parts.push('<th class="corner-cell" scope="col"></th>');
  for (let c = range.s.c; c <= range.e.c; c++) {
    parts.push(`<th class="col-header" scope="col">${XLSX.utils.encode_col(c)}</th>`);
  }
  parts.push('</tr></thead>');

  // tbody — data rows
  parts.push('<tbody>');
  for (let r = range.s.r; r <= range.e.r; r++) {
    const rowInfo = rows[r];
    const rowStyle = rowInfo?.hpx ? ` style="height:${rowInfo.hpx}px"` : '';
    parts.push(`<tr${rowStyle}>`);

    // Row number header
    parts.push(`<td class="row-header" scope="row">${r + 1}</td>`);

    for (let c = range.s.c; c <= range.e.c; c++) {
      const key = `${r},${c}`;
      if (skipCell.has(key)) continue;

      const addr   = XLSX.utils.encode_cell({ r, c });
      const cell   = ws[addr];
      const merge  = mergeSpan.get(key);

      // Build td attributes
      let attrs = ' class="data-cell"';
      if (merge) {
        if (merge.rowspan > 1) attrs += ` rowspan="${merge.rowspan}"`;
        if (merge.colspan > 1) attrs += ` colspan="${merge.colspan}"`;
      }

      // Inline style: start with base padding, then add cell style
      let inlineStyle = 'padding:2px 4px';
      if (cell?.s) {
        const extra = styleToCSS(cell.s);
        if (extra) inlineStyle += ';' + extra;
      }
      attrs += ` style="${inlineStyle}"`;

      const display = cell ? esc(cellDisplay(cell)) : '';
      parts.push(`<td${attrs}>${display}</td>`);
    }

    parts.push('</tr>');
  }
  parts.push('</tbody></table>');

  return parts.join('');
}
