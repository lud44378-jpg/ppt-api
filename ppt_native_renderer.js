#!/usr/bin/env node
'use strict';

/*
 * This renderer deliberately accepts a small teaching-storyboard schema rather
 * than arbitrary coordinates from an LLM.  It keeps every text box, shape and
 * image editable in PowerPoint while the layout system owns readability.
 */
const fs = require('fs');
const pptxgen = require('pptxgenjs');

const [inputPath, outputPath] = process.argv.slice(2);
if (!inputPath || !outputPath) throw new Error('Usage: node ppt_native_renderer.js input.json output.pptx');

const deck = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = '师助AI';
pptx.subject = deck.title || '教学演示文稿';
pptx.title = deck.title || '教学演示文稿';
pptx.company = '师助AI';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN'
};

const palettes = {
  safety:  { ink: '8A2B16', brand: 'D95F02', soft: 'FFF1E8', accent: 'F6A400', dark: '542013' },
  exam:    { ink: '15245D', brand: '3056B3', soft: 'EEF2FF', accent: 'E45050', dark: '101A46' },
  health:  { ink: '174D36', brand: '27865B', soft: 'EAF7EF', accent: 'E6A23C', dark: '113B29' },
  holiday: { ink: '74183B', brand: 'B93666', soft: 'FFF0F5', accent: 'F29A49', dark: '4A1026' },
  plan:    { ink: '44306A', brand: '6D50A2', soft: 'F3EEFF', accent: 'DE8D47', dark: '2E204B' },
  default: { ink: '16375E', brand: '207C9C', soft: 'EAF6F9', accent: 'EE9B45', dark: '0F2B4A' }
};
const c = palettes[deck.theme] || palettes.default;
const W = 13.333, H = 7.5;

function safeColor(value, fallback) {
  return /^[0-9A-Fa-f]{6}$/.test(String(value || '').replace('#', ''))
    ? String(value).replace('#', '').toUpperCase() : fallback;
}
function addText(slide, text, x, y, w, h, options = {}) {
  const opts = Object.assign({
    x, y, w, h, margin: 0, breakLine: false, fit: 'shrink',
    fontFace: 'Microsoft YaHei', color: c.ink, fontSize: 18,
    valign: 'mid', paraSpaceAfterPt: 5, transparency: 0
  }, options);
  slide.addText(String(text || ''), opts);
}
function addCard(slide, x, y, w, h, text, options = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: options.fill || c.soft },
    line: { color: options.line || c.soft, transparency: options.lineTransparency || 100 },
    radius: 0.12
  });
  addText(slide, text, x + 0.22, y + 0.16, w - 0.44, h - 0.32, {
    fontSize: options.fontSize || 16,
    color: options.color || c.ink,
    bold: Boolean(options.bold),
    valign: options.valign || 'mid',
    align: options.align || 'left'
  });
}
function addHeader(slide, title, page, total) {
  addText(slide, title, 0.72, 0.44, 10.8, 0.55, { fontSize: 29, bold: true, color: c.ink });
  addText(slide, String(page) + ' / ' + String(total), 11.82, 6.92, 0.7, 0.22, {
    fontSize: 9, color: c.brand, align: 'right', valign: 'mid'
  });
}
function addImage(slide, item, x, y, w, h) {
  if (!item || !item._image_path || !fs.existsSync(item._image_path)) return false;
  try {
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.12,
      fill: { color: 'FFFFFF' }, line: { color: c.soft, transparency: 0 } });
    slide.addImage({ path: item._image_path, x: x + 0.04, y: y + 0.04, w: w - 0.08, h: h - 0.08, sizing: { type: 'cover', x: x + 0.04, y: y + 0.04, w: w - 0.08, h: h - 0.08 } });
    return true;
  } catch (error) {
    console.error('[native-ppt] image skipped:', error.message);
    return false;
  }
}
function blocksOf(item, max = 6) {
  return (Array.isArray(item.blocks) ? item.blocks : []).map(v => String(v).trim()).filter(Boolean).slice(0, max);
}
function addBulletRows(slide, blocks, x, y, w, h) {
  const count = Math.max(blocks.length, 1);
  const rowH = Math.min(0.92, Math.max(0.58, (h - 0.14 * (count - 1)) / count));
  blocks.forEach((block, index) => {
    const rowY = y + index * (rowH + 0.14);
    slide.addShape(pptx.ShapeType.ellipse, { x, y: rowY + 0.22, w: 0.17, h: 0.17,
      fill: { color: c.accent }, line: { color: c.accent } });
    addText(slide, block, x + 0.35, rowY, w - 0.35, rowH, { fontSize: block.length > 72 ? 15 : 17, color: c.ink });
  });
}
function renderContent(slide, item, page, total) {
  const layout = item.layout || 'explain';
  const blocks = blocksOf(item);
  addHeader(slide, item.title || '教学要点', page, total);
  if (layout === 'story' || layout === 'scenario') {
    const hasImage = addImage(slide, item, 8.55, 1.48, 3.82, 4.75);
    addCard(slide, 0.86, 1.45, hasImage ? 7.18 : 11.6, 4.85, blocks.join('\n\n'), {
      fontSize: blocks.join('').length > 240 ? 16 : 19, fill: c.soft
    });
    if (layout === 'scenario') addText(slide, '请先独立思考，再与同伴交流。', 1.12, 6.42, 6.4, 0.3, { fontSize: 11, color: c.brand });
    return;
  }
  if (layout === 'cards' || layout === 'action') {
    const count = Math.min(Math.max(blocks.length, 1), 6);
    const cols = count <= 2 ? 2 : 3;
    const rows = Math.ceil(count / cols);
    const cw = (11.65 - (cols - 1) * 0.26) / cols;
    const ch = Math.min(2.05, (4.85 - (rows - 1) * 0.28) / rows);
    blocks.forEach((block, index) => {
      const col = index % cols, row = Math.floor(index / cols);
      const x = 0.86 + col * (cw + 0.26), y = 1.5 + row * (ch + 0.28);
      addCard(slide, x, y, cw, ch, block, { fontSize: block.length > 72 ? 14 : 16, fill: index % 2 ? 'FFFFFF' : c.soft, line: c.soft });
      slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.18, y: y + 0.16, w: 0.28, h: 0.28,
        fill: { color: c.accent }, line: { color: c.accent } });
    });
    return;
  }
  if (layout === 'steps' || layout === 'timeline') {
    const count = Math.min(Math.max(blocks.length, 1), 5);
    const gap = 0.25, cw = (11.62 - gap * (count - 1)) / count;
    blocks.slice(0, count).forEach((block, index) => {
      const x = 0.86 + index * (cw + gap);
      if (index < count - 1) slide.addShape(pptx.ShapeType.line, { x: x + cw, y: 3.46, w: gap, h: 0, line: { color: c.brand, width: 1.2, beginArrowType: 'none', endArrowType: 'triangle' } });
      slide.addShape(pptx.ShapeType.ellipse, { x: x + cw / 2 - 0.28, y: 1.63, w: 0.56, h: 0.56,
        fill: { color: c.brand }, line: { color: c.brand } });
      addText(slide, String(index + 1), x + cw / 2 - 0.2, 1.71, 0.4, 0.22, { fontSize: 14, bold: true, color: 'FFFFFF', align: 'center' });
      addCard(slide, x, 2.42, cw, 2.25, block, { fontSize: block.length > 58 ? 14 : 16, align: 'center', fill: c.soft });
    });
    return;
  }
  if (layout === 'compare') {
    const mid = Math.ceil(blocks.length / 2);
    const left = blocks.slice(0, mid), right = blocks.slice(mid);
    addCard(slide, 0.86, 1.48, 5.65, 4.85, left.join('\n\n'), { fontSize: 17, fill: c.soft });
    addCard(slide, 6.82, 1.48, 5.65, 4.85, right.join('\n\n'), { fontSize: 17, fill: 'FFFFFF', line: c.soft });
    return;
  }
  addCard(slide, 0.86, 1.45, 11.6, 4.95, '', { fill: c.soft });
  addBulletRows(slide, blocks, 1.22, 1.8, 10.85, 4.1);
}
function addNotes(slide, text) {
  if (text && typeof slide.addNotes === 'function') slide.addNotes(String(text));
}
function renderCover() {
  const slide = pptx.addSlide();
  slide.background = { color: c.dark };
  slide.addShape(pptx.ShapeType.arc, { x: 8.4, y: -1.0, w: 5.8, h: 5.8, adjustPoint: 0.2,
    line: { color: c.brand, transparency: 0, width: 3 }, rotate: 30 });
  addText(slide, deck.title || '教学演示文稿', 1.0, 2.05, 10.1, 1.45, { fontSize: 42, bold: true, color: 'FFFFFF' });
  if (deck.subtitle) addText(slide, deck.subtitle, 1.04, 3.85, 8.5, 0.58, { fontSize: 20, color: c.soft });
  addText(slide, '师助AI · 教学支持', 1.04, 6.52, 3.4, 0.26, { fontSize: 10, color: c.accent });
}
function renderSources(page, total) {
  const sources = Array.isArray(deck.sources) ? deck.sources.slice(0, 6) : [];
  if (!sources.length) return;
  const slide = pptx.addSlide();
  addHeader(slide, '资料来源与延伸阅读', page, total);
  sources.forEach((source, index) => {
    const domain = String(source.url || '').replace(/^https?:\/\//, '').split('/')[0];
    addCard(slide, 0.92, 1.48 + index * 0.74, 11.5, 0.55,
      `${source.title || domain}${domain ? '  ·  ' + domain : ''}`, { fontSize: 13, fill: index % 2 ? 'FFFFFF' : c.soft, line: c.soft });
  });
}
function renderEnd() {
  const slide = pptx.addSlide();
  slide.background = { color: c.soft };
  slide.addShape(pptx.ShapeType.ellipse, { x: 9.58, y: 1.22, w: 1.5, h: 1.5, fill: { color: c.accent, transparency: 8 }, line: { color: c.accent, transparency: 100 } });
  addText(slide, '把今天的理解，变成明天的行动', 1.1, 2.7, 11.1, 0.75, { fontSize: 35, bold: true, color: c.dark, align: 'center' });
}

const content = Array.isArray(deck.slides) ? deck.slides : [];
const sourceCount = Array.isArray(deck.sources) && deck.sources.length ? 1 : 0;
const total = content.length + 2 + sourceCount;
renderCover();
content.forEach((item, index) => {
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  renderContent(slide, item, index + 2, total);
  addNotes(slide, item.speaker_note);
});
if (sourceCount) renderSources(content.length + 2, total);
renderEnd();
pptx.writeFile({ fileName: outputPath });
