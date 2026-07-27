/* .pptx の組み立て。
   PptxGenJS は vendor/pptxgen.bundle.js がグローバルに読み込む。

   deck = { title, subtitle, theme, slides: [{ heading, body, media: [] }] } */

import { getTheme } from './theme.js';
import {
  SLIDE_W, SLIDE_H, M, CONTENT_W, CONTENT_TOP, CONTENT_H, FOOTER_Y,
  COVER_BAND_Y, COVER_TITLE_Y, COVER_TITLE_H, COVER_SUB_PT,
  COVER_RULE_TOP_Y, COVER_RULE_BOT_Y, COVER_ACCENT_W, COVER_ACCENT_H, COVER_ACCENT_Y,
  HEAD_MID, RULE_Y, RULE_ACCENT_Y, RULE_ACCENT_W, RULE_ACCENT_H,
  HAIRLINE_H, TEXT_COL_W, DIVIDER_MID, DIVIDER_RULE_Y, DIVIDER_RULE_W, PAGE_NO_DY,
  gridCells, fitRect, visibleMedia, slideKind, mediaRegion,
  titleSize, headingSize, bodySize, bodyLines, paraSpacePt,
  BODY_LINE_SPACING,
} from './layout.js';

/* Noto Sans JP を指定する。web で広く配布されているフォントなので、
   pptx を読み込むアプリ側で置き換えが起きにくい。 */
const FONT = 'Noto Sans JP';

/* テキストボックスの内側余白。既定値はアプリによって違い、そのぶん
   折り返し位置がずれる。0 に固定して、指定した幅をそのまま使わせる。 */
const INSET = { margin: 0 };

export async function buildPptx(deck, outputType = 'blob') {
  const Ctor = globalThis.PptxGenJS;
  if (!Ctor) throw new Error('PptxGenJS が読み込まれていません');

  const t = getTheme(deck.theme);
  const pptx = new Ctor();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = deck.title || 'スライド';

  addCoverSlide(pptx, t, deck);
  (deck.slides || []).forEach((slide, idx) => addContentSlide(pptx, t, slide, idx + 2));

  return pptx.write({ outputType });
}

export function safeFileName(name) {
  const cleaned = String(name || '').replace(/[\\/:*?"<>|\r\n]/g, '_').trim().slice(0, 60);
  return cleaned || 'スライド';
}

/* ---------- 表紙 ---------- */

function addCoverSlide(pptx, t, deck) {
  const s = pptx.addSlide();
  s.background = { color: t.bg };

  const titlePt = titleSize(deck.title);

  // タイトルを上下の細い罫線ではさむ
  hairline(pptx, s, M, COVER_RULE_TOP_Y, CONTENT_W, t.line);

  s.addText(deck.title || '', {
    ...INSET,
    x: M, y: COVER_TITLE_Y, w: CONTENT_W, h: COVER_TITLE_H,
    fontFace: FONT, fontSize: titlePt, bold: true,
    color: t.ink, align: 'left', valign: 'middle', fit: 'shrink',
    lineSpacing: Math.round(titlePt * 1.15),
  });

  hairline(pptx, s, M, COVER_RULE_BOT_Y, CONTENT_W, t.line);
  s.addShape(pptx.ShapeType.rect, {
    x: M, y: COVER_ACCENT_Y, w: COVER_ACCENT_W, h: COVER_ACCENT_H, fill: { color: t.accent },
  });

  // 下端いっぱいの色帯。テーマの色がいちばん大きく出るところ
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: COVER_BAND_Y, w: SLIDE_W, h: SLIDE_H - COVER_BAND_Y, fill: { color: t.accent },
  });

  if (deck.subtitle) {
    s.addText(deck.subtitle, {
      ...INSET,
      x: M, y: COVER_BAND_Y, w: CONTENT_W, h: SLIDE_H - COVER_BAND_Y,
      fontFace: FONT, fontSize: COVER_SUB_PT, color: t.bg,
      align: 'left', valign: 'middle', fit: 'shrink',
    });
  }
}

/* ---------- 本文 ---------- */

function addContentSlide(pptx, t, slide, pageNo) {
  const s = pptx.addSlide();
  s.background = { color: t.bg };

  const kind = slideKind(slide);
  const media = visibleMedia(slide);

  // 見出しだけ → 章の区切り（中扉）。背景をテーマ色で塗って抜き文字にする
  if (kind === 'divider') {
    s.background = { color: t.accent };
    s.addText(slide.heading || '', {
      ...INSET,
      x: M, y: DIVIDER_MID - 0.75, w: CONTENT_W, h: 1.5,
      fontFace: FONT, fontSize: titleSize(slide.heading), bold: true,
      color: t.bg, align: 'center', valign: 'middle', fit: 'shrink',
    });
    hairline(pptx, s, (SLIDE_W - DIVIDER_RULE_W) / 2, DIVIDER_RULE_Y, DIVIDER_RULE_W, t.bg);
    addPageNo(s, pageNo, t.bg);
    return;
  }

  s.addText(slide.heading || '', {
    ...INSET,
    x: M, y: HEAD_MID - 0.33, w: CONTENT_W, h: 0.66,
    fontFace: FONT, fontSize: headingSize(slide.heading), bold: true,
    color: t.accent, align: 'left', valign: 'middle', fit: 'shrink',
  });
  // 細い線を全幅に引き、左端だけ太い色付きの線を重ねる
  hairline(pptx, s, M, RULE_Y, CONTENT_W, t.line);
  s.addShape(pptx.ShapeType.rect, {
    x: M, y: RULE_ACCENT_Y, w: RULE_ACCENT_W, h: RULE_ACCENT_H, fill: { color: t.accent },
  });

  if (kind === 'split') {
    addBody(s, t, slide.body, TEXT_COL_W);
    layoutMedia(pptx, s, media, mediaRegion(true));
  } else if (kind === 'media') {
    layoutMedia(pptx, s, media, mediaRegion(false));
  } else {
    addBody(s, t, slide.body, CONTENT_W);
  }

  hairline(pptx, s, M, FOOTER_Y, CONTENT_W, t.line);
  addPageNo(s, pageNo, t.muted);
}

// 本文。枠に収まる大きさをあらかじめ計算しているので自動縮小に頼らない
function addBody(s, t, body, boxW) {
  const pt = bodySize(body, boxW, CONTENT_H);
  s.addText(bodyParagraphs(body, pt), {
    ...INSET,
    x: M, y: CONTENT_TOP, w: boxW, h: CONTENT_H,
    fontFace: FONT, fontSize: pt, color: t.ink,
    valign: 'top', fit: 'shrink',
    lineSpacing: Math.round(pt * BODY_LINE_SPACING),
  });
}

function hairline(pptx, s, x, y, w, color) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h: HAIRLINE_H, fill: { color } });
}

function addPageNo(s, pageNo, color) {
  s.addText(String(pageNo), {
    ...INSET,
    x: SLIDE_W - M - 0.5, y: FOOTER_Y + PAGE_NO_DY, w: 0.5, h: 0.3,
    fontFace: FONT, fontSize: 9, color, align: 'right',
  });
}

function layoutMedia(pptx, s, media, region) {
  const cells = gridCells(region, media.length);

  media.forEach((m, k) => {
    const box = fitRect(cells[k], m.aspect);

    if (m.kind === 'image' || !m.kind) {
      // pptxgenjs の sizing:contain は data URL だと寸法を計算できず、
      // 指定した箱にそのまま引き伸ばされてしまう。自分で矩形を出して渡す
      s.addImage({ data: m.data, ...box });
      return;
    }

    if (m.kind === 'video') {
      const opts = { type: 'video', data: m.data, extn: videoExtension(m), ...box };
      if (m.cover) opts.cover = m.cover;
      s.addMedia(opts);
    }
  });
}

// 動画機能は現在無効だが、将来の再有効化に備えて埋め込み処理を残す。
function videoExtension(media) {
  const fromName = String(media.name || '').match(/\.([a-z0-9]{2,5})$/i);
  if (fromName) return fromName[1].toLowerCase();
  if (media.extn) return String(media.extn).toLowerCase();

  const subtype = String(media.data || '').match(/^data:video\/([^;,]+)/i)?.[1]?.toLowerCase();
  if (subtype === 'quicktime') return 'mov';
  if (subtype === 'x-m4v') return 'm4v';
  return subtype || 'mp4';
}

// 改行を段落に。箇条書きの行には中黒を付ける
function bodyParagraphs(body, sizePt) {
  const lines = bodyLines(body);
  const space = paraSpacePt(sizePt);
  return lines.map((line, i) => ({
    text: line.text,
    options: {
      bullet: line.bullet ? { characterCode: '2022' } : false,
      breakLine: i < lines.length - 1,
      paraSpaceAfter: space,
    },
  }));
}
