/* 確認画面に出すスライドのプレビュー。
   pptx.js と同じ layout.js の値を使って canvas に描くので、
   実際の仕上がりとずれない。 */

import { getTheme } from './theme.js';
import {
  SLIDE_W, SLIDE_H, M, CONTENT_W, CONTENT_TOP, CONTENT_H, FOOTER_Y,
  COVER_BAND_Y, COVER_TITLE_MID, COVER_SUB_PT,
  COVER_RULE_TOP_Y, COVER_RULE_BOT_Y, COVER_ACCENT_W, COVER_ACCENT_H, COVER_ACCENT_Y,
  HEAD_MID, RULE_Y, RULE_ACCENT_Y,
  RULE_ACCENT_W, RULE_ACCENT_H, HAIRLINE_H, TEXT_COL_W,
  DIVIDER_MID, DIVIDER_RULE_Y, DIVIDER_RULE_W, PAGE_NO_DY,
  gridCells, fitRect, visibleMedia, slideKind, mediaRegion,
  titleSize, headingSize, bodySize, bodyLines, paraSpacePt,
  BODY_LINE_SPACING, BULLET_INDENT,
} from './layout.js';

const FONT = '"Hiragino Sans", "Noto Sans JP", "Yu Gothic UI", Meiryo, sans-serif';

const imgCache = new Map();

function loadImage(src) {
  let p = imgCache.get(src);
  if (!p) {
    p = new Promise((resolve) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => resolve(null);
      i.src = src;
    });
    imgCache.set(src, p);
  }
  return p;
}

// canvas を実寸に合わせ、インチ→px（S）と pt→px（PT）の係数を返す
function setup(canvas) {
  const cssW = canvas.clientWidth || (canvas.parentElement && canvas.parentElement.clientWidth) || 320;
  const cssH = cssW * SLIDE_H / SLIDE_W;
  const dpr = Math.min(window.devicePixelRatio || 1, 2);

  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);
  canvas.style.height = cssH + 'px';

  const ctx = canvas.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);

  const S = cssW / SLIDE_W;
  return { ctx, S, PT: S / 72 };
}

/** 表紙を描く */
export function drawCover(canvas, deck) {
  const t = getTheme(deck.theme);
  const { ctx, S, PT } = setup(canvas);

  fill(ctx, t.bg, 0, 0, SLIDE_W * S, SLIDE_H * S);

  // タイトルを上下の細い罫線ではさむ
  fill(ctx, t.line, M * S, COVER_RULE_TOP_Y * S, CONTENT_W * S, Math.max(1, HAIRLINE_H * S));

  ctx.fillStyle = '#' + t.ink;
  ctx.font = `bold ${titleSize(deck.title) * PT}px ${FONT}`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText(deck.title || '', M * S, COVER_TITLE_MID * S);

  fill(ctx, t.line, M * S, COVER_RULE_BOT_Y * S, CONTENT_W * S, Math.max(1, HAIRLINE_H * S));
  fill(ctx, t.accent, M * S, COVER_ACCENT_Y * S, COVER_ACCENT_W * S, Math.max(2, COVER_ACCENT_H * S));

  fill(ctx, t.accent, 0, COVER_BAND_Y * S, SLIDE_W * S, (SLIDE_H - COVER_BAND_Y) * S);

  if (deck.subtitle) {
    ctx.fillStyle = '#' + t.bg;
    ctx.font = `${COVER_SUB_PT * PT}px ${FONT}`;
    ctx.fillText(deck.subtitle, M * S, (COVER_BAND_Y + (SLIDE_H - COVER_BAND_Y) / 2) * S);
  }
}

/** 本文スライドを描く（画像の読み込みを待つので非同期） */
export async function drawSlide(canvas, deck, slide, pageNo) {
  const t = getTheme(deck.theme);
  const { ctx, S, PT } = setup(canvas);

  const kind = slideKind(slide);
  const media = visibleMedia(slide);

  if (kind === 'divider') {
    fill(ctx, t.accent, 0, 0, SLIDE_W * S, SLIDE_H * S);
    ctx.fillStyle = '#' + t.bg;
    ctx.font = `bold ${titleSize(slide.heading) * PT}px ${FONT}`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(slide.heading || '', (SLIDE_W / 2) * S, DIVIDER_MID * S);
    fill(ctx, t.bg, ((SLIDE_W - DIVIDER_RULE_W) / 2) * S, DIVIDER_RULE_Y * S,
         DIVIDER_RULE_W * S, Math.max(1, HAIRLINE_H * S));
    pageNumber(ctx, S, PT, pageNo, t.bg);
    return;
  }

  fill(ctx, t.bg, 0, 0, SLIDE_W * S, SLIDE_H * S);

  ctx.fillStyle = '#' + t.accent;
  ctx.font = `bold ${headingSize(slide.heading) * PT}px ${FONT}`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText(slide.heading || '', M * S, HEAD_MID * S);

  fill(ctx, t.line, M * S, RULE_Y * S, CONTENT_W * S, Math.max(1, HAIRLINE_H * S));
  fill(ctx, t.accent, M * S, RULE_ACCENT_Y * S, RULE_ACCENT_W * S, Math.max(2, RULE_ACCENT_H * S));

  if (kind === 'split') {
    drawBody(ctx, slide.body, S, PT, TEXT_COL_W, t.ink);
    await drawMedia(ctx, S, media, mediaRegion(true));
  } else if (kind === 'media') {
    await drawMedia(ctx, S, media, mediaRegion(false));
  } else {
    drawBody(ctx, slide.body, S, PT, CONTENT_W, t.ink);
  }

  fill(ctx, t.line, M * S, FOOTER_Y * S, CONTENT_W * S, Math.max(1, HAIRLINE_H * S));
  pageNumber(ctx, S, PT, pageNo, t.muted);
}

/* ---------- 部品 ---------- */

function fill(ctx, color, x, y, w, h) {
  ctx.fillStyle = '#' + color;
  ctx.fillRect(x, y, w, h);
}

function pageNumber(ctx, S, PT, pageNo, color) {
  ctx.fillStyle = '#' + color;
  ctx.font = `${9 * PT}px ${FONT}`;
  ctx.textAlign = 'right';
  ctx.textBaseline = 'top';
  ctx.fillText(String(pageNo), (SLIDE_W - M) * S, (FOOTER_Y + PAGE_NO_DY + 0.04) * S);
}

// 日本語は単語区切りがないので1文字ずつ測って折り返す
function wrapText(ctx, text, maxW) {
  const lines = [];
  let line = '';
  for (const ch of text) {
    const test = line + ch;
    if (ctx.measureText(test).width > maxW && line) {
      lines.push(line);
      line = ch;
    } else {
      line = test;
    }
  }
  if (line) lines.push(line);
  return lines;
}

// pptx 側と同じ bodySize / 行送りを使うので、プレビューと仕上がりがそろう
function drawBody(ctx, body, S, PT, boxW, color) {
  const pt = bodySize(body, boxW, CONTENT_H);
  const fontPx = pt * PT;
  const x = M * S;
  const y = CONTENT_TOP * S;
  const w = boxW * S;
  const h = CONTENT_H * S;

  ctx.fillStyle = '#' + color;
  ctx.font = `${fontPx}px ${FONT}`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'top';

  const lh = fontPx * BODY_LINE_SPACING;
  const paraGap = paraSpacePt(pt) * PT;
  let cy = y;

  for (const line of bodyLines(body)) {
    if (!line.text) { cy += lh * 0.6; continue; }

    const indent = line.bullet ? fontPx * BULLET_INDENT : 0;
    const wrapped = wrapText(ctx, line.text, w - indent);

    for (let k = 0; k < wrapped.length; k++) {
      if (cy + lh > y + h) return;           // 入りきらない分は描かない
      if (line.bullet && k === 0) ctx.fillText('•', x, cy);
      ctx.fillText(wrapped[k], x + indent, cy);
      cy += lh;
    }
    cy += paraGap;
  }
}

async function drawMedia(ctx, S, media, region) {
  const cells = gridCells(region, media.length);

  for (let k = 0; k < media.length; k++) {
    const m = media[k];
    const box = fitRect(cells[k], m.aspect);
    const src = m.kind === 'video' ? (m.cover || null) : m.data;

    if (src) {
      const img = await loadImage(src);
      if (img) ctx.drawImage(img, box.x * S, box.y * S, box.w * S, box.h * S);
    } else {
      fill(ctx, 'DDDDE2', box.x * S, box.y * S, box.w * S, box.h * S);
    }
    if (m.kind === 'video') playBadge(ctx, box, S);
  }
}

function playBadge(ctx, box, S) {
  const cx = (box.x + box.w / 2) * S;
  const cy = (box.y + box.h / 2) * S;
  const r = Math.max(6, Math.min(box.w, box.h) * S * 0.16);

  ctx.fillStyle = 'rgba(0,0,0,.55)';
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = '#fff';
  ctx.beginPath();
  ctx.moveTo(cx - r * 0.3, cy - r * 0.45);
  ctx.lineTo(cx + r * 0.5, cy);
  ctx.lineTo(cx - r * 0.3, cy + r * 0.45);
  ctx.closePath();
  ctx.fill();
}
