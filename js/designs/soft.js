/* デザイン「やわらか」。

   角丸の面に中身を載せて組む。罫線は引かず、境目は背景にわずかに
   色を混ぜた面で示す。見出しは本文色にして、色は面に回す。
   社内向けの提案資料や、やさしい印象を出したい場面に向く。

   「標準」が線で組み、「くっきり」が色面で組むのに対して、
   こちらは淡い面の重なりで組む。3種の骨格はここで分かれる。 */

import { cardSurface, fillOf, onFillOf } from '../theme.js';
import { SLIDE_W, visibleMedia, slideKind, bodyStyle, fitFontSize } from '../layout.js';
import { drawBody, drawPlain, drawMedia } from './common.js';

export const key = 'soft';
export const label = 'やわらか';

const M = 0.85;
const CONTENT_W = SLIDE_W - M * 2;

const PANEL_R = 0.3;              // 大きな面の角丸

const MARK_W = 0.5;               // 表紙と見出しに添える小さな丸帯
const MARK_H = 0.13;
const MARK_R = 0.065;

const HEAD_Y = 0.34;
const HEAD_H = 0.72;
const HEAD_MAX_PT = 26;

/* 中身を載せる大きな角丸の面。この面が骨格そのもの。
   カード組みのときはカードが面の役目を果たすので、二重にはしない。 */
const PANEL_Y = 1.18;
const PANEL_BOTTOM = 5.06;
const PANEL_H = PANEL_BOTTOM - PANEL_Y;
const PANEL_PAD = 0.42;

const CONTENT_TOP = PANEL_Y + PANEL_PAD;
const CONTENT_BOTTOM = PANEL_BOTTOM - PANEL_PAD;
const CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP;

// カード組みのときは面を敷かないので、上下いっぱいまで使える
const OPEN_TOP = 1.36;
const OPEN_BOTTOM = 5.02;
const OPEN_H = OPEN_BOTTOM - OPEN_TOP;

const TEXT_COL_W = 3.85;
const MEDIA_GAP = 0.5;

const COVER_TITLE_Y = 1.15;
const COVER_TITLE_H = 1.85;
const COVER_TITLE_MAX_PT = 38;
const COVER_PANEL_Y = 3.35;
const COVER_PANEL_H = 1.72;
const COVER_SUB_PT = 15;

const DIVIDER_PANEL_Y = 1.05;
const DIVIDER_PANEL_H = 3.5;
const DIVIDER_HEAD_MAX_PT = 34;

const PAGE_NO_PT = 9;
const PAGE_NO_Y = 5.18;

/* 画像だけのスライドの領域。面には載せず、本文より上下に広く取る。
   ページ番号には重ねないので、そのぶんは残す。 */
const MEDIA_ONLY = { x: M, y: 1.22, w: CONTENT_W, h: (PAGE_NO_Y - 0.08) - 1.22 };

// 面の上に置くとき（リード文・数値組み・ベタ組み）
const ON_PANEL = {
  region: { x: M + PANEL_PAD, y: CONTENT_TOP, w: CONTENT_W - PANEL_PAD * 2, h: CONTENT_H },
  lead: 'plain',
  leadMaxPt: 28,
  stat: 'plain',
};

// カード組みのとき。カード自体が面になるので下地は敷かない
const AS_CARDS = {
  region: { x: M, y: OPEN_TOP, w: CONTENT_W, h: OPEN_H },
  card: 'fill',
  cardRadius: 0.18,
  cardNumber: 'badge',
};

export function cover(surface, t, deck) {
  surface.background(t.bg);

  surface.roundRect(M, 0.72, MARK_W, MARK_H, MARK_R, fillOf(t));

  surface.text(deck.title, {
    x: M, y: COVER_TITLE_Y, w: CONTENT_W, h: COVER_TITLE_H,
    pt: fitFontSize(deck.title, CONTENT_W, COVER_TITLE_MAX_PT, 18),
    bold: true, color: t.ink, align: 'left', valign: 'middle', lineSpacing: 1.2,
  });

  // 下端に浮かせた角丸の面。帯を切り落とさないことで柔らかく見せる
  surface.roundRect(M, COVER_PANEL_Y, CONTENT_W, COVER_PANEL_H, PANEL_R, fillOf(t));

  surface.text(deck.subtitle, {
    x: M + 0.42, y: COVER_PANEL_Y, w: CONTENT_W - 0.84, h: COVER_PANEL_H,
    pt: COVER_SUB_PT, color: onFillOf(t), align: 'left', valign: 'middle',
  });
}

export async function slide(surface, t, slide, pageNo) {
  const kind = slideKind(slide);
  const media = visibleMedia(slide);

  surface.background(t.bg);

  if (kind === 'divider') {
    surface.roundRect(M, DIVIDER_PANEL_Y, CONTENT_W, DIVIDER_PANEL_H, PANEL_R, fillOf(t));
    surface.text(slide.heading, {
      x: M + 0.6, y: DIVIDER_PANEL_Y, w: CONTENT_W - 1.2, h: DIVIDER_PANEL_H,
      pt: fitFontSize(slide.heading, CONTENT_W - 1.2, DIVIDER_HEAD_MAX_PT, 18),
      bold: true, color: onFillOf(t), align: 'center', valign: 'middle', lineSpacing: 1.2,
    });
    pageNumber(surface, t, pageNo);
    return;
  }

  surface.roundRect(M, HEAD_Y + 0.05, MARK_W, MARK_H, MARK_R, fillOf(t));

  surface.text(slide.heading, {
    x: M + MARK_W + 0.22, y: HEAD_Y, w: CONTENT_W - MARK_W - 0.22, h: HEAD_H,
    pt: fitFontSize(slide.heading, CONTENT_W - MARK_W - 0.22, HEAD_MAX_PT, 14),
    bold: true, color: t.ink, align: 'left', valign: 'middle',
  });

  if (kind === 'split') {
    panel(surface, t);
    drawPlain(surface, t, slide.body, {
      x: M + PANEL_PAD, y: CONTENT_TOP, w: TEXT_COL_W, h: CONTENT_H,
    });
    const x = M + PANEL_PAD + TEXT_COL_W + MEDIA_GAP;
    await drawMedia(surface, media, {
      x, y: CONTENT_TOP, w: SLIDE_W - M - PANEL_PAD - x, h: CONTENT_H,
    });
  } else if (kind === 'media') {
    // 写真だけのときは面に載せず、そのまま大きく置く
    await drawMedia(surface, media, MEDIA_ONLY);
  } else if (bodyStyle(slide) === 'cards') {
    drawBody(surface, t, slide, AS_CARDS);
  } else {
    panel(surface, t);
    drawBody(surface, t, slide, ON_PANEL);
  }

  pageNumber(surface, t, pageNo);
}

// 中身を載せる淡い面。背景との差はわずかで、面があることだけが分かる濃さ
function panel(surface, t) {
  surface.roundRect(M, PANEL_Y, CONTENT_W, PANEL_H, PANEL_R, cardSurface(t));
}

function pageNumber(surface, t, pageNo) {
  surface.text(String(pageNo), {
    x: SLIDE_W - M - 0.5, y: PAGE_NO_Y, w: 0.5, h: 0.28,
    pt: PAGE_NO_PT, color: t.muted, align: 'right', valign: 'middle',
  });
}
