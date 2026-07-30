/* デザイン「くっきり」。

   罫線を使わず、テーマ色のベタ塗りだけで組む。
   本文スライドは上端に色帯を渡して見出しを抜き文字にし、
   カードもテーマ色で塗りつぶす。中扉は全面をテーマ色にする。
   遠くから見る発表や、白地が続くと眠くなる場面に向く。

   「標準」が線で組み、「やわらか」が角丸の面に載せるのに対して、
   こちらは色面そのもので組む。3種の骨格はここで分かれる。 */

import { mix, fillOf, onFillOf } from '../theme.js';
import { SLIDE_W, visibleMedia, slideKind, mediaFullRegion, fitFontSize } from '../layout.js';
import { drawBody, drawPlain, drawMedia } from './common.js';

export const key = 'bold';
export const label = 'くっきり';

const M = 0.8;
const CONTENT_W = SLIDE_W - M * 2;

const BAND_H = 1.15;              // 上端の色帯。見出しはこの中に抜き文字で置く
const HEAD_Y = 0.2;
const HEAD_H = 0.75;
const HEAD_W = CONTENT_W - 1.0;   // 右端はページ番号のために空ける
const HEAD_MAX_PT = 30;

const CONTENT_TOP = 1.52;
const CONTENT_BOTTOM = 5.18;
const CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP;

const TEXT_COL_W = 3.7;
const MEDIA_GAP = 0.5;

const COVER_TITLE_Y = 1.28;
const COVER_TITLE_H = 2.1;
const COVER_TITLE_MAX_PT = 56;
const COVER_RULE_Y = 3.6;
const COVER_RULE_W = 1.5;
const COVER_RULE_H = 0.07;
const COVER_SUB_Y = 3.95;
const COVER_SUB_PT = 15;

const DIVIDER_MARK_Y = 1.18;
const DIVIDER_MARK_W = 1.2;
const DIVIDER_HEAD_Y = 1.62;
const DIVIDER_HEAD_H = 2.4;
const DIVIDER_HEAD_MAX_PT = 44;

const PAGE_NO_PT = 12;
const MEDIA_FULL_TOP = BAND_H + 0.16;   // 色帯のすぐ下から下端まで

const BODY_OPT = {
  region: { x: M, y: CONTENT_TOP, w: CONTENT_W, h: CONTENT_H },
  lead: 'rule',
  leadMaxPt: 34,
  stat: 'plain',
  statRule: 'top',
  card: 'solid',
};

export function cover(surface, t, deck) {
  const f = fillOf(t);
  const on = onFillOf(t);

  // 全面をテーマ色で塗り、その上に載る文字は onFill で抜く
  surface.background(f);

  surface.text(deck.title, {
    x: M, y: COVER_TITLE_Y, w: CONTENT_W, h: COVER_TITLE_H,
    pt: fitFontSize(deck.title, CONTENT_W, COVER_TITLE_MAX_PT, 22),
    bold: true, color: on, align: 'left', valign: 'middle', lineSpacing: 1.08,
  });

  surface.rect(M, COVER_RULE_Y, COVER_RULE_W, COVER_RULE_H, on);

  surface.text(deck.subtitle, {
    x: M, y: COVER_SUB_Y, w: CONTENT_W, h: 0.7,
    // 抜き文字をわずかに沈ませて、タイトルとの主従をはっきりさせる
    pt: COVER_SUB_PT, color: mix(on, f, 0.78), align: 'left', valign: 'top',
  });
}

export async function slide(surface, t, slide, pageNo) {
  const kind = slideKind(slide);
  const media = visibleMedia(slide);
  const f = fillOf(t);
  const on = onFillOf(t);

  // 見出しだけ → 中扉。全面をテーマ色で塗る
  if (kind === 'divider') {
    surface.background(f);
    surface.rect((SLIDE_W - DIVIDER_MARK_W) / 2, DIVIDER_MARK_Y, DIVIDER_MARK_W, COVER_RULE_H, on);
    surface.text(slide.heading, {
      x: M + 0.5, y: DIVIDER_HEAD_Y, w: CONTENT_W - 1.0, h: DIVIDER_HEAD_H,
      pt: fitFontSize(slide.heading, CONTENT_W - 1.0, DIVIDER_HEAD_MAX_PT, 18),
      bold: true, color: on, align: 'center', valign: 'middle', lineSpacing: 1.15,
    });
    surface.text(String(pageNo), {
      x: SLIDE_W - M - 0.6, y: 5.05, w: 0.6, h: 0.3,
      pt: 10, bold: true, color: mix(on, f, 0.6), align: 'right', valign: 'middle',
    });
    return;
  }

  surface.background(t.bg);

  // 上端の色帯。ページをまたいで通る唯一の目印
  surface.rect(0, 0, SLIDE_W, BAND_H, f);

  surface.text(slide.heading, {
    x: M, y: HEAD_Y, w: HEAD_W, h: HEAD_H,
    pt: fitFontSize(slide.heading, HEAD_W, HEAD_MAX_PT, 15),
    bold: true, color: on, align: 'left', valign: 'middle', lineSpacing: 1.15,
  });

  // ページ番号も帯の中に入れる。下端に罫線を引かずに済む
  surface.text(String(pageNo), {
    x: SLIDE_W - M - 0.8, y: HEAD_Y, w: 0.8, h: HEAD_H,
    pt: PAGE_NO_PT, bold: true, color: mix(on, f, 0.68),
    align: 'right', valign: 'middle',
  });

  // 画像だけのときは下端まで裁ち落とす。上の帯にページ番号があるので番号は残る
  if (kind === 'media') {
    await drawMedia(surface, media, mediaFullRegion(MEDIA_FULL_TOP));
    return;
  }

  if (kind === 'split') {
    drawPlain(surface, t, slide.body, { x: M, y: CONTENT_TOP, w: TEXT_COL_W, h: CONTENT_H });
    const x = M + TEXT_COL_W + MEDIA_GAP;
    await drawMedia(surface, media, { x, y: CONTENT_TOP, w: SLIDE_W - M - x, h: CONTENT_H });
  } else {
    drawBody(surface, t, slide, BODY_OPT);
  }
}
