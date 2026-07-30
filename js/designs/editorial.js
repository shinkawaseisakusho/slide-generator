/* デザイン「標準」。

   余白と細い罫線だけで組む、雑誌の記事のような骨格。
   色の面積は表紙の帯と中扉に絞り、残りは余白で見せる。
   従来からある唯一のデザインで、見た目は変えていない。 */

import {
  SLIDE_W, SLIDE_H, M, CONTENT_W, CONTENT_TOP, CONTENT_H, FOOTER_Y,
  COVER_BAND_Y, COVER_TITLE_Y, COVER_TITLE_H, COVER_SUB_PT,
  COVER_RULE_TOP_Y, COVER_RULE_BOT_Y, COVER_ACCENT_W, COVER_ACCENT_H, COVER_ACCENT_Y,
  HEAD_MID, RULE_Y, RULE_ACCENT_Y, RULE_ACCENT_W, RULE_ACCENT_H,
  HAIRLINE_H, TEXT_COL_W, DIVIDER_MID, DIVIDER_RULE_Y, DIVIDER_RULE_W, PAGE_NO_DY,
  visibleMedia, slideKind, mediaRegion,
  titleSize, headingSize,
} from '../layout.js';
import { fillOf, onFillOf } from '../theme.js';
import { drawBody, drawPlain, drawMedia } from './common.js';

export const key = 'editorial';
export const label = '標準';

/* 画像だけのスライドの領域。本文より上下に広く取る。

   16:9 では正方形に近い写真が高さで頭を打つので、縦を広げるのが唯一効く。
   下端の罫線とページ番号には重ねないので、そのぶんは残す。 */
const MEDIA_ONLY = {
  x: M,
  y: RULE_Y + 0.14,
  w: CONTENT_W,
  h: (FOOTER_Y - 0.1) - (RULE_Y + 0.14),
};

const BODY_OPT = {
  region: { x: M, y: CONTENT_TOP, w: CONTENT_W, h: CONTENT_H },
  lead: 'rule',
  stat: 'divider',
  card: 'fill',
};

export function cover(surface, t, deck) {
  surface.background(t.bg);

  // タイトルを上下の細い罫線ではさむ
  surface.rect(M, COVER_RULE_TOP_Y, CONTENT_W, HAIRLINE_H, t.line);

  const titlePt = titleSize(deck.title);
  surface.text(deck.title, {
    x: M, y: COVER_TITLE_Y, w: CONTENT_W, h: COVER_TITLE_H,
    pt: titlePt, bold: true, color: t.ink,
    align: 'left', valign: 'middle', lineSpacing: 1.15,
  });

  surface.rect(M, COVER_RULE_BOT_Y, CONTENT_W, HAIRLINE_H, t.line);
  surface.rect(M, COVER_ACCENT_Y, COVER_ACCENT_W, COVER_ACCENT_H, fillOf(t));

  // 下端いっぱいの色帯。テーマの色がいちばん大きく出るところ
  surface.rect(0, COVER_BAND_Y, SLIDE_W, SLIDE_H - COVER_BAND_Y, fillOf(t));

  surface.text(deck.subtitle, {
    x: M, y: COVER_BAND_Y, w: CONTENT_W, h: SLIDE_H - COVER_BAND_Y,
    pt: COVER_SUB_PT, color: onFillOf(t), align: 'left', valign: 'middle',
  });
}

export async function slide(surface, t, slide, pageNo) {
  const kind = slideKind(slide);
  const media = visibleMedia(slide);

  // 見出しだけ → 章の区切り（中扉）。背景をテーマ色で塗って抜き文字にする
  if (kind === 'divider') {
    const on = onFillOf(t);
    surface.background(fillOf(t));
    surface.text(slide.heading, {
      x: M, y: DIVIDER_MID - 0.75, w: CONTENT_W, h: 1.5,
      pt: titleSize(slide.heading), bold: true, color: on,
      align: 'center', valign: 'middle',
    });
    surface.rect((SLIDE_W - DIVIDER_RULE_W) / 2, DIVIDER_RULE_Y, DIVIDER_RULE_W, HAIRLINE_H, on);
    pageNumber(surface, pageNo, on);
    return;
  }

  surface.background(t.bg);

  surface.text(slide.heading, {
    x: M, y: HEAD_MID - 0.33, w: CONTENT_W, h: 0.66,
    pt: headingSize(slide.heading), bold: true, color: t.accent,
    align: 'left', valign: 'middle',
  });

  // 細い線を全幅に引き、左端だけ太い色付きの線を重ねる
  surface.rect(M, RULE_Y, CONTENT_W, HAIRLINE_H, t.line);
  surface.rect(M, RULE_ACCENT_Y, RULE_ACCENT_W, RULE_ACCENT_H, fillOf(t));

  if (kind === 'split') {
    drawPlain(surface, t, slide.body, { x: M, y: CONTENT_TOP, w: TEXT_COL_W, h: CONTENT_H });
    await drawMedia(surface, media, mediaRegion(true));
  } else if (kind === 'media') {
    await drawMedia(surface, media, MEDIA_ONLY);
  } else {
    drawBody(surface, t, slide, BODY_OPT);
  }

  surface.rect(M, FOOTER_Y, CONTENT_W, HAIRLINE_H, t.line);
  pageNumber(surface, pageNo, t.muted);
}

function pageNumber(surface, pageNo, color) {
  surface.text(String(pageNo), {
    x: SLIDE_W - M - 0.5, y: FOOTER_Y + PAGE_NO_DY, w: 0.5, h: 0.3,
    pt: 9, color, align: 'right',
  });
}
