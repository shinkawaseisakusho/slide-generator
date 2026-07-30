/* デザイン3種で共通の描画部品。

   本文の組み方（リード文・数値組み・カード組み・ベタ組み）と
   メディアの配置は、どのデザインでも中身の判定は同じにしたい。
   違うのは「どんな装飾で見せるか」だけなので、その差は引数で受ける。

   surface は surface.js の API。pptx でも canvas でも同じコードが動く。 */

import { cardSurface, fillOf, onFillOf, mix } from '../theme.js';
import {
  MAX_MEDIA,
  gridCells, fitRect, bodyItems, bodyStyle, bodySize, bodyLines, paraSpacePt,
  leadLayout, statLayout, cardLayout,
  BODY_LINE_SPACING, BULLET_INDENT,
  CARD_PAD, CARD_RADIUS, CARD_GAP, HAIRLINE_H,
  LEAD_RULE_W, LEAD_RULE_H, LEAD_LINE_SPACING, STAT_RULE_H,
} from '../layout.js';

const STAT_TOP_RULE_W = 0.6;      // 数値の上に置く短い帯

/**
 * 文章だけのスライドの本文を描く。
 *
 * @param opt.region     本文を置く領域
 * @param opt.lead       'rule'（罫線を添える）| 'plain'
 * @param opt.leadMaxPt  リード文の上限
 * @param opt.stat       'divider'（縦罫線で区切る）| 'plain'
 * @param opt.statRule   'top' なら値の上に短い色の帯を置く
 * @param opt.card       'fill'（淡い下地）| 'solid'（テーマ色で塗って抜き文字）
 *                       | 'bar'（上に色の帯）| 'outline'
 * @param opt.cardRadius 角丸の半径
 * @param opt.cardNumber 連番を出すか。'badge' なら丸で囲む
 */
export function drawBody(surface, t, slide, opt) {
  const style = bodyStyle(slide);
  const items = bodyItems(slide.body);

  if (style === 'lead') return drawLead(surface, t, items[0].text, opt);
  if (style === 'stats') return drawStats(surface, t, items, opt);
  if (style === 'cards') return drawCards(surface, t, items, opt);
  drawPlain(surface, t, slide.body, opt.region);
}

// ベタ組み。枠に収まる大きさをあらかじめ計算しているので自動縮小に頼らない
export function drawPlain(surface, t, body, region) {
  const pt = bodySize(body, region.w, region.h);
  surface.paragraphs(bodyLines(body), {
    x: region.x, y: region.y, w: region.w, h: region.h,
    pt, color: t.ink, valign: 'top',
    lineSpacing: BODY_LINE_SPACING,
    paraSpace: paraSpacePt(pt),
    bulletIndent: BULLET_INDENT,
  });
}

// 短い一文。本文領域の中央に大きく置く
function drawLead(surface, t, text, opt) {
  const L = leadLayout(text, opt.region, opt.leadMaxPt);

  if (opt.lead !== 'plain') {
    surface.rect(opt.region.x, L.ruleY, LEAD_RULE_W, LEAD_RULE_H, fillOf(t));
  }
  surface.text(text, {
    x: opt.region.x, y: L.textY, w: L.w, h: L.textH,
    pt: L.pt, bold: true, color: opt.leadColor || t.ink,
    align: 'left', valign: 'top', lineSpacing: LEAD_LINE_SPACING,
  });
}

// 数値の並び。値をそろえて大きく出す
function drawStats(surface, t, items, opt) {
  const L = statLayout(items, opt.region);

  L.cells.forEach((cell, k) => {
    if (k > 0 && opt.stat !== 'plain') {
      surface.rect(
        cell.x - CARD_GAP / 2 - HAIRLINE_H / 2,
        opt.region.y + (opt.region.h - STAT_RULE_H) / 2,
        HAIRLINE_H, STAT_RULE_H, t.line,
      );
    }

    if (opt.statRule === 'top') {
      surface.rect(cell.x + (cell.w - STAT_TOP_RULE_W) / 2, L.valueY - 0.26,
                   STAT_TOP_RULE_W, 0.055, fillOf(t));
    }

    surface.text(L.stats[k].value, {
      x: cell.x, y: L.valueY, w: cell.w, h: L.valueH,
      pt: L.valuePt, bold: true, color: t.accent, align: 'center', valign: 'middle',
    });
    surface.text(L.stats[k].label, {
      x: cell.x, y: L.labelY, w: cell.w, h: L.labelH,
      pt: L.labelPt, color: t.muted, align: 'center', valign: 'middle',
    });
  });
}

// 短い箇条書きをカードに並べ替える
function drawCards(surface, t, items, opt) {
  const L = cardLayout(items, opt.region);
  const radius = opt.cardRadius == null ? CARD_RADIUS : opt.cardRadius;
  const numbered = opt.cardNumber !== false;

  /* テーマ色でベタ塗りするときは、その上の文字を onFill で抜く。
     説明文と連番は少し沈ませて、見出しとの主従を保つ。 */
  const solid = opt.card === 'solid';
  const f = fillOf(t);
  const on = onFillOf(t);
  const titleColor = solid ? on : t.ink;
  const descColor = solid ? mix(on, f, 0.75) : t.muted;
  const numColor = solid ? mix(on, f, 0.55) : t.accent;

  L.cells.forEach((cell, k) => {
    const card = L.cards[k];
    const x = cell.x + CARD_PAD;
    let numY = cell.y + L.numY;
    let titleY = cell.y + L.titleY;
    let descY = cell.y + L.descY;

    if (opt.card === 'bar') {
      // 下地を敷かず、上端に太い色の帯を渡す。線だけで組む骨格に合う
      surface.rect(cell.x, cell.y, cell.w, 0.055, f);
      // 帯のぶん、中身をすこし下げる
      numY += 0.06;
      titleY += 0.06;
      descY += 0.06;
    } else if (opt.card === 'outline') {
      surface.roundRect(cell.x, cell.y, cell.w, cell.h, radius, t.line);
      surface.roundRect(cell.x + HAIRLINE_H, cell.y + HAIRLINE_H,
                        cell.w - HAIRLINE_H * 2, cell.h - HAIRLINE_H * 2, radius, t.bg);
    } else {
      surface.roundRect(cell.x, cell.y, cell.w, cell.h, radius, solid ? f : cardSurface(t));
    }

    if (numbered) {
      if (opt.cardNumber === 'badge') {
        const d = L.numH * 1.06;
        surface.ellipse(x, numY, d, d, solid ? on : f);
        surface.text(String(k + 1).padStart(2, '0'), {
          x, y: numY, w: d, h: d,
          pt: L.numPt - 1, bold: true, color: solid ? f : on,
          align: 'center', valign: 'middle',
        });
      } else {
        surface.text(String(k + 1).padStart(2, '0'), {
          x, y: numY, w: L.innerW, h: L.numH,
          pt: L.numPt, bold: true, color: numColor, align: 'left', valign: 'middle',
        });
      }
    }

    surface.text(card.title, {
      x, y: titleY, w: L.innerW, h: L.titleH,
      pt: L.titlePt, bold: true, color: titleColor,
      align: 'left', valign: 'top', lineSpacing: 1.3,
    });

    if (card.desc) {
      surface.text(card.desc, {
        x, y: descY, w: L.innerW, h: L.descH,
        pt: L.descPt, color: descColor, align: 'left', valign: 'top',
        lineSpacing: BODY_LINE_SPACING,
      });
    }
  });
}

/* ---------- メディア ---------- */

// 画像を領域に敷き詰める。縦横比は保ったままセルの中央に収める
export async function drawMedia(surface, media, region) {
  const cells = gridCells(region, Math.min(media.length, MAX_MEDIA));

  for (let k = 0; k < cells.length; k++) {
    await surface.media(media[k], fitRect(cells[k], media[k].aspect));
  }
}
