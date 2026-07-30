/* 確認画面に出すスライドのプレビュー。

   pptx.js と同じ designs/*.js を呼ぶので、実際の仕上がりとずれない。
   このファイルの仕事は canvas を実寸に合わせることだけ。 */

import { getTheme } from './theme.js';
import { getDesign } from './designs/index.js';
import { canvasSurface } from './surface.js';
import { SLIDE_W, SLIDE_H } from './layout.js';

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
  return canvasSurface(ctx, S, S / 72, SLIDE_W, SLIDE_H);
}

/** 表紙を描く */
export function drawCover(canvas, deck) {
  getDesign(deck.design).cover(setup(canvas), getTheme(deck.theme), deck);
}

/** 本文スライドを描く（画像の読み込みを待つので非同期） */
export async function drawSlide(canvas, deck, slide, pageNo) {
  await getDesign(deck.design).slide(setup(canvas), getTheme(deck.theme), slide, pageNo);
}
