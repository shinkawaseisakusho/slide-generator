/* .pptx の組み立て。
   PptxGenJS は vendor/pptxgen.bundle.js がグローバルに読み込む。

   実際の描画は designs/*.js が担い、ここは箱を用意するだけ。
   同じデザインのコードが preview.js からも呼ばれるので、
   プレビューと仕上がりがずれない。

   deck = { title, subtitle, theme, design, slides: [{ heading, body, media: [] }] } */

import { getTheme } from './theme.js';
import { getDesign } from './designs/index.js';
import { pptxSurface } from './surface.js';

export async function buildPptx(deck, outputType = 'blob') {
  const Ctor = globalThis.PptxGenJS;
  if (!Ctor) throw new Error('PptxGenJS が読み込まれていません');

  const t = getTheme(deck.theme);
  const design = getDesign(deck.design);

  const pptx = new Ctor();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = deck.title || 'スライド';

  design.cover(pptxSurface(pptx, pptx.addSlide()), t, deck);

  const slides = deck.slides || [];
  for (let i = 0; i < slides.length; i++) {
    await design.slide(pptxSurface(pptx, pptx.addSlide()), t, slides[i], i + 2);
  }

  return pptx.write({ outputType });
}

export function safeFileName(name) {
  const cleaned = String(name || '').replace(/[\\/:*?"<>|\r\n]/g, '_').trim().slice(0, 60);
  return cleaned || 'スライド';
}
