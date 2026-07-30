/* デザインの一覧。

   各デザインは cover(surface, theme, deck) と
   slide(surface, theme, slide, pageNo) の2つを持つ。
   描画先の違い（pptx / canvas）は surface.js が吸収するので、
   デザインは1回書けば両方に反映される。

   テーマ（配色）とは独立した軸で、掛け合わせて使う。 */

import * as editorial from './editorial.js';
import * as bold from './bold.js';
import * as soft from './soft.js';

export const DESIGNS = { editorial, bold, soft };
export const DEFAULT_DESIGN = 'editorial';

// 旧バージョンで保存されたデザイン名
const LEGACY_DESIGNS = { impact: 'bold' };

export function normalizeDesign(key) {
  const k = LEGACY_DESIGNS[key] || key;
  return DESIGNS[k] ? k : DEFAULT_DESIGN;
}

export function getDesign(key) {
  return DESIGNS[normalizeDesign(key)];
}
