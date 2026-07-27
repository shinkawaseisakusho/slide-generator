/* スライドの寸法とレイアウト計算。
   pptx 生成（pptx.js）と画面プレビュー（preview.js）の両方から使う。
   ここを共通化しておかないと、プレビューと実際の仕上がりがずれる。
   単位はインチ（16:9 = 10 x 5.625）。 */

export const SLIDE_W = 10;
export const SLIDE_H = 5.625;
export const M = 0.75;                       // 左右の余白
export const CONTENT_W = SLIDE_W - M * 2;    // 8.5
export const GAP = 0.14;                     // メディア同士のすき間
export const CONTENT_TOP = 1.6;
export const FOOTER_Y = 5.16;                // 下端の罫線。本文はこの上まで使える
export const CONTENT_H = FOOTER_Y - 0.16 - CONTENT_TOP;
export const MAX_MEDIA = 4;                  // 1枚のスライドに載せられるメディアの数

export const HAIRLINE_H = 0.012;             // 細い罫線の太さ

export const HEAD_MID = 0.72;                // 見出しの中心
export const RULE_Y = 1.19;                  // 見出し下の細い罫線
export const RULE_ACCENT_W = 0.85;           // 左端に重ねる太い線
export const RULE_ACCENT_H = 0.042;
export const RULE_ACCENT_Y = RULE_Y - (RULE_ACCENT_H - HAIRLINE_H) / 2;

export const COVER_BAND_Y = 4.3;             // 表紙の色帯の上端
export const COVER_SUB_PT = 16;              // 色帯の中のサブタイトル

/* 表紙タイトルは、色帯より上の白地（0〜COVER_BAND_Y）の中央に置く。
   上下に細い罫線を渡し、下の罫線の左端だけ太いアクセントにする。
   本文スライドの見出しと同じ組み方なので、表紙だけ浮かない。 */
export const COVER_TITLE_H = 1.6;
export const COVER_TITLE_Y = (COVER_BAND_Y - COVER_TITLE_H) / 2;
export const COVER_TITLE_MID = COVER_TITLE_Y + COVER_TITLE_H / 2;
export const COVER_RULE_TOP_Y = COVER_TITLE_Y;
export const COVER_RULE_BOT_Y = COVER_TITLE_Y + COVER_TITLE_H;
export const COVER_ACCENT_W = 1.1;
export const COVER_ACCENT_H = 0.05;
export const COVER_ACCENT_Y = COVER_RULE_BOT_Y - (COVER_ACCENT_H - HAIRLINE_H) / 2;
export const TEXT_COL_W = 3.95;              // 文章とメディアを並べるときの文章側の幅
export const DIVIDER_MID = 2.7;              // 中扉の見出しの中心
export const DIVIDER_RULE_Y = 3.55;
export const DIVIDER_RULE_W = 1.05;
export const PAGE_NO_DY = 0.04;

/* ---------- スライドの種類 ---------- */

// 描けないメディア（保存で中身を落とした動画など）は数に入れない。
// 入れてしまうと「メディアあり」判定になり、空の余白だけのスライドになる
export function visibleMedia(slide) {
  return (slide.media || []).filter(m => m && m.data).slice(0, MAX_MEDIA);
}

// 内容とメディアの有無で4通り。'divider' は見出しだけの中扉
export function slideKind(slide) {
  const hasText = Boolean((slide.body || '').trim());
  const hasMedia = visibleMedia(slide).length > 0;
  if (!hasText && !hasMedia) return 'divider';
  if (hasText && hasMedia) return 'split';
  return hasMedia ? 'media' : 'text';
}

// メディアを置く領域。文章と並ぶときは右半分、単独なら全幅
export function mediaRegion(withText) {
  const x = withText ? M + TEXT_COL_W + 0.6 : M;
  return { x, y: CONTENT_TOP, w: SLIDE_W - M - x, h: CONTENT_H };
}

/* ---------- メディアの配置 ---------- */

// 領域を n 分割する。横長の領域なら横並び、縦長なら縦積み。4点は2x2。
export function gridCells(region, n) {
  const cells = [];
  const push = (x, y, w, h) => cells.push({ x, y, w, h });

  if (n === 1) { push(region.x, region.y, region.w, region.h); return cells; }

  if (n === 4) {
    const w = (region.w - GAP) / 2;
    const h = (region.h - GAP) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) push(region.x + c * (w + GAP), region.y + r * (h + GAP), w, h);
    }
    return cells;
  }

  // 2点・3点は領域の縦横比で並べ方を決める
  const horizontal = region.w / region.h > 1.5;
  if (horizontal) {
    const w = (region.w - GAP * (n - 1)) / n;
    for (let k = 0; k < n; k++) push(region.x + k * (w + GAP), region.y, w, region.h);
  } else {
    const h = (region.h - GAP * (n - 1)) / n;
    for (let k = 0; k < n; k++) push(region.x, region.y + k * (h + GAP), region.w, h);
  }
  return cells;
}

// 縦横比を保ったままセルの中央に収める。切り抜きも引き伸ばしもしない
export function fitRect(cell, aspect) {
  const a = aspect && isFinite(aspect) && aspect > 0 ? aspect : 4 / 3;
  let w = cell.w;
  let h = w / a;
  if (h > cell.h) { h = cell.h; w = h * a; }
  return { x: cell.x + (cell.w - w) / 2, y: cell.y + (cell.h - h) / 2, w, h };
}

/* ---------- 文字の大きさ ---------- */

/* 文字幅を「全角1文字ぶん（1em）」を単位として見積もる。
   フォントの実測値を使わないので、どのアプリで開いても同じ答えになる。
   実測より少し大きめに出て、結果として折り返しに余裕が出る側へ倒している。 */
export function textUnits(t) {
  let w = 0;
  for (const ch of String(t || '')) {
    if (ch === ' ') w += 0.28;
    else if (/[A-Z]/.test(ch)) w += 0.72;
    else if (/[a-z]/.test(ch)) w += 0.52;
    else if (/[0-9]/.test(ch)) w += 0.56;
    else if (/[!-~]/.test(ch)) w += 0.45;   // 半角記号
    else if (/[｡-ﾟ]/.test(ch)) w += 0.5;    // 半角カナ
    else w += 1;                             // 全角
  }
  return w;
}

/* 1行に収まる大きさを文字幅から決める。
   PowerPoint の自動縮小は高さがあふれた時しか効かず、幅の折り返しは防げないため
   自前で計算する。min を下回るほど長い文字列は、無理に縮めず折り返させる。 */
export function fitFontSize(text, boxW, max, min) {
  const units = textUnits(text);
  if (!units) return max;
  const byWidth = (boxW * 72 * 0.97) / units;   // 72pt = 1インチ。端に少し余裕を残す
  return Math.max(min, Math.min(max, Math.floor(byWidth)));
}

export function titleSize(t, boxW) {
  return fitFontSize(t, boxW || CONTENT_W, 40, 18);
}

export function headingSize(t, boxW) {
  return fitFontSize(t, boxW || CONTENT_W, 24, 13);
}

/* ---------- 本文の大きさ ----------

   PowerPoint の自動縮小（normAutofit）は、これを解釈しないアプリでは効かない。
   そのアプリでは指定どおりの大きさで描かれ、枠からあふれてレイアウトが崩れる。
   そこで「何行になるか」を自前で見積もり、最初から確実に収まる大きさを選ぶ。
   これで自動縮小に頼らなくてよくなり、どのアプリでも同じ見た目になる。 */

export const BODY_MAX_PT = 16;
export const BODY_MIN_PT = 9;
export const BODY_LINE_SPACING = 1.35;   // 行送り（フォントサイズに対する倍率）
export const BULLET_INDENT = 1.4;        // 箇条書きのぶら下げ幅（em）

/* 段落のあとの空き。固定値にすると、箇条書きが増えたときに
   この空きだけで枠を食いつぶしてしまうため、文字の大きさに比例させる。 */
export function paraSpacePt(sizePt) {
  return Math.round(sizePt * 0.3);
}

// 指定の大きさで何行になるかを見積もる
export function estimateBodyLines(body, boxW, sizePt) {
  const perLine = (boxW * 72) / sizePt;   // 1行に入る em 数
  let lines = 0;
  for (const line of bodyLines(body)) {
    if (!line.text) { lines += 0.6; continue; }   // 空行は詰めて数える
    const avail = Math.max(1, perLine - (line.bullet ? BULLET_INDENT : 0));
    lines += Math.max(1, Math.ceil(textUnits(line.text) / avail));
  }
  return lines;
}

// 指定の大きさで本文全体が何ポイントの高さになるか
export function bodyHeightPt(body, boxW, sizePt) {
  const paras = Math.max(0, bodyLines(body).length - 1);
  return estimateBodyLines(body, boxW, sizePt) * sizePt * BODY_LINE_SPACING
       + paras * paraSpacePt(sizePt);
}

// 枠に収まる最大の大きさ。最小まで下げても入らない場合は最小を返す
export function bodySize(body, boxW, boxH) {
  const limitPt = boxH * 72;
  for (let pt = BODY_MAX_PT; pt >= BODY_MIN_PT; pt--) {
    if (bodyHeightPt(body, boxW, pt) <= limitPt) return pt;
  }
  return BODY_MIN_PT;
}

// 1行を「箇条書きかどうか」と本文に分ける。行頭の ・ - * • を印として扱う
export function parseLine(line) {
  const trimmed = String(line).trim();
  const isBullet = /^([・\-*•])\s*/.test(trimmed);
  return { bullet: isBullet, text: isBullet ? trimmed.replace(/^([・\-*•])\s*/, '') : trimmed };
}

export function bodyLines(body) {
  return String(body || '').split(/\r?\n/).map(parseLine);
}
