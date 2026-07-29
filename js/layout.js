import { ENABLE_VIDEO } from './config.js';

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

// 読み込める画像だけを数える。旧下書きに残った別形式はここでも除外する。
export function visibleMedia(slide) {
  return (slide.media || [])
    .filter(m => (
      m && m.data
      && (m.kind === 'image' || !m.kind || (ENABLE_VIDEO && m.kind === 'video'))
    ))
    .slice(0, MAX_MEDIA);
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
export function gridCells(region, n, gap = GAP) {
  const cells = [];
  const push = (x, y, w, h) => cells.push({ x, y, w, h });

  if (n === 1) { push(region.x, region.y, region.w, region.h); return cells; }

  if (n === 4) {
    const w = (region.w - gap) / 2;
    const h = (region.h - gap) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) push(region.x + c * (w + gap), region.y + r * (h + gap), w, h);
    }
    return cells;
  }

  // 2点・3点は領域の縦横比で並べ方を決める
  const horizontal = region.w / region.h > 1.5;
  if (horizontal) {
    const w = (region.w - gap * (n - 1)) / n;
    for (let k = 0; k < n; k++) push(region.x + k * (w + gap), region.y, w, region.h);
  } else {
    const h = (region.h - gap * (n - 1)) / n;
    for (let k = 0; k < n; k++) push(region.x, region.y + k * (h + gap), region.w, h);
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

// 空行を除いた本文の行
export function bodyItems(body) {
  return bodyLines(body).filter(l => l.text);
}

/* ---------- 本文の見せ方 ----------

   同じ「文章だけのスライド」でも、書かれ方によって最適な組み方は違う。
   短い一文はリード文、数値の並びは大きな数字、短い箇条書きはカード。
   ここで種類を決め、pptx.js と preview.js が同じ判定で描く。 */

export const CARD_MIN = 2;                   // カード／数値組みにする行数の下限
export const CARD_MAX = 4;                   // 同・上限（これを超えるとベタ組み）
export const CARD_GAP = 0.22;
export const CARD_PAD = 0.3;
export const CARD_RADIUS = 0.09;
export const CARD_NUM_PT = 11;               // 01, 02 … の連番
export const CARD_TITLE_MAX_PT = 17;
export const CARD_TITLE_MIN_PT = 10;
export const CARD_DESC_MAX_PT = 11;
export const CARD_DESC_MIN_PT = 8;

export const STAT_MAX_PT = 54;
export const STAT_MIN_PT = 22;
export const STAT_LABEL_PT = 12;
export const STAT_RULE_H = 1.15;             // 数値の間に立てる縦の細い罫線

export const LEAD_W = 7.2;                   // リード文の折り返し幅
export const LEAD_MAX_PT = 30;
export const LEAD_MIN_PT = 15;
export const LEAD_LINE_SPACING = 1.4;
export const LEAD_RULE_W = 0.9;
export const LEAD_RULE_H = 0.05;
export const LEAD_RULE_GAP = 0.28;           // 罫線と本文の間

const LEAD_MAX_UNITS = 46;                   // これより長い一文はベタ組みに戻す
const CARD_MAX_UNITS = 40;                   // カード1枚に入れられる長さ

/* 数値と単位。「30%」「1,200件」のような、そのままで意味が通る値だけを拾う。
   年号や日付を数値として拾わないよう、値とラベルの間に空白を必須にしている。 */
const UNIT = '%|％|倍|件|人|名|社|校|円|万円|億円|万|億|時間|分|秒|日|年|ヶ月|か月|カ月|週|回|台|個|pt|ポイント|x|×';
const VALUE = `[+\\-−]?[0-9][0-9,.]*\\s*(?:${UNIT})?`;
const VALUE_ONLY_RE = new RegExp(`^${VALUE}$`);
const VALUE_HEAD_RE = new RegExp(`^(${VALUE})[ 　]+(.+)$`);

/* 「30% 前年比」「売上：30%」「30%：前年比」を {value, label} に分ける。
   数値として読めない行は null を返し、呼び出し側でカード組みに落とす。 */
export function parseStat(text) {
  const s = String(text || '').trim();

  const parts = s.split(/\s*[:：]\s*/);
  if (parts.length === 2 && parts[0] && parts[1]) {
    if (VALUE_ONLY_RE.test(parts[1])) return { value: parts[1], label: parts[0] };
    if (VALUE_ONLY_RE.test(parts[0])) return { value: parts[0], label: parts[1] };
    return null;
  }

  const m = s.match(VALUE_HEAD_RE);
  return m ? { value: m[1].trim(), label: m[2].trim() } : null;
}

// 「見出し：説明」をカードの2段に分ける。区切りがなければ見出しだけのカード
export function parseCard(text) {
  const s = String(text || '').trim();
  const i = s.search(/[:：]/);
  if (i <= 0) return { title: s, desc: '' };
  return { title: s.slice(0, i).trim(), desc: s.slice(i + 1).trim() };
}

/* 本文の組み方を決める。カード・数値組みは「利用者が箇条書きとして書いた」
   ことを条件にしている。ふつうの文章が勝手にカードへ化けると読みにくい。 */
export function bodyStyle(slide) {
  if (slideKind(slide) !== 'text') return 'text';

  const items = bodyItems(slide.body);
  if (!items.length) return 'text';

  if (items.length === 1 && !items[0].bullet && textUnits(items[0].text) <= LEAD_MAX_UNITS) {
    return 'lead';
  }
  if (items.length < CARD_MIN || items.length > CARD_MAX) return 'text';
  if (!items.every(l => l.bullet)) return 'text';
  if (items.every(l => parseStat(l.text))) return 'stats';
  if (items.every(l => textUnits(l.text) <= CARD_MAX_UNITS)) return 'cards';
  return 'text';
}

export function contentRegion() {
  return { x: M, y: CONTENT_TOP, w: CONTENT_W, h: CONTENT_H };
}

/* ---------- リード文 ---------- */

// 罫線と本文をひとかたまりとして、本文領域の中央に置く
export function leadLayout(text) {
  const pt = fitFontSize(text, LEAD_W, LEAD_MAX_PT, LEAD_MIN_PT);
  const perLine = (LEAD_W * 72) / pt;
  const lines = Math.max(1, Math.ceil(textUnits(text) / perLine));
  const lh = (pt * LEAD_LINE_SPACING) / 72;
  const textH = lines * lh;
  const top = CONTENT_TOP + (CONTENT_H - (LEAD_RULE_H + LEAD_RULE_GAP + textH)) / 2;

  return { pt, lines, lh, textH, ruleY: top, textY: top + LEAD_RULE_H + LEAD_RULE_GAP };
}

/* ---------- 数値組み ---------- */

export function statLayout(items) {
  const stats = items.map(l => parseStat(l.text));
  const cells = gridCells(contentRegion(), stats.length, CARD_GAP);
  const colW = cells[0].w;

  // いちばん長い値に合わせる。数字の大きさがそろっていないと比較して見えない
  const valuePt = Math.min(...stats.map(s => fitFontSize(s.value, colW - 0.2, STAT_MAX_PT, STAT_MIN_PT)));
  const labelPt = Math.min(...stats.map(s => fitFontSize(s.label, colW - 0.2, STAT_LABEL_PT, 8)));

  const valueH = (valuePt * 1.2) / 72;
  const labelH = (labelPt * 1.5) / 72;
  const groupH = valueH + 0.12 + labelH;
  const top = CONTENT_TOP + (CONTENT_H - groupH) / 2;

  return { stats, cells, valuePt, labelPt, valueH, labelH, valueY: top, labelY: top + valueH + 0.12 };
}

/* ---------- カード組み ---------- */

export function cardLayout(items) {
  const cards = items.map(l => parseCard(l.text));
  const region = contentRegion();
  const grid = gridCells(region, cards.length, CARD_GAP);
  const rows = cards.length === 4 ? 2 : 1;
  const availH = grid[0].h;
  const innerW = grid[0].w - CARD_PAD * 2;

  const titlePt = Math.min(...cards.map(c => fitFontSize(c.title, innerW, CARD_TITLE_MAX_PT, CARD_TITLE_MIN_PT)));
  const titleLines = Math.max(...cards.map(
    c => Math.min(2, Math.max(1, Math.ceil(textUnits(c.title) / ((innerW * 72) / titlePt))))));

  const numH = (CARD_NUM_PT * 1.6) / 72;
  const titleH = (titleLines * titlePt * 1.3) / 72;
  const headH = CARD_PAD + numH + 0.1 + titleH;
  const hasDesc = cards.some(c => c.desc);
  const descY = headH + (hasDesc ? 0.12 : 0);

  /* 説明文は全カードで同じ大きさにそろえる。いちばん長いカードを基準に、
     カードの高さに収まるところまで下げる。 */
  let descPt = CARD_DESC_MAX_PT;
  let descH = 0;
  for (; descPt > CARD_DESC_MIN_PT; descPt--) {
    descH = Math.max(0, ...cards.map(c => (c.desc
      ? (estimateBodyLines(c.desc, innerW, descPt) * descPt * BODY_LINE_SPACING) / 72
      : 0)));
    if (descY + descH + CARD_PAD <= availH) break;
  }

  /* カードの高さは中身ぶんだけにして、余ったぶんは上下に振り分ける。
     領域いっぱいに引き伸ばすと、短い文言のときに間延びして見える。 */
  const cardH = Math.min(availH, descY + descH + CARD_PAD);
  const usedH = cardH * rows + (rows - 1) * CARD_GAP;
  const top = region.y + (region.h - usedH) / 2;
  const cells = grid.map((c, i) => ({
    x: c.x, w: c.w, h: cardH,
    y: top + (cards.length === 4 ? Math.floor(i / 2) : 0) * (cardH + CARD_GAP),
  }));

  return {
    cards, cells, innerW, titlePt, titleH, descPt,
    descH: Math.min(descH, cardH - descY - CARD_PAD),
    numY: CARD_PAD,
    titleY: CARD_PAD + numH + 0.1,
    descY,
  };
}
