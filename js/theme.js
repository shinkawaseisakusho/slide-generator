/* テーマは配色だけを差し替える。余白と罫線で構成する骨格は共通。

   bg=背景 / ink=本文 / muted=補助テキスト / line=細い罫線
   accent=文字に使うテーマ色（見出し、数値、連番）
   fill  =面と線に使うテーマ色（表紙の帯、中扉の背景、カードの塗り）
   onFill=fill の上に載せる文字の色

   fill と onFill は省略でき、その場合は accent と bg を使う。
   いまはどのテーマも省略している。accent を鮮やかにすると
   その上の白抜き文字が読めなくなる配色を足したくなったときに、
   面と文字の色を切り離すための逃げ道として用意してある。

   追加・変更するときの条件：
     - accent と bg のコントラスト比を 4.5 以上
     - fill を指定する場合、fill と onFill のコントラスト比を 4.5 以上
     - accent はどれも暗めになるため、色相だけでなく明度（L*）にも
       差をつける。既存のどの accent とも知覚差 ΔE が 30 以上を目安に
       （オレンジと赤だけは色の指定を優先していて 26）
     - 彩度（C*）が 20 を下回ると、暗い色は黒っぽく見えて濁る */

export const THEMES = {
  mono:   { label: '白黒',     bg: 'FFFFFF', ink: '111113', muted: '86868C', line: 'E4E4E7', accent: '111113' },
  blue:   { label: '青',       bg: 'FFFFFF', ink: '15243A', muted: '7A8698', line: 'E2E7EE', accent: '1F5FD0' },
  green:  { label: '緑',       bg: 'FFFFFF', ink: '152A20', muted: '78887F', line: 'E0E8E3', accent: '1E7A4B' },
  orange: { label: 'オレンジ', bg: 'FFFFFF', ink: '2E1D16', muted: '93807A', line: 'EFE3DC', accent: 'B54A17' },
  red:    { label: '赤',       bg: 'FFFFFF', ink: '2F1719', muted: '947E81', line: 'F0DCDE', accent: 'AB1F2E' },
  purple: { label: '紫',       bg: 'FFFFFF', ink: '221A33', muted: '837A94', line: 'E6E0F0', accent: '6A3FA0' },
  beige:  { label: 'ベージュ', bg: 'F7F4EF', ink: '24211C', muted: '8B8377', line: 'E6DFD3', accent: '7B6338' },
  dark:   { label: '黒',       bg: '111113', ink: 'F4F4F5', muted: '9A9AA1', line: '2E2E34', accent: 'F4F4F5' },
};

export const DEFAULT_THEME = 'mono';

// 旧バージョンで保存されたテーマ名
const LEGACY_THEMES = { navy: 'blue', forest: 'green', terra: 'orange', sand: 'beige', night: 'dark' };

export function normalizeTheme(key) {
  const k = LEGACY_THEMES[key] || key;
  return THEMES[k] ? k : DEFAULT_THEME;
}

export function getTheme(key) {
  return THEMES[normalizeTheme(key)];
}

/* 2色を混ぜる。ratio は a の割合。テーマに色を足さずに
   「背景をほんの少しだけアクセント寄りにした面」を作るために使う。 */
export function mix(a, b, ratio) {
  const hex = h => [0, 2, 4].map(i => parseInt(String(h).slice(i, i + 2), 16));
  const [ar, ag, ab] = hex(a);
  const [br, bg, bb] = hex(b);
  const ch = (x, y) => Math.round(x * ratio + y * (1 - ratio)).toString(16).padStart(2, '0');
  return (ch(ar, br) + ch(ag, bg) + ch(ab, bb)).toUpperCase();
}

/** 面と線に使うテーマ色。省略しているテーマは accent をそのまま使う */
export function fillOf(t) {
  return t.fill || t.accent;
}

/** fill の上に載せる文字の色。省略しているテーマは背景色で抜く */
export function onFillOf(t) {
  return t.onFill || t.bg;
}

// カードの下地。背景との差はごくわずかで、面があることだけが分かる濃さ
export function cardSurface(t) {
  return mix(fillOf(t), t.bg, 0.07);
}
