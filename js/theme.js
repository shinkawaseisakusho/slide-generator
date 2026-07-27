/* テーマは配色だけを差し替える。余白と罫線で構成する骨格は共通。
   bg=背景 / ink=本文 / muted=補助テキスト / line=細い罫線
   accent=見出しの文字色、表紙の帯、中扉の背景。色の面積が出る主役

   テーマ色の上には bg の抜き文字を置くため、追加するときは
   bg と accent のコントラスト比を 4.5 以上にすること。 */

export const THEMES = {
  mono:   { label: '白黒',     bg: 'FFFFFF', ink: '111113', muted: '86868C', line: 'E4E4E7', accent: '111113' },
  blue:   { label: '青',       bg: 'FFFFFF', ink: '15243A', muted: '7A8698', line: 'E2E7EE', accent: '1F5FD0' },
  green:  { label: '緑',       bg: 'FFFFFF', ink: '152A20', muted: '78887F', line: 'E0E8E3', accent: '1E7A4B' },
  orange: { label: 'オレンジ', bg: 'FFFFFF', ink: '2E1D16', muted: '93807A', line: 'EFE3DC', accent: 'B54A17' },
  red:    { label: '赤',       bg: 'FFFFFF', ink: '2F1817', muted: '94807F', line: 'F0DEDD', accent: 'B5302A' },
  purple: { label: '紫',       bg: 'FFFFFF', ink: '221A33', muted: '837A94', line: 'E6E0F0', accent: '6A3FA0' },
  beige:  { label: 'ベージュ', bg: 'F7F4EF', ink: '24211C', muted: '8B8377', line: 'E6DFD3', accent: '6E5F45' },
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
