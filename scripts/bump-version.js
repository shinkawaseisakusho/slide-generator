/* sw.js のキャッシュ版数を上げる。
   公開のたびに手で書き換えると忘れるので、`npm run bump` で自動化する。

   使い方:
     npm run bump          → v2 → v3 のように 1 つ上げる
     npm run bump -- 7     → v7 に設定する
*/
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
const SW = path.join(here, '..', 'sw.js');
const PATTERN = /(const CACHE = 'slide-generator-v)(\d+)(';)/;

const src = fs.readFileSync(SW, 'utf8');
const m = PATTERN.exec(src);

if (!m) {
  console.error('sw.js の CACHE 行が見つかりませんでした。');
  console.error("期待する形: const CACHE = 'slide-generator-v1';");
  process.exit(1);
}

const current = Number(m[2]);
const arg = process.argv[2];
const next = arg === undefined ? current + 1 : Number(arg);

if (!Number.isInteger(next) || next < 1) {
  console.error(`版数が不正です: ${arg}`);
  process.exit(1);
}

if (next === current) {
  console.log(`版数は v${current} のままです（変更なし）。`);
  process.exit(0);
}

fs.writeFileSync(SW, src.replace(PATTERN, `$1${next}$3`));
console.log(`キャッシュ版数を v${current} → v${next} に更新しました。`);
