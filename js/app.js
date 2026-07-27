/* =========================================================
   スライドジェネレーター — 画面の制御
   質問に順番に答えるだけで .pptx を作る（サーバー・ログイン不要）

   ここは状態と DOM の担当。実際の描画・生成・保存は各モジュールへ。
     layout.js  … 寸法とレイアウト計算（pptx と プレビューで共有）
     theme.js   … 配色
     pptx.js    … .pptx の組み立て
     preview.js … 確認画面のプレビュー描画
     media.js   … 画像・動画の取り込み
     storage.js … 下書きの保存
   ========================================================= */

import { THEMES, DEFAULT_THEME, normalizeTheme } from './theme.js';
import { MAX_MEDIA } from './layout.js';
import { buildPptx, safeFileName } from './pptx.js';
import { drawCover, drawSlide } from './preview.js';
import { readMedia, ensureAspects } from './media.js';
import { saveDraft, loadDraft, clearDraft } from './storage.js';

const MAX_HISTORY = 40;
const INPUT_MAX_H = 134;
const SAVE_DEBOUNCE = 300;
const PREVIEW_DEBOUNCE = 250;
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

// ---------- 要素 ----------
const el = {
  chat:      document.getElementById('chat'),
  viewChat:  document.getElementById('view-chat'),
  viewRev:   document.getElementById('view-review'),
  composer:  document.getElementById('composer'),
  quick:     document.getElementById('quick'),
  input:     document.getElementById('input'),
  send:      document.getElementById('btn-send'),
  back:      document.getElementById('btn-back'),
  reset:     document.getElementById('btn-reset'),
  progress:  document.getElementById('progress'),
  file:      document.getElementById('file'),
  cards:     document.getElementById('cards'),
  themes:    document.getElementById('themes'),
  add:       document.getElementById('btn-add'),
  download:  document.getElementById('btn-download'),
  howto:     document.getElementById('howto'),
  toast:     document.getElementById('toast'),
  overlay:   document.getElementById('overlay'),
  overlayTx: document.getElementById('overlay-text'),
};

const isTouch = window.matchMedia('(pointer: coarse)').matches;

/* ファイル入力の設定は HTML にも書いてあるが、ここでも設定しておく。
   Service Worker が古い index.html を返しても、複数選択が効くようにするため。 */
el.file.multiple = true;
el.file.accept = 'image/*,video/*';

/* ---------- 状態 ----------
   phase: 'title' | 'subtitle' | 'heading' | 'body' | 'media' | 'review'
   i:     何枚目の本文スライドを聞いているか (0 始まり)
   slide = { heading, body, media: [] } */
let state = newState();
let history = [];
let saveFailed = false;
let pickTarget = null;   // 確認画面から追加するときのスライド添え字（チャット中は null）
let lastDeleted = null;  // 「元に戻す」用

function newState() {
  return {
    title: '', subtitle: '', slides: [],
    phase: 'title', i: 0, theme: DEFAULT_THEME, skipMedia: false, log: [],
  };
}

const currentSlide = () => state.slides[state.i];

// ---------- 質問文 ----------
const QUESTIONS = {
  title: () => ({
    text: 'まずは表紙です。\nスライドのタイトルを教えてください。',
    sub: '例）60期 設備管理委員',
  }),
  subtitle: () => ({
    text: '表紙のサブタイトルはありますか？',
    sub: '日付・部署名・発表者名など。なければ「なし」でOK',
  }),
  heading: (i) => ({
    text: `${i + 2}枚目の見出しを教えてください。`,
    sub: i === 0 ? '例）今期の振り返り' : null,
  }),
  body: (i) => ({
    text: `${i + 2}枚目の内容はありますか？`,
    sub: '任意です。改行で箇条書きになります（行の先頭に「・」でもOK）',
  }),
  media: (i) => ({
    text: `${i + 2}枚目に画像や動画を入れますか？`,
    sub: `任意です。1枚のスライドに最大${MAX_MEDIA}点まで並べられます`,
  }),
};

/* =========================================================
   チャット描画
   ========================================================= */

function pushLog(entry) {
  state.log.push(entry);
  renderMessage(entry);
  scrollToBottom();
}

function renderMessage(entry) {
  const box = document.createElement('div');
  box.className = 'msg ' + entry.role;

  if (entry.img) {
    const img = document.createElement('img');
    img.src = entry.img;
    img.alt = '選んだメディア';
    box.appendChild(img);
  }
  if (entry.text) box.appendChild(document.createTextNode(entry.text));
  if (entry.sub) {
    const s = document.createElement('span');
    s.className = 'sub';
    s.textContent = entry.sub;
    box.appendChild(s);
  }
  el.chat.appendChild(box);
}

function renderAllMessages() {
  el.chat.textContent = '';
  state.log.forEach(renderMessage);
  scrollToBottom();
}

function scrollToBottom() {
  requestAnimationFrame(() => { el.viewChat.scrollTop = el.viewChat.scrollHeight; });
}

function ask() {
  const q = QUESTIONS[state.phase](state.i);
  pushLog({ role: 'bot', text: q.text, sub: q.sub || null });
  updateComposer();
}

/* =========================================================
   入力欄の状態
   ========================================================= */

function updateComposer() {
  el.quick.textContent = '';
  el.back.disabled = history.length === 0;
  updateProgress();

  const chips = [];

  if (state.phase === 'media') {
    const media = (currentSlide() && currentSlide().media) || [];

    el.input.disabled = true;
    el.input.value = '';
    el.input.placeholder = '下のボタンから選んでください';
    el.send.disabled = true;

    if (media.length) {
      chips.push([`次へ（${media.length}点）`, 'accent', () => { beginTurn(); next(); }]);
    }
    if (media.length < MAX_MEDIA) {
      chips.push(['メディアを追加', media.length ? '' : 'accent', () => pickFile(null)]);
    }
    if (!media.length) {
      chips.push(['メディアなし（スキップ）', '', () => {
        beginTurn();
        pushLog({ role: 'user', text: 'メディアなし' });
        next();
      }]);
      chips.push(['以降すべてなし', '', () => {
        beginTurn();
        pushLog({ role: 'user', text: '以降すべてメディアなし' });
        state.skipMedia = true;
        pushLog({
          role: 'bot',
          text: 'これ以降は画像・動画の質問をとばします。',
          sub: '確認画面からはいつでも追加できます',
        });
        next();
      }]);
    }
  } else {
    el.input.disabled = false;
    el.send.disabled = false;
    el.input.placeholder = placeholderFor(state.phase);

    if (state.phase === 'subtitle') {
      chips.push(['サブタイトルなし', '', () => {
        beginTurn();
        pushLog({ role: 'user', text: 'なし' });
        state.subtitle = '';
        next();
      }]);
    }
    if (state.phase === 'body') {
      chips.push(['内容なし（スキップ）', '', () => {
        beginTurn();
        pushLog({ role: 'user', text: '内容なし' });
        currentSlide().body = '';
        next();
      }]);
    }
    if (state.phase === 'heading') {
      if (state.i >= 1) {
        chips.push([`ここまでで完成（全${state.i + 1}枚）`, 'accent', finish]);
      }
      if (state.skipMedia) {
        chips.push(['画像・動画の質問を戻す', '', () => {
          state.skipMedia = false;
          pushLog({ role: 'bot', text: '画像・動画の質問を再開します。' });
          save();
          updateComposer();
        }]);
      }
    }
  }

  chips.forEach(([label, cls, fn]) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'chip' + (cls ? ' ' + cls : '');
    b.textContent = label;
    b.addEventListener('click', fn);
    el.quick.appendChild(b);
  });

  autosize();
}

function placeholderFor(phase) {
  switch (phase) {
    case 'title':    return 'タイトルを入力…';
    case 'subtitle': return 'サブタイトルを入力…';
    case 'heading':  return '見出しを入力…';
    case 'body':     return '内容を入力（改行OK）…';
    default:         return 'ここに入力…';
  }
}

// 「いま何枚できているか」をヘッダーに出す
function updateProgress() {
  if (state.phase === 'review' || state.phase === 'title') {
    el.progress.textContent = '';
    return;
  }
  const done = state.slides.filter(s => s && s.heading).length;
  el.progress.textContent = `表紙＋${done}枚`;
}

/* =========================================================
   進行
   ========================================================= */

// 「戻る」用に、回答を記録する前の状態を控えておく。
// ログに吹き出しを足す前に呼ぶこと（後だと回答が残ったまま巻き戻ってしまう）
function beginTurn() {
  history.push(snapshot());
  if (history.length > MAX_HISTORY) history.shift();
}

function snapshot() {
  return {
    title: state.title,
    subtitle: state.subtitle,
    slides: state.slides.map(cloneSlide),
    phase: state.phase,
    i: state.i,
    theme: state.theme,
    skipMedia: state.skipMedia,
    log: state.log.slice(),
  };
}

function cloneSlide(s) {
  return { heading: s.heading, body: s.body, media: (s.media || []).map(m => ({ ...m })) };
}

function next() {
  switch (state.phase) {
    case 'title':
      state.phase = 'subtitle';
      break;
    case 'subtitle':
      state.phase = 'heading';
      state.i = 0;
      break;
    case 'heading':
      state.phase = 'body';
      break;
    case 'body':
      // 「以降すべてなし」を選んでいたらメディアの質問はとばす
      if (state.skipMedia) {
        state.i += 1;
        state.phase = 'heading';
      } else {
        state.phase = 'media';
      }
      break;
    case 'media':
      state.i += 1;
      state.phase = 'heading';
      break;
  }
  save();
  ask();
}

function handleSend() {
  const value = el.input.value.trim();
  if (!value || state.phase === 'media') return;

  beginTurn();
  pushLog({ role: 'user', text: value });
  el.input.value = '';
  autosize();

  switch (state.phase) {
    case 'title':
      state.title = value;
      break;
    case 'subtitle':
      state.subtitle = (value === 'なし' || value === 'ナシ') ? '' : value;
      break;
    case 'heading':
      state.slides[state.i] = { heading: value, body: '', media: [] };
      break;
    case 'body':
      currentSlide().body = value;
      break;
  }
  next();
}

function goBack() {
  if (!history.length) return;
  const prev = history.pop();
  state = { ...prev, slides: prev.slides.map(cloneSlide), log: prev.log.slice() };
  renderAllMessages();
  updateComposer();
  save();
}

function finish() {
  // 見出しだけ入って途中で終わった場合の後始末
  state.slides = state.slides.filter(s => s && s.heading);
  state.phase = 'review';
  save();
  showReview();
}

/* =========================================================
   メディアの追加
   ========================================================= */

function addMediaToSlide(idx, item) {
  const slide = state.slides[idx];
  if (!slide) return false;
  if (!Array.isArray(slide.media)) slide.media = [];
  if (slide.media.length >= MAX_MEDIA) return false;
  slide.media.push(item);
  return true;
}

// チャット中にメディアを足したあとの案内。上限に達したら自動で次へ進む
function noteMediaAdded(skipped) {
  if (skipped > 0) toast(`${skipped}点は上限（${MAX_MEDIA}点）を超えるため追加しませんでした`);

  if (currentSlide().media.length >= MAX_MEDIA) {
    pushLog({ role: 'bot', text: `${MAX_MEDIA}点になりました。次へ進みます。` });
    next();
  } else {
    pushLog({
      role: 'bot',
      text: `追加しました（${currentSlide().media.length}点）。`,
      sub: 'さらに追加するか、「次へ」で進んでください',
    });
    save();
    updateComposer();
  }
}

/* ファイル選択。target は確認画面のスライド添え字、チャット中は null。
   ここで必ず上書きするので、前回ダイアログをキャンセルした値が残らない
   （残ると次に選んだメディアが別のスライドに入ってしまう） */
function pickFile(target) {
  pickTarget = (typeof target === 'number') ? target : null;
  el.file.value = '';
  el.file.click();
}

el.file.addEventListener('change', async () => {
  const files = Array.from(el.file.files || []);
  if (!files.length) return;

  const target = pickTarget;
  const slideIdx = (target === null) ? state.i : target;
  const slide = state.slides[slideIdx];
  if (!slide) { el.file.value = ''; pickTarget = null; return; }

  const room = MAX_MEDIA - (slide.media || []).length;
  const accepted = files.slice(0, Math.max(0, room));
  const skipped = files.length - accepted.length;

  if (!accepted.length) {
    toast(`1枚につき${MAX_MEDIA}点までです`);
    el.file.value = '';
    pickTarget = null;
    return;
  }

  // チャット中は、読み込みが1点でも成功してから履歴を積む
  let turnStarted = false;
  let added = 0;

  for (let k = 0; k < accepted.length; k++) {
    const file = accepted[k];
    const isVideo = file.type.startsWith('video/');
    const count = accepted.length > 1 ? `（${k + 1}/${accepted.length}）` : '';
    showOverlay(isVideo ? `動画を読み込み中…${count}` : `画像を読み込み中…${count}`);

    try {
      const item = await readMedia(file);
      if (target === null && !turnStarted) { beginTurn(); turnStarted = true; }
      if (!addMediaToSlide(slideIdx, item)) break;
      added++;
      if (target === null) {
        pushLog({ role: 'user', img: item.cover || item.data, text: isVideo ? file.name : null });
      }
    } catch (err) {
      console.error(err);
      toast(err && err.message ? err.message : `${file.name} は読み込めませんでした`);
    }
  }

  hideOverlay();
  el.file.value = '';
  pickTarget = null;

  if (!added) return;
  if (target === null) {
    noteMediaAdded(skipped);
  } else {
    if (skipped > 0) toast(`${skipped}点は上限（${MAX_MEDIA}点）を超えるため追加しませんでした`);
    save();
    renderCards();
  }
});

/* =========================================================
   確認・編集画面
   ========================================================= */

function showReview() {
  el.viewChat.classList.add('hidden');
  el.composer.classList.add('hidden');
  el.viewRev.classList.remove('hidden');
  updateProgress();
  renderCards();
  el.viewRev.scrollTop = 0;
}

function showChat() {
  el.viewRev.classList.add('hidden');
  el.viewChat.classList.remove('hidden');
  el.composer.classList.remove('hidden');
  updateComposer();
  scrollToBottom();
}

function renderCards() {
  el.cards.textContent = '';
  el.cards.appendChild(coverCard());
  state.slides.forEach((slide, idx) => el.cards.appendChild(slideCard(slide, idx)));
  el.download.disabled = !state.title.trim();
  renderThemes();
  refreshPreviews();
}

function coverCard() {
  const card = div('card');
  card.dataset.preview = 'cover';

  const head = div('card-head');
  head.appendChild(badge('表紙'));
  card.appendChild(head);
  card.appendChild(previewBox(card));

  card.appendChild(textField('タイトル', state.title, (v) => {
    state.title = v;
    el.download.disabled = !v.trim();
    save();
    schedulePreview(card);
  }));
  card.appendChild(textField('サブタイトル', state.subtitle, (v) => {
    state.subtitle = v;
    save();
    schedulePreview(card);
  }));
  return card;
}

function slideCard(slide, idx) {
  const card = div('card');
  card.dataset.preview = String(idx);

  const head = div('card-head');
  head.appendChild(badge(`${idx + 2}枚目`));
  head.appendChild(miniBtn('↑', idx === 0, () => moveSlide(idx, -1)));
  head.appendChild(miniBtn('↓', idx === state.slides.length - 1, () => moveSlide(idx, 1)));
  head.appendChild(miniBtn('✕', false, () => deleteSlide(idx), 'del'));
  card.appendChild(head);

  card.appendChild(previewBox(card));

  card.appendChild(textField('見出し', slide.heading, (v) => {
    slide.heading = v;
    save();
    schedulePreview(card);
  }));
  card.appendChild(textField('内容', slide.body, (v) => {
    slide.body = v;
    save();
    schedulePreview(card);
  }, 3));

  const media = slide.media || [];
  if (media.length) {
    const grid = div('media-grid');
    media.forEach((m, mi) => grid.appendChild(mediaThumb(m, idx, mi, media.length)));
    card.appendChild(grid);
  }

  const add = div('media-add');
  if (media.length < MAX_MEDIA) {
    add.appendChild(linkBtn('＋ 画像・動画を追加', () => pickFile(idx)));
  } else {
    const note = document.createElement('span');
    note.className = 'media-note';
    note.textContent = `${MAX_MEDIA}点まで`;
    add.appendChild(note);
  }
  card.appendChild(add);

  return card;
}

function mediaThumb(m, slideIdx, mediaIdx, total) {
  const wrap = div('media-item');

  const frame = div('media-frame');
  if (m.cover || m.data) {
    const img = document.createElement('img');
    img.src = m.cover || m.data;
    img.alt = m.kind === 'video' ? '動画' : '画像';
    frame.appendChild(img);
  } else {
    const ph = div('media-ph');
    ph.textContent = '▶';
    frame.appendChild(ph);
  }
  if (m.kind === 'video') {
    const kind = document.createElement('span');
    kind.className = 'media-kind';
    kind.textContent = '動画';
    frame.appendChild(kind);
  }
  wrap.appendChild(frame);

  const bar = div('media-bar');
  bar.appendChild(miniBtn('‹', mediaIdx === 0, () => moveMedia(slideIdx, mediaIdx, -1), 'tiny'));
  bar.appendChild(miniBtn('✕', false, () => deleteMedia(slideIdx, mediaIdx), 'tiny del'));
  bar.appendChild(miniBtn('›', mediaIdx === total - 1, () => moveMedia(slideIdx, mediaIdx, 1), 'tiny'));
  wrap.appendChild(bar);

  return wrap;
}

/* ---------- 編集操作（削除は取り消せるようにする） ---------- */

function deleteSlide(idx) {
  const [removed] = state.slides.splice(idx, 1);
  lastDeleted = { type: 'slide', idx, data: removed };
  save();
  renderCards();
  toast(`${idx + 2}枚目を削除しました`, '元に戻す', undoDelete);
}

function deleteMedia(slideIdx, mediaIdx) {
  const [removed] = state.slides[slideIdx].media.splice(mediaIdx, 1);
  lastDeleted = { type: 'media', idx: slideIdx, mediaIdx, data: removed };
  save();
  renderCards();
  toast('メディアを削除しました', '元に戻す', undoDelete);
}

function undoDelete() {
  if (!lastDeleted) return;
  if (lastDeleted.type === 'slide') {
    state.slides.splice(lastDeleted.idx, 0, lastDeleted.data);
  } else {
    const slide = state.slides[lastDeleted.idx];
    if (!slide) { lastDeleted = null; return; }
    if (!Array.isArray(slide.media)) slide.media = [];
    slide.media.splice(lastDeleted.mediaIdx, 0, lastDeleted.data);
  }
  lastDeleted = null;
  save();
  renderCards();
}

function moveSlide(idx, delta) {
  const target = idx + delta;
  if (target < 0 || target >= state.slides.length) return;
  const [item] = state.slides.splice(idx, 1);
  state.slides.splice(target, 0, item);
  save();
  renderCards();
}

function moveMedia(slideIdx, mediaIdx, delta) {
  const media = state.slides[slideIdx].media;
  const target = mediaIdx + delta;
  if (target < 0 || target >= media.length) return;
  const [item] = media.splice(mediaIdx, 1);
  media.splice(target, 0, item);
  save();
  renderCards();
}

/* ---------- テーマ選択 ---------- */

function renderThemes() {
  el.themes.textContent = '';

  Object.keys(THEMES).forEach((key) => {
    const t = THEMES[key];
    const on = key === state.theme;

    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'theme-swatch' + (on ? ' on' : '');
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', String(on));
    b.title = t.label;

    // 表紙の見た目をそのまま縮めたプレビュー（本文＋下の色帯）
    const prev = div('theme-prev');
    prev.style.background = '#' + t.bg;
    for (const [cls, color] of [['theme-bar', t.ink], ['theme-line', t.line], ['theme-band', t.accent]]) {
      const part = div(cls);
      part.style.background = '#' + color;
      prev.appendChild(part);
    }

    const name = document.createElement('span');
    name.textContent = t.label;

    b.appendChild(prev);
    b.appendChild(name);
    b.addEventListener('click', () => {
      state.theme = key;
      save();
      renderThemes();
      refreshPreviews();
    });

    el.themes.appendChild(b);
  });
}

/* ---------- プレビュー ---------- */

const previewTimers = new WeakMap();

function previewBox(card) {
  const box = div('card-preview');
  const canvas = document.createElement('canvas');
  canvas.className = 'preview-canvas';
  box.appendChild(canvas);
  card._canvas = canvas;
  return box;
}

function schedulePreview(card) {
  clearTimeout(previewTimers.get(card));
  previewTimers.set(card, setTimeout(() => drawCard(card), PREVIEW_DEBOUNCE));
}

function refreshPreviews() {
  el.cards.querySelectorAll('.card').forEach(drawCard);
}

function drawCard(card) {
  const canvas = card._canvas;
  if (!canvas) return;
  const key = card.dataset.preview;
  if (key === 'cover') { drawCover(canvas, state); return; }
  const slide = state.slides[Number(key)];
  if (slide) drawSlide(canvas, state, slide, Number(key) + 2);
}

/* ---------- 小さなDOMヘルパー ---------- */

function div(cls) { const d = document.createElement('div'); d.className = cls; return d; }

function badge(text) { const b = div('badge'); b.textContent = text; return b; }

function miniBtn(label, disabled, fn, extra) {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'mini-btn' + (extra ? ' ' + extra : '');
  b.textContent = label;
  b.disabled = disabled;
  b.addEventListener('click', fn);
  return b;
}

function linkBtn(label, fn, extra) {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'link-btn' + (extra ? ' ' + extra : '');
  b.textContent = label;
  b.addEventListener('click', fn);
  return b;
}

function textField(label, value, onChange, rows) {
  const f = div('field');
  const l = document.createElement('label');
  l.textContent = label;
  const t = document.createElement('textarea');
  t.rows = rows || 1;
  t.value = value || '';
  t.addEventListener('input', () => onChange(t.value));
  f.appendChild(l);
  f.appendChild(t);
  return f;
}

/* =========================================================
   ダウンロード
   ========================================================= */

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function download() {
  if (!state.title.trim()) { toast('タイトルを入力してください'); return; }

  const hasVideo = state.slides.some(s => (s.media || []).some(m => m.kind === 'video'));
  showOverlay(hasVideo ? '動画を含めて作成中…' : 'スライドを作成中…');

  try {
    await ensureAspects(state.slides);
    const generated = await buildPptx(state);
    const blob = new Blob([generated], { type: PPTX_MIME });
    const fileName = safeFileName(state.title) + '.pptx';

    triggerDownload(blob, fileName);
    toast('PowerPointファイルをダウンロードしました');
    el.howto.open = true;
  } catch (err) {
    console.error(err);
    toast('作成に失敗しました。画像や動画を減らして試してください');
  } finally {
    hideOverlay();
  }
}

/* =========================================================
   保存 / 復元
   ========================================================= */

let saveTimer = null;

function save() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, SAVE_DEBOUNCE);
}

async function saveNow() {
  clearTimeout(saveTimer);
  const where = await saveDraft(snapshot());

  if (where === 'failed') {
    if (!saveFailed) toast('容量不足のため下書きを保存できませんでした');
    saveFailed = true;
    return;
  }
  if (where === 'local' && !saveFailed) {
    toast('この端末では動画を含まない下書きのみ保存します');
  }
  saveFailed = false;
}

// 旧形式（slide.image に1枚だけ）を media 配列へ移す
function migrateSlide(s) {
  const slide = { heading: s.heading || '', body: s.body || '', media: [] };
  if (Array.isArray(s.media)) {
    // data を持たないもの（保存対象外だった動画・廃止したYouTubeリンク）は落とす
    slide.media = s.media.filter(m => m && m.data && m.kind !== 'online');
  } else if (s.image) {
    slide.media = [{ kind: 'image', data: s.image }];
  }
  return slide;
}

function applyDraft(data) {
  if (!data || typeof data !== 'object' || !Array.isArray(data.log)) return false;
  if (!data.title && !data.log.length) return false;

  state = {
    title: data.title || '',
    subtitle: data.subtitle || '',
    slides: (Array.isArray(data.slides) ? data.slides : []).filter(Boolean).map(migrateSlide),
    phase: data.phase === 'image' ? 'media' : (data.phase || 'title'),
    i: typeof data.i === 'number' ? data.i : 0,
    theme: normalizeTheme(data.theme),
    skipMedia: Boolean(data.skipMedia),
    log: data.log,
  };
  return true;
}

async function reset() {
  if (!confirm('入力した内容をすべて消して、最初からやり直しますか？')) return;
  clearTimeout(saveTimer);
  await clearDraft();
  state = newState();
  history = [];
  lastDeleted = null;
  saveFailed = false;
  el.chat.textContent = '';
  showChat();
  start();
}

/* =========================================================
   UI 小物
   ========================================================= */

let toastTimer = null;

function toast(message, actionLabel, onAction) {
  el.toast.textContent = '';
  el.toast.appendChild(document.createTextNode(message));

  if (actionLabel && onAction) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'toast-action';
    b.textContent = actionLabel;
    b.addEventListener('click', () => {
      el.toast.classList.add('hidden');
      clearTimeout(toastTimer);
      onAction();
    });
    el.toast.appendChild(b);
  }

  el.toast.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.toast.classList.add('hidden'), actionLabel ? 6000 : 2600);
}

function showOverlay(text) {
  el.overlayTx.textContent = text;
  el.overlay.classList.remove('hidden');
}

function hideOverlay() { el.overlay.classList.add('hidden'); }

function autosize() {
  const t = el.input;
  t.style.height = 'auto';
  // border-box なので scrollHeight（＝ボーダーを含まない）に上下ボーダーを足さないと
  // 1〜2px 足りず、常にスクロールバーが出てしまう
  const cs = getComputedStyle(t);
  const border = parseFloat(cs.borderTopWidth) + parseFloat(cs.borderBottomWidth);
  const needed = t.scrollHeight + border;
  t.style.height = Math.min(needed, INPUT_MAX_H) + 'px';
  t.style.overflowY = needed > INPUT_MAX_H ? 'auto' : 'hidden';
}

/* =========================================================
   イベント
   ========================================================= */

el.send.addEventListener('click', handleSend);
el.back.addEventListener('click', goBack);
el.reset.addEventListener('click', reset);
el.input.addEventListener('input', autosize);
el.download.addEventListener('click', download);

el.input.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter') return;
  // スマホでは Enter は改行、PC では Enter で送信（Shift+Enter で改行）
  if (isTouch || e.shiftKey || e.isComposing) return;
  e.preventDefault();
  handleSend();
});

el.add.addEventListener('click', () => {
  state.phase = 'heading';
  state.i = state.slides.length;
  showChat();
  ask();
});

// 画面幅が変わるとプレビューの解像度が合わなくなるので描き直す
let resizeTimer = null;
window.addEventListener('resize', () => {
  if (el.viewRev.classList.contains('hidden')) return;
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(refreshPreviews, 200);
});

// 保存はまとめて行うので、離れる直前に書き切る
window.addEventListener('pagehide', saveNow);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') saveNow();
});

// 下書きは自動保存されるため、保存できていない時だけ引き止める
window.addEventListener('beforeunload', (e) => {
  if (saveFailed && (state.title || state.slides.length)) {
    e.preventDefault();
    e.returnValue = '';
  }
});

/* =========================================================
   起動
   ========================================================= */

function start() {
  pushLog({
    role: 'bot',
    text: 'こんにちは！質問に順番に答えるだけでスライドができます。',
    sub: '入力した内容はこの端末の中だけで処理されます（送信されません）',
  });
  ask();
}

async function init() {
  if (applyDraft(await loadDraft())) {
    renderAllMessages();
    if (state.phase === 'review') {
      showReview();
    } else {
      updateComposer();
      toast('前回の続きから再開しました');
    }
  } else {
    start();
  }
}

init();
