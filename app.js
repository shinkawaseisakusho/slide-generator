/* =========================================================
   スライドジェネレーター
   質問に順番に答えるだけで .pptx を作る（サーバー・ログイン不要）
   ========================================================= */
(() => {
  'use strict';

  const STORAGE_KEY = 'slide-generator/v1';
  const MAX_IMAGE_EDGE = 1400;   // 画像の長辺の上限(px)
  const JPEG_QUALITY = 0.82;
  const MAX_HISTORY = 40;

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
    file:      document.getElementById('file'),
    cards:     document.getElementById('cards'),
    add:       document.getElementById('btn-add'),
    download:  document.getElementById('btn-download'),
    toast:     document.getElementById('toast'),
    overlay:   document.getElementById('overlay'),
    overlayTx: document.getElementById('overlay-text'),
  };

  const isTouch = window.matchMedia('(pointer: coarse)').matches;

  // ---------- 状態 ----------
  //   phase: 'title' | 'subtitle' | 'heading' | 'body' | 'image' | 'review'
  //   i:     何枚目の本文スライドを聞いているか (0 始まり)
  let state = newState();
  let history = [];
  let storageOK = true;
  let pendingImageTarget = null;  // 確認画面から画像を差し替えるときの添え字

  function newState() {
    return { title: '', subtitle: '', slides: [], phase: 'title', i: 0, log: [] };
  }

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
    image: (i) => ({
      text: `${i + 2}枚目に入れる画像はありますか？`,
      sub: '任意です。写真を撮る・アルバムから選ぶ、どちらでもOK',
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
    const div = document.createElement('div');
    div.className = 'msg ' + entry.role;

    if (entry.img) {
      const img = document.createElement('img');
      img.src = entry.img;
      img.alt = '選んだ画像';
      div.appendChild(img);
    }
    if (entry.text) {
      div.appendChild(document.createTextNode(entry.text));
    }
    if (entry.sub) {
      const s = document.createElement('span');
      s.className = 'sub';
      s.textContent = entry.sub;
      div.appendChild(s);
    }
    el.chat.appendChild(div);
  }

  function renderAllMessages() {
    el.chat.textContent = '';
    state.log.forEach(renderMessage);
    scrollToBottom();
  }

  function scrollToBottom() {
    requestAnimationFrame(() => {
      el.viewChat.scrollTop = el.viewChat.scrollHeight;
    });
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

    const chips = [];

    if (state.phase === 'image') {
      el.input.disabled = true;
      el.input.value = '';
      el.input.placeholder = '下のボタンから選んでください';
      el.send.disabled = true;
      chips.push(['画像を選ぶ', 'accent', pickImage]);
      chips.push(['画像なし（スキップ）', '', () => {
        beginTurn();
        pushLog({ role: 'user', text: '画像なし' });
        state.slides[state.i].image = null;
        next();
      }]);
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
          state.slides[state.i].body = '';
          next();
        }]);
      }
      if (state.phase === 'heading' && state.i >= 1) {
        chips.push([`ここまでで完成（全${state.i + 1}枚）`, 'accent', finish]);
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
      slides: state.slides.map(s => ({ ...s })),
      phase: state.phase,
      i: state.i,
      log: state.log.slice(),
    };
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
        state.phase = 'image';
        break;
      case 'image':
        state.i += 1;
        state.phase = 'heading';
        break;
    }
    save();
    ask();
  }

  function handleSend() {
    const value = el.input.value.trim();
    if (!value || state.phase === 'image') return;

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
        state.slides[state.i] = { heading: value, body: '', image: null };
        break;
      case 'body':
        state.slides[state.i].body = value;
        break;
    }
    next();
  }

  function goBack() {
    if (!history.length) return;
    const prev = history.pop();
    state = { ...prev, slides: prev.slides.map(s => ({ ...s })), log: prev.log.slice() };
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
     画像
     ========================================================= */

  function pickImage() {
    pendingImageTarget = null;
    el.file.value = '';
    el.file.click();
  }

  el.file.addEventListener('change', async () => {
    const file = el.file.files && el.file.files[0];
    if (!file) return;

    showOverlay('画像を読み込み中…');
    try {
      const dataUrl = await compressImage(file);

      if (pendingImageTarget !== null) {
        // 確認画面からの差し替え
        state.slides[pendingImageTarget].image = dataUrl;
        pendingImageTarget = null;
        save();
        showReview();
      } else {
        beginTurn();
        pushLog({ role: 'user', img: dataUrl });
        state.slides[state.i].image = dataUrl;
        next();
      }
    } catch (err) {
      console.error(err);
      toast('この画像は読み込めませんでした');
    } finally {
      hideOverlay();
      el.file.value = '';
    }
  });

  function compressImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('読み込み失敗'));
      reader.onload = () => {
        const img = new Image();
        img.onerror = () => reject(new Error('画像として開けません'));
        img.onload = () => {
          const scale = Math.min(1, MAX_IMAGE_EDGE / Math.max(img.width, img.height));
          const w = Math.max(1, Math.round(img.width * scale));
          const h = Math.max(1, Math.round(img.height * scale));

          // 透過を保ちたい小さめの PNG はそのまま PNG で出す
          const keepPng = file.type === 'image/png' && file.size < 1.5 * 1024 * 1024;

          const canvas = document.createElement('canvas');
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');
          if (!keepPng) {
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, w, h);
          }
          ctx.drawImage(img, 0, 0, w, h);

          resolve(keepPng
            ? canvas.toDataURL('image/png')
            : canvas.toDataURL('image/jpeg', JPEG_QUALITY));
        };
        img.src = reader.result;
      };
      reader.readAsDataURL(file);
    });
  }

  /* =========================================================
     確認・編集画面
     ========================================================= */

  function showReview() {
    el.viewChat.classList.add('hidden');
    el.composer.classList.add('hidden');
    el.viewRev.classList.remove('hidden');
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
  }

  function coverCard() {
    const card = div('card');
    const head = div('card-head');
    head.appendChild(badge('表紙'));
    card.appendChild(head);
    card.appendChild(textField('タイトル', state.title, v => { state.title = v; save(); el.download.disabled = !v.trim(); }));
    card.appendChild(textField('サブタイトル', state.subtitle, v => { state.subtitle = v; save(); }));
    return card;
  }

  function slideCard(slide, idx) {
    const card = div('card');

    const head = div('card-head');
    head.appendChild(badge(`${idx + 2}枚目`));
    head.appendChild(miniBtn('↑', idx === 0, () => move(idx, -1)));
    head.appendChild(miniBtn('↓', idx === state.slides.length - 1, () => move(idx, 1)));
    head.appendChild(miniBtn('✕', false, () => {
      if (!confirm(`${idx + 2}枚目を削除しますか？`)) return;
      state.slides.splice(idx, 1);
      save();
      renderCards();
    }, 'del'));
    card.appendChild(head);

    card.appendChild(textField('見出し', slide.heading, v => { slide.heading = v; save(); }));
    card.appendChild(textField('内容', slide.body, v => { slide.body = v; save(); }, 3));

    const row = div('thumb-row');
    if (slide.image) {
      const img = document.createElement('img');
      img.className = 'thumb';
      img.src = slide.image;
      img.alt = `${idx + 2}枚目の画像`;
      row.appendChild(img);
      row.appendChild(linkBtn('変更', () => replaceImage(idx)));
      row.appendChild(linkBtn('削除', () => { slide.image = null; save(); renderCards(); }, 'del'));
    } else {
      row.appendChild(linkBtn('＋ 画像を追加する', () => replaceImage(idx)));
    }
    card.appendChild(row);

    return card;
  }

  function replaceImage(idx) {
    pendingImageTarget = idx;
    el.file.value = '';
    el.file.click();
  }

  function move(idx, delta) {
    const target = idx + delta;
    if (target < 0 || target >= state.slides.length) return;
    const [item] = state.slides.splice(idx, 1);
    state.slides.splice(target, 0, item);
    save();
    renderCards();
  }

  // --- 小さなDOMヘルパー ---
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
     pptx 生成
     ========================================================= */

  const FONT = 'Noto Sans JP';
  const COLOR_TEXT = '1A1C23';
  const COLOR_MUTED = '6B7180';
  const COLOR_ACCENT = '4F46E5';
  const SLIDE_W = 10;
  const SLIDE_H = 5.625;

  async function buildPptx() {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = state.title || 'スライド';

    addCoverSlide(pptx);
    state.slides.forEach((slide, idx) => addContentSlide(pptx, slide, idx + 2));

    return pptx.write({ outputType: 'blob' });
  }

  function addCoverSlide(pptx) {
    const s = pptx.addSlide();
    s.background = { color: 'FFFFFF' };

    // 左端のアクセントバー
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 0.26, h: SLIDE_H, fill: { color: COLOR_ACCENT },
    });

    s.addText(state.title || '', {
      x: 0.95, y: 1.55, w: 8.4, h: 1.75,
      fontFace: FONT, fontSize: titleSize(state.title), bold: true,
      color: COLOR_TEXT, valign: 'bottom', align: 'left', fit: 'shrink',
    });

    if (state.subtitle) {
      s.addShape(pptx.ShapeType.rect, {
        x: 0.97, y: 3.48, w: 1.1, h: 0.045, fill: { color: COLOR_ACCENT },
      });
      s.addText(state.subtitle, {
        x: 0.95, y: 3.72, w: 8.4, h: 0.9,
        fontFace: FONT, fontSize: 17, color: COLOR_MUTED,
        valign: 'top', align: 'left', fit: 'shrink',
      });
    }
  }

  // 内容・画像はどちらも任意。有無の組み合わせでレイアウトを4通りに出し分ける
  function addContentSlide(pptx, slide, pageNo) {
    const s = pptx.addSlide();
    s.background = { color: 'FFFFFF' };

    const hasText = Boolean((slide.body || '').trim());
    const hasImage = Boolean(slide.image);

    // 見出しだけ → 章の区切り（中扉）として中央に大きく置く
    if (!hasText && !hasImage) {
      s.addText(slide.heading || '', {
        x: 0.9, y: 1.85, w: 8.2, h: 1.6,
        fontFace: FONT, fontSize: titleSize(slide.heading), bold: true,
        color: COLOR_TEXT, align: 'center', valign: 'middle', fit: 'shrink',
      });
      s.addShape(pptx.ShapeType.rect, {
        x: (SLIDE_W - 0.85) / 2, y: 3.62, w: 0.85, h: 0.05, fill: { color: COLOR_ACCENT },
      });
      addPageNo(s, pageNo);
      return;
    }

    s.addText(slide.heading || '', {
      x: 0.62, y: 0.38, w: 8.76, h: 0.8,
      fontFace: FONT, fontSize: headingSize(slide.heading), bold: true,
      color: COLOR_TEXT, valign: 'middle', align: 'left', fit: 'shrink',
    });
    s.addShape(pptx.ShapeType.rect, {
      x: 0.64, y: 1.24, w: 0.85, h: 0.05, fill: { color: COLOR_ACCENT },
    });

    const top = 1.58;
    const height = 3.42;

    if (hasText && hasImage) {
      s.addText(bodyParagraphs(slide.body), {
        x: 0.64, y: top, w: 4.36, h: height,
        fontFace: FONT, fontSize: bodySize(slide.body, true),
        color: COLOR_TEXT, valign: 'top', fit: 'shrink', lineSpacingMultiple: 1.25,
      });
      s.addImage({
        data: slide.image,
        x: 5.28, y: top, w: 4.1, h: height,
        sizing: { type: 'contain', w: 4.1, h: height },
      });
    } else if (hasImage) {
      s.addImage({
        data: slide.image,
        x: 0.64, y: top, w: 8.72, h: height,
        sizing: { type: 'contain', w: 8.72, h: height },
      });
    } else {
      s.addText(bodyParagraphs(slide.body), {
        x: 0.64, y: top, w: 8.72, h: height,
        fontFace: FONT, fontSize: bodySize(slide.body, false),
        color: COLOR_TEXT, valign: 'top', fit: 'shrink', lineSpacingMultiple: 1.3,
      });
    }

    addPageNo(s, pageNo);
  }

  function addPageNo(s, pageNo) {
    s.addText(String(pageNo), {
      x: 8.9, y: 5.12, w: 0.5, h: 0.3,
      fontFace: FONT, fontSize: 10, color: COLOR_MUTED, align: 'right',
    });
  }

  // 改行を段落に。「・」「-」始まりの行は箇条書きにする
  function bodyParagraphs(body) {
    const lines = String(body || '').split(/\r?\n/);
    return lines.map((line, i) => {
      const trimmed = line.trim();
      const isBullet = /^([・\-*•])\s*/.test(trimmed);
      return {
        text: isBullet ? trimmed.replace(/^([・\-*•])\s*/, '') : trimmed,
        options: {
          bullet: isBullet ? { characterCode: '2022' } : false,
          breakLine: i < lines.length - 1,
          paraSpaceAfter: 4,
        },
      };
    });
  }

  function titleSize(t) {
    const n = (t || '').length;
    if (n <= 16) return 40;
    if (n <= 28) return 32;
    return 26;
  }

  function headingSize(t) {
    const n = (t || '').length;
    if (n <= 20) return 26;
    if (n <= 34) return 22;
    return 18;
  }

  function bodySize(body, narrow) {
    const n = (body || '').length;
    const limit = narrow ? 0.55 : 1;
    if (n <= 90 * limit) return 17;
    if (n <= 200 * limit) return 15;
    if (n <= 400 * limit) return 13;
    return 11;
  }

  async function download() {
    if (!state.title.trim()) { toast('タイトルを入力してください'); return; }
    showOverlay('スライドを作成中…');
    try {
      const blob = await buildPptx();
      const name = safeFileName(state.title) + '.pptx';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      toast('ダウンロードしました');
      document.getElementById('howto').open = true;
    } catch (err) {
      console.error(err);
      toast('作成に失敗しました。画像を減らして試してください');
    } finally {
      hideOverlay();
    }
  }

  function safeFileName(name) {
    const cleaned = name.replace(/[\\/:*?"<>|\r\n]/g, '_').trim().slice(0, 60);
    return cleaned || 'スライド';
  }

  /* =========================================================
     保存 / 復元
     ========================================================= */

  function save() {
    if (!storageOK) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        title: state.title, subtitle: state.subtitle,
        slides: state.slides, phase: state.phase, i: state.i, log: state.log,
      }));
    } catch (err) {
      // 画像が多いと容量超過する。以後は保存をあきらめる（動作には影響なし）
      storageOK = false;
      try { localStorage.removeItem(STORAGE_KEY); } catch (_) { /* noop */ }
      toast('画像が多いため、下書きの自動保存は停止しました');
    }
  }

  function restore() {
    let raw = null;
    try { raw = localStorage.getItem(STORAGE_KEY); } catch (_) { return false; }
    if (!raw) return false;
    try {
      const data = JSON.parse(raw);
      if (!data || typeof data !== 'object' || !Array.isArray(data.log)) return false;
      if (!data.title && !data.log.length) return false;
      state = {
        title: data.title || '',
        subtitle: data.subtitle || '',
        slides: Array.isArray(data.slides) ? data.slides.filter(Boolean) : [],
        phase: data.phase || 'title',
        i: typeof data.i === 'number' ? data.i : 0,
        log: data.log,
      };
      return true;
    } catch (_) {
      return false;
    }
  }

  function reset() {
    if (!confirm('入力した内容をすべて消して、最初からやり直しますか？')) return;
    try { localStorage.removeItem(STORAGE_KEY); } catch (_) { /* noop */ }
    state = newState();
    history = [];
    storageOK = true;
    el.chat.textContent = '';
    showChat();
    start();
  }

  /* =========================================================
     UI 小物
     ========================================================= */

  let toastTimer = null;
  function toast(message) {
    el.toast.textContent = message;
    el.toast.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.toast.classList.add('hidden'), 2600);
  }

  function showOverlay(text) {
    el.overlayTx.textContent = text;
    el.overlay.classList.remove('hidden');
  }

  function hideOverlay() { el.overlay.classList.add('hidden'); }

  const INPUT_MAX_H = 134;

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

  el.download.addEventListener('click', download);

  window.addEventListener('beforeunload', (e) => {
    if (state.title || state.slides.length) { e.preventDefault(); e.returnValue = ''; }
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

  function init() {
    if (restore()) {
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
})();
