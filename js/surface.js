/* 描画の抽象化層。

   デザイン（designs/*.js）はここで定義する API に向かって描き、
   pptx への出力と画面プレビューの違いはこのファイルだけが知っている。

   これがないと「デザインの数 × 出力先の数」だけ描画コードが増え、
   プレビューと実際の仕上がりが必ずどこかでずれる。

   座標と大きさの単位はインチ、文字の大きさはポイント。 */

/* Noto Sans JP を指定する。web で広く配布されているフォントなので、
   pptx を読み込むアプリ側で置き換えが起きにくい。 */
export const PPTX_FONT = 'Noto Sans JP';
export const CANVAS_FONT = '"Hiragino Sans", "Noto Sans JP", "Yu Gothic UI", Meiryo, sans-serif';

/* テキストボックスの内側余白。既定値はアプリによって違い、そのぶん
   折り返し位置がずれる。0 に固定して、指定した幅をそのまま使わせる。 */
const INSET = { margin: 0 };

/* ---------- pptx 用 ---------- */

export function pptxSurface(pptx, slide) {
  const textOpts = (o) => ({
    ...INSET,
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: PPTX_FONT,
    fontSize: o.pt,
    bold: Boolean(o.bold),
    color: o.color,
    align: o.align || 'left',
    valign: o.valign || 'top',
    charSpacing: o.charSpacing || 0,
    fit: o.shrink === false ? 'none' : 'shrink',
    ...(o.lineSpacing ? { lineSpacing: Math.round(o.pt * o.lineSpacing) } : {}),
  });

  return {
    background(color) {
      slide.background = { color };
    },

    rect(x, y, w, h, color) {
      slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color } });
    },

    roundRect(x, y, w, h, r, color) {
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y, w, h, rectRadius: r, fill: { color }, line: { color },
      });
    },

    ellipse(x, y, w, h, color) {
      slide.addShape(pptx.ShapeType.ellipse, { x, y, w, h, fill: { color }, line: { color } });
    },

    text(str, o) {
      if (!String(str || '')) return;
      slide.addText(String(str), textOpts(o));
    },

    /* 箇条書きを含む本文。pptx では段落ごとの options で表現する。
       中黒は PowerPoint 側の箇条書き機能に任せる（ぶら下げが自然になる）。 */
    paragraphs(lines, o) {
      if (!lines.length) return;
      const space = Math.round(o.paraSpace || 0);
      slide.addText(lines.map((line, i) => ({
        text: line.text,
        options: {
          bullet: line.bullet ? { characterCode: '2022' } : false,
          breakLine: i < lines.length - 1,
          paraSpaceAfter: space,
        },
      })), textOpts(o));
    },

    async media(m, box) {
      if (m.kind === 'image' || !m.kind) {
        // pptxgenjs の sizing:contain は data URL だと寸法を計算できず、
        // 指定した箱にそのまま引き伸ばされてしまう。自分で矩形を出して渡す
        slide.addImage({ data: m.data, ...box });
        return;
      }
      if (m.kind === 'video') {
        const opts = { type: 'video', data: m.data, extn: videoExtension(m), ...box };
        if (m.cover) opts.cover = m.cover;
        slide.addMedia(opts);
      }
    },
  };
}

// 動画機能は現在無効だが、将来の再有効化に備えて埋め込み処理を残す。
function videoExtension(media) {
  const fromName = String(media.name || '').match(/\.([a-z0-9]{2,5})$/i);
  if (fromName) return fromName[1].toLowerCase();
  if (media.extn) return String(media.extn).toLowerCase();

  const subtype = String(media.data || '').match(/^data:video\/([^;,]+)/i)?.[1]?.toLowerCase();
  if (subtype === 'quicktime') return 'mov';
  if (subtype === 'x-m4v') return 'm4v';
  return subtype || 'mp4';
}

/* ---------- canvas 用 ---------- */

const imgCache = new Map();

function loadImage(src) {
  let p = imgCache.get(src);
  if (!p) {
    p = new Promise((resolve) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => resolve(null);
      i.src = src;
    });
    imgCache.set(src, p);
  }
  return p;
}

// 日本語は単語区切りがないので1文字ずつ測って折り返す
function wrap(ctx, text, maxW) {
  const out = [];
  let line = '';
  for (const ch of String(text)) {
    const test = line + ch;
    if (ctx.measureText(test).width > maxW && line) {
      out.push(line);
      line = ch;
    } else {
      line = test;
    }
  }
  if (line) out.push(line);
  return out;
}

/**
 * @param ctx 2D コンテキスト
 * @param S   インチ→px の係数
 * @param PT  ポイント→px の係数
 */
export function canvasSurface(ctx, S, PT, slideW, slideH) {
  // pptx の align / valign に合わせて、文字列の描画位置を決める
  function anchorX(o) {
    if (o.align === 'center') return (o.x + o.w / 2) * S;
    if (o.align === 'right') return (o.x + o.w) * S;
    return o.x * S;
  }

  function setFont(o) {
    ctx.font = `${o.bold ? 'bold ' : ''}${o.pt * PT}px ${CANVAS_FONT}`;
    ctx.fillStyle = '#' + o.color;
    ctx.textAlign = o.align || 'left';
    ctx.textBaseline = 'top';
  }

  return {
    background(color) {
      ctx.fillStyle = '#' + color;
      ctx.fillRect(0, 0, slideW * S, slideH * S);
    },

    rect(x, y, w, h, color) {
      ctx.fillStyle = '#' + color;
      // 罫線が消えないよう、1px は必ず描く
      ctx.fillRect(x * S, y * S, Math.max(w * S, w > 0 ? 1 : 0), Math.max(h * S, h > 0 ? 1 : 0));
    },

    roundRect(x, y, w, h, r, color) {
      ctx.fillStyle = '#' + color;
      if (ctx.roundRect) {
        ctx.beginPath();
        ctx.roundRect(x * S, y * S, w * S, h * S, r * S);
        ctx.fill();
      } else {
        ctx.fillRect(x * S, y * S, w * S, h * S);   // 古いブラウザでは角丸なしで描く
      }
    },

    ellipse(x, y, w, h, color) {
      ctx.fillStyle = '#' + color;
      ctx.beginPath();
      ctx.ellipse((x + w / 2) * S, (y + h / 2) * S, (w / 2) * S, (h / 2) * S, 0, 0, Math.PI * 2);
      ctx.fill();
    },

    text(str, o) {
      const s = String(str || '');
      if (!s) return;
      setFont(o);
      if (o.charSpacing) ctx.letterSpacing = (o.charSpacing * PT) + 'px';

      const lh = o.pt * (o.lineSpacing || 1.2) * PT;
      const lines = wrap(ctx, s, o.w * S);
      const total = lines.length * lh;
      // pptx の valign と同じく、はみ出す分は下へ伸ばす
      let cy = o.valign === 'middle' ? (o.y + o.h / 2) * S - total / 2 : o.y * S;

      for (const line of lines) {
        ctx.fillText(line, anchorX(o), cy);
        cy += lh;
      }
      if (o.charSpacing) ctx.letterSpacing = '0px';
    },

    paragraphs(lines, o) {
      if (!lines.length) return;
      setFont(o);

      const fontPx = o.pt * PT;
      const lh = fontPx * (o.lineSpacing || 1.2);
      const paraGap = (o.paraSpace || 0) * PT;
      const x = o.x * S;
      const limit = (o.y + o.h) * S;
      let cy = o.y * S;

      for (const line of lines) {
        if (!line.text) { cy += lh * 0.6; continue; }   // 空行は詰めて数える

        const indent = line.bullet ? fontPx * (o.bulletIndent || 0) : 0;
        const wrapped = wrap(ctx, line.text, o.w * S - indent);

        for (let k = 0; k < wrapped.length; k++) {
          if (cy + lh > limit) return;                  // 入りきらない分は描かない
          if (line.bullet && k === 0) ctx.fillText('•', x, cy);
          ctx.fillText(wrapped[k], x + indent, cy);
          cy += lh;
        }
        cy += paraGap;
      }
    },

    async media(m, box) {
      const src = m.kind === 'video' ? (m.cover || null) : m.data;
      if (src) {
        const img = await loadImage(src);
        if (img) ctx.drawImage(img, box.x * S, box.y * S, box.w * S, box.h * S);
      } else {
        ctx.fillStyle = '#DDDDE2';
        ctx.fillRect(box.x * S, box.y * S, box.w * S, box.h * S);
      }
      if (m.kind === 'video') playBadge(ctx, box, S);
    },
  };
}

// 動画機能は現在無効だが、将来の再有効化に備えて表示処理を残す。
function playBadge(ctx, box, S) {
  const cx = (box.x + box.w / 2) * S;
  const cy = (box.y + box.h / 2) * S;
  const r = Math.max(6, Math.min(box.w, box.h) * S * 0.16);

  ctx.fillStyle = 'rgba(0,0,0,.55)';
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = '#fff';
  ctx.beginPath();
  ctx.moveTo(cx - r * 0.3, cy - r * 0.45);
  ctx.lineTo(cx + r * 0.5, cy);
  ctx.lineTo(cx - r * 0.3, cy + r * 0.45);
  ctx.closePath();
  ctx.fill();
}
