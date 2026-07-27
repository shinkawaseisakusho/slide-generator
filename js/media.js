/* 画像・動画の取り込み。
   端末の中だけで完結させるため、縮小もポスター抽出も canvas で行う。

   返す形:
     { kind: 'image', data, aspect }
     { kind: 'video', data, cover, aspect, name } */

const MAX_IMAGE_EDGE = 1400;   // 画像の長辺の上限(px)
const JPEG_QUALITY = 0.82;
export const MAX_VIDEO_BYTES = 25 * 1024 * 1024;

function videoExtension(file) {
  const fromName = String(file.name || '').match(/\.([a-z0-9]{2,5})$/i);
  if (fromName) return fromName[1].toLowerCase();

  const subtype = String(file.type || '').split('/')[1]?.toLowerCase();
  if (subtype === 'quicktime') return 'mov';
  if (subtype === 'x-m4v') return 'm4v';
  return subtype || 'mp4';
}

/** ファイルの MIME から画像／動画を判別して取り込む */
export function readMedia(file) {
  return file.type.startsWith('video/') ? readVideo(file) : readImage(file);
}

export function readImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('画像を読み込めませんでした'));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error('画像として開けませんでした'));
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

        resolve({
          kind: 'image',
          aspect: w / h,
          data: keepPng ? canvas.toDataURL('image/png') : canvas.toDataURL('image/jpeg', JPEG_QUALITY),
        });
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

export async function readVideo(file) {
  if (file.size > MAX_VIDEO_BYTES) {
    const mb = (file.size / 1024 / 1024).toFixed(0);
    throw new Error(`動画は${Math.round(MAX_VIDEO_BYTES / 1024 / 1024)}MBまでです（${mb}MB）`);
  }

  const poster = await videoPoster(file).catch(() => null);
  const data = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('動画を読み込めませんでした'));
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(file);
  });

  return {
    kind: 'video',
    name: file.name,
    extn: videoExtension(file),
    data,
    cover: poster ? poster.data : null,
    aspect: poster ? poster.aspect : 16 / 9,
  };
}

// 動画の1コマ目を取り出してポスター画像にする（失敗しても致命的ではない）
function videoPoster(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const v = document.createElement('video');
    const done = (fn) => (arg) => { URL.revokeObjectURL(url); fn(arg); };
    const fail = done(reject);
    const ok = done(resolve);

    const timer = setTimeout(() => fail(new Error('timeout')), 8000);

    v.preload = 'metadata';
    v.muted = true;
    v.playsInline = true;
    v.onerror = () => { clearTimeout(timer); fail(new Error('動画を開けませんでした')); };
    v.onloadeddata = () => {
      // 先頭は真っ黒なことが多いので少しだけ進める
      v.currentTime = Math.min(0.2, (v.duration || 1) / 2);
    };
    v.onseeked = () => {
      clearTimeout(timer);
      try {
        const scale = Math.min(1, MAX_IMAGE_EDGE / Math.max(v.videoWidth, v.videoHeight));
        const w = Math.max(1, Math.round(v.videoWidth * scale));
        const h = Math.max(1, Math.round(v.videoHeight * scale));
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        canvas.getContext('2d').drawImage(v, 0, 0, w, h);
        ok({ data: canvas.toDataURL('image/jpeg', JPEG_QUALITY), aspect: w / h });
      } catch (e) {
        fail(e);
      }
    };
    v.src = url;
  });
}

// 旧データには aspect がないので、生成前に実寸を測って補う
function measureAspect(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve(img.naturalWidth / img.naturalHeight || 4 / 3);
    img.onerror = () => resolve(4 / 3);
    img.src = dataUrl;
  });
}

export async function ensureAspects(slides) {
  const jobs = [];
  (slides || []).forEach((s) => {
    (s.media || []).forEach((m) => {
      if (m && m.data && !m.aspect) {
        jobs.push(measureAspect(m.cover || m.data).then((a) => { m.aspect = a; }));
      }
    });
  });
  if (jobs.length) await Promise.all(jobs);
}
