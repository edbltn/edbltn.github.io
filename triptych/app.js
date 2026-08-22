/* Triptych — a PDF presenter with two kinds of transition.
 *
 *   rotation    the panes move across the triptych: next slide swings into the
 *               centre, the centre shrinks into the left pane.
 *   within-slide the centre pane advances one build step; nothing moves.
 *
 * Which junction is which is inferred from the deck itself (see the grouping
 * section below) and can be overridden by hand in the structure inspector.
 *
 * A PDF export drops the deck's video. Hand it the .pptx as well and the video
 * comes back: pptx.js pulls the media parts and their placement out of the zip,
 * and they are laid over the matching page (see attachPptx).
 */

import * as pdfjsLib from './vendor/pdf.min.mjs';
import { readPptx } from './pptx.js';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  new URL('./vendor/pdf.worker.min.mjs', import.meta.url).toString();

/* ═══════════════════════════ settings ═══════════════════════════ */

const DEFAULTS = {
  fracs: [0.21, 0.58, 0.21],
  sideScale: 1,
  dim: 0.5,
  fit: 'fill',
  foldDeg: 46,
  depth: 1.0,
  rotateMs: 620,
  stepMs: 220,
  stepStyle: 'crossfade',
  gutter: 10,
  bg: '#0d0e10',
  videoAuto: true,
  videoLoop: false,
  webEmbed: '',
  webEmbedBlank: true,
  embedZoom: 1.5,
  embedPreload: 2,
  threshold: 0.88,
};

const SETTINGS_KEY = 'triptych:settings';
const SETTINGS_VERSION = 3;

/* Saved settings shadow the defaults, so a default that changes later would
 * never reach anyone who has used the app before. Bumping the version keeps
 * the one thing worth carrying over — the pane widths you tuned — and takes
 * the new defaults for everything else. */
const stored = readJSON(SETTINGS_KEY) || {};
const carried = stored.version === SETTINGS_VERSION ? stored : { fracs: stored.fracs };
const settings = Object.assign({}, DEFAULTS, carried);
for (const key of Object.keys(settings)) {
  if (!(key in DEFAULTS)) delete settings[key];
}
settings.version = SETTINGS_VERSION;
settings.fracs = normalizeFracs(settings.fracs);

function readJSON(k) { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } }
function writeJSON(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* quota */ } }
const saveSettings = debounce(() => writeJSON(SETTINGS_KEY, settings), 250);

/* ═══════════════════════════ deck state ═══════════════════════════ */

/** @type {{url:string, thumb:string, hash:Uint8Array, tokens:Set<string>, title:string, w:number, h:number}[]} */
let pages = [];
/** junction k sits between page k and page k+1 */
let junctions = [];
/** @type {number[][]} slides[i] = array of page indices */
let slides = [];
/** @type {Record<number,'build'|'rotate'>} */
let overrides = {};
let deckKey = null;
let deckName = '';

/** page index → media laid over that page, in slide-relative fractions */
let mediaByPage = [];
let pptxNote = '';
let soundOn = true;

let slideIdx = 0;
let stepIdx = 0;

/** @type {null | {kind:'rotate', dir:number, t0:number, dur:number}
 *              | {kind:'step', from:string, dir:number, t0:number, dur:number}} */
let anim = null;
let dirty = true;

/* ═══════════════════════════ dom ═══════════════════════════ */

const $ = (id) => document.getElementById(id);
const stage = $('stage');
const panesEl = $('panes');
const embedLayer = $('embeds');
const dropzone = $('dropzone');

/* ═══════════════════════════ pdf ingest ═══════════════════════════ */

/** ?debug=1 logs per-page ingest timings to the console. */
const DEBUG = new URLSearchParams(location.search).has('debug');
const mark = () => performance.now();
const since = (t) => Math.round(performance.now() - t);

const RENDER_WIDTH = Math.max(1400, Math.min(2400,
  Math.round(window.screen.width * (window.devicePixelRatio || 1) * 1.15)));
const HASH_N = 48;          // hash grid is HASH_N × HASH_N
const THUMB_W = 200;

async function loadPdf(source, name) {
  deckName = name || 'deck.pdf';
  dropzone.classList.add('hidden');
  $('loading').classList.remove('hidden');
  $('loadTitle').textContent = 'Reading ' + deckName;
  setLoadProgress(0, 0);

  const doc = await pdfjsLib.getDocument(source).promise;
  const n = doc.numPages;
  const out = [];

  const hashCanvas = document.createElement('canvas');
  hashCanvas.width = hashCanvas.height = HASH_N;
  const hashCtx = hashCanvas.getContext('2d', { willReadFrequently: true });

  for (let i = 1; i <= n; i++) {
    $('loadTitle').textContent = 'Rendering pages…';
    setLoadProgress(i - 1, n);

    const t = mark();
    const page = await doc.getPage(i);
    const tGot = mark();
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(RENDER_WIDTH / base.width, 3.5);
    const vp = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    canvas.width = Math.round(vp.width);
    canvas.height = Math.round(vp.height);
    const ctx = canvas.getContext('2d', { alpha: false });
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // intent:'print' is not about printing here — it is the one render path
    // pdf.js paces on microtasks instead of requestAnimationFrame. With the
    // default 'display' intent, ingest stalls completely whenever the tab is
    // hidden (rAF never fires) and crawls even when it isn't. Same raster,
    // ~20-200ms a page instead of tens of seconds.
    await page.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
    const tRendered = mark();

    // perceptual hash off the same raster
    hashCtx.fillStyle = '#ffffff';
    hashCtx.fillRect(0, 0, HASH_N, HASH_N);
    hashCtx.drawImage(canvas, 0, 0, HASH_N, HASH_N);
    const px = hashCtx.getImageData(0, 0, HASH_N, HASH_N).data;
    const hash = new Uint8Array(HASH_N * HASH_N);
    for (let p = 0, q = 0; p < px.length; p += 4, q++) {
      hash[q] = (px[p] * 77 + px[p + 1] * 151 + px[p + 2] * 28) >> 8;
    }

    const tHashed = mark();
    const text = await page.getTextContent();
    const { tokens, title } = digestText(text);
    const tText = mark();

    const url = await encodeToUrl(canvas);
    const thumb = await encodeToUrl(downscale(canvas, THUMB_W), 0.72);
    if (DEBUG) {
      console.log(`[ingest] p${i} ${canvas.width}×${canvas.height} ` +
        `getPage ${since(t) - since(tGot)}ms · render ${since(tGot) - since(tRendered)}ms · ` +
        `hash ${since(tRendered) - since(tHashed)}ms · text ${since(tHashed) - since(tText)}ms · ` +
        `encode ${since(tText)}ms`);
    }

    out.push({ url, thumb, hash, tokens, title, w: canvas.width, h: canvas.height });
    page.cleanup();
    await breathe();
  }

  setLoadProgress(n, n);
  pages.forEach(p => { URL.revokeObjectURL(p.url); URL.revokeObjectURL(p.thumb); });
  pages = out;

  deckKey = 'triptych:deck:' + hashString(deckName + ':' + n + ':' +
    pages.map(p => p.tokens.size).join(','));
  overrides = readJSON(deckKey) || {};

  for (const frame of embedFrames.values()) frame.remove();
  embedFrames.clear();

  junctions = [];
  for (let k = 0; k + 1 < pages.length; k++) junctions.push(similarity(pages[k], pages[k + 1]));

  regroup(0);
  $('loading').classList.add('hidden');
  document.title = deckName + ' — Triptych';
  dirty = true;
  buildInspector();
}

function setLoadProgress(i, n) {
  $('loadFill').style.width = n ? (100 * i / n) + '%' : '0%';
  $('loadSub').textContent = n ? `${i} / ${n}` : '';
}

function downscale(canvas, width) {
  const c = document.createElement('canvas');
  c.width = width;
  c.height = Math.round(canvas.height * width / canvas.width);
  const x = c.getContext('2d', { alpha: false });
  x.imageSmoothingQuality = 'high';
  x.drawImage(canvas, 0, 0, c.width, c.height);
  return c;
}

/** WebP where available (crisper text per byte), JPEG otherwise. */
let encodeType = null;
function encodeToUrl(canvas, quality = 0.92) {
  return new Promise((resolve) => {
    const attempt = (type) => canvas.toBlob((blob) => {
      if (!blob) return resolve('');
      if (encodeType === null) encodeType = blob.type === 'image/webp' ? 'image/webp' : 'image/jpeg';
      resolve(URL.createObjectURL(blob));
    }, type, quality);
    attempt(encodeType || 'image/webp');
  });
}

/* ═══════════════════════════ text digest ═══════════════════════════ */

function digestText(content) {
  const items = content.items.filter(it => it.str && it.str.trim());
  const tokens = new Set();
  for (const it of items) {
    for (const t of it.str.toLowerCase().match(/[a-z0-9][a-z0-9'’\-]*/g) || []) {
      if (t.length > 1) tokens.add(t);
    }
  }
  // topmost line ≈ the slide's title; a shared title is strong evidence of a build
  let title = '';
  if (items.length) {
    const withY = items.map(it => ({ y: it.transform[5], x: it.transform[4], s: it.str }));
    const topY = Math.max(...withY.map(v => v.y));
    title = withY.filter(v => Math.abs(v.y - topY) < 4)
      .sort((a, b) => a.x - b.x).map(v => v.s).join(' ')
      .toLowerCase().replace(/\s+/g, ' ').trim();
  }
  return { tokens, title };
}

/* ═══════════════════════════ grouping ═══════════════════════════ */

/** How alike are two consecutive pages? Returns the evidence, not the verdict. */
function similarity(a, b) {
  let changed = 0;
  for (let i = 0; i < a.hash.length; i++) if (Math.abs(a.hash[i] - b.hash[i]) > 22) changed++;
  const visual = 1 - changed / a.hash.length;

  const A = a.tokens, B = b.tokens;
  let inter = 0;
  for (const t of A) if (B.has(t)) inter++;
  const union = A.size + B.size - inter;
  const jaccard = union ? inter / union : 1;

  const small = A.size <= B.size ? A : B;
  const subset = small.size >= 3 && inter === small.size && A.size !== B.size;
  const titleSame = !!a.title && a.title === b.title;

  const hasText = A.size >= 3 && B.size >= 3;
  const textSim = subset ? Math.max(jaccard, 0.92) : jaccard;

  const score = hasText
    ? Math.min(1, 0.55 * visual + 0.45 * textSim + (titleSame ? 0.03 : 0))
    : visual;

  return { visual, jaccard, subset, titleSame, score };
}

function autoIsBuild(k) { return junctions[k].score >= settings.threshold; }
function isBuild(k) {
  const o = overrides[k];
  return o ? o === 'build' : autoIsBuild(k);
}

/** Rebuild slides[] from the junctions, keeping `keepPage` on screen. */
function regroup(keepPage) {
  if (keepPage === undefined) keepPage = currentPage();
  slides = [];
  let cur = [0];
  for (let k = 0; k < junctions.length; k++) {
    if (isBuild(k)) cur.push(k + 1);
    else { slides.push(cur); cur = [k + 1]; }
  }
  if (pages.length) slides.push(cur);

  pageToSlide = [];
  slides.forEach((group, i) => group.forEach(page => { pageToSlide[page] = i; }));

  outer: for (let i = 0; i < slides.length; i++) {
    for (let j = 0; j < slides[i].length; j++) {
      if (slides[i][j] === keepPage) { slideIdx = i; stepIdx = j; break outer; }
    }
  }
  slideIdx = clamp(slideIdx, 0, Math.max(0, slides.length - 1));
  stepIdx = clamp(stepIdx, 0, Math.max(0, (slides[slideIdx] || []).length - 1));
  anim = null;
  dirty = true;
}

const currentPage = () => (slides[slideIdx] || [0])[stepIdx] ?? 0;

/* ═══════════════════════════ geometry ═══════════════════════════ */

const SHUT_TURN = 22;     // extra degrees a panel turns as it closes on its way out

/* Anchor rects for the slots a pane can occupy. Slot 0 is the centre;
 * ±1 are the wings; ±2 sit just off-stage, which is what makes a rotation
 * read as panes sliding across the triptych rather than a crossfade. */
function anchor(slot, W) {
  const [fl, fc, fr] = settings.fracs;
  if (slot <= -1) return { x: 0, w: fl * W };
  if (slot >= 1) return { x: (fl + fc) * W, w: fr * W };
  return { x: fl * W, w: fc * W };
}

/* The wings fold *toward* the viewer, the way an altarpiece opens out, so the
 * outer edge is the near one — wider and taller than the seam.
 *
 * How wide must a panel be, unfolded, for its outer edge to project exactly
 * onto `outer`? A point d along the panel sits at z = +d·sin θ, so its screen x
 * is cx + (hinge + side·d·cos θ − cx) · P/(P − d·sin θ). That grows without
 * bound as the panel reaches the camera plane, so bisect below it. */
function flapWidth(hinge, outer, side, theta, W, P) {
  const cx = W / 2;
  const cos = Math.cos(theta), sin = Math.sin(theta);
  const at = (d) => cx + (hinge + side * d * cos - cx) * P / (P - d * sin);
  let lo = 0, hi = Math.min(W * 6, sin > 1e-6 ? 0.82 * P / sin : W * 6);
  for (let i = 0; i < 28; i++) {
    const mid = (lo + hi) / 2;
    const short = side < 0 ? at(mid) > outer : at(mid) < outer;
    if (short) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

function geomAt(s, W) {
  // Beyond a wing a panel stops travelling: it stays in the wing's section and
  // keeps folding instead, so it swings shut on its hinge rather than sliding
  // off the stage.
  const c = clamp(s, -1, 1);
  const lo = Math.floor(c), hi = Math.ceil(c), t = c - lo;
  const a = anchor(lo, W), b = anchor(hi, W);
  return { x: lerp(a.x, b.x, t), w: lerp(a.w, b.w, t) };
}

/* ═══════════════════════════ pane pool ═══════════════════════════ */

const paneFor = new Map();   // slide index → element

function acquirePane(j) {
  let el = paneFor.get(j);
  if (!el) {
    el = document.createElement('div');
    el.className = 'pane';
    el.innerHTML = '<div class="pane-inner">' +
      '<img class="under" alt=""><img class="over" alt=""><div class="shade"></div></div>';
    el.addEventListener('click', () => {
      if (j < slideIdx) prevSlide();
      else if (j > slideIdx) nextSlide();
      else next();
    });
    panesEl.appendChild(el);
    paneFor.set(j, el);
  }
  return el;
}

function setLayer(img, url) {
  if (img.dataset.url !== url) { img.dataset.url = url; img.src = url || ''; }
}

/** Which page of slide j should be showing right now. */
function pageOfSlide(j) {
  const s = slides[j];
  if (!s) return null;
  if (j < slideIdx) return s[s.length - 1];
  if (j > slideIdx) return s[0];
  return s[clamp(stepIdx, 0, s.length - 1)];
}

function render() {
  if (!pages.length) return;
  const W = stage.clientWidth, H = stage.clientHeight;
  const g = settings.gutter;

  let offset = 0, blend = 0, underUrl = '', lift = 0;
  if (anim) {
    const t = clamp((performance.now() - anim.t0) / anim.dur, 0, 1);
    const e = easeInOutCubic(t);
    if (anim.kind === 'rotate') offset = anim.dir * e;
    else {
      underUrl = anim.from;
      blend = settings.stepStyle === 'cut' ? (t < 1 ? 0 : 1) : e;
      if (settings.stepStyle === 'lift') lift = (1 - e) * 14 * anim.dir;
    }
    if (t >= 1) { commitAnim(); return render(); }
  }

  panesEl.style.perspective = (settings.depth * W).toFixed(0) + 'px';

  let centreBox = null;
  const live = new Set();
  for (let j = slideIdx - 2; j <= slideIdx + 2; j++) {
    if (j < 0 || j >= slides.length) continue;
    const s = (j - slideIdx) - offset;
    if (Math.abs(s) > 2) continue;
    live.add(j);

    const el = acquirePane(j);
    const { x, w } = geomAt(s, W);
    const k = clamp(Math.abs(s), 0, 1);
    const sectionX = x + g / 2;
    const sectionW = Math.max(0, w - g);

    /* The wings are hinged flaps: each is rotated about the edge it shares
     * with the centre, so perspective turns it into a trapezoid. `flapWidth`
     * works out how wide the *unfolded* panel has to be for its far edge to
     * land exactly on the section boundary — otherwise folding one back would
     * pull it away from the divider you placed. */
    const side = s < 0 ? -1 : 1;
    const hinge = s < 0 ? sectionX + sectionW : sectionX;
    const outer = s < 0 ? sectionX : sectionX + sectionW;

    /* Up to the wing, the panel is sized to fill its section at the fold angle.
     * Past it, the panel leaves by closing against its *outer* edge: that edge
     * stays where it is — the near, tall one — while everything sweeps outward
     * into it and the panel narrows to nothing. The squeeze is applied to the
     * contents (below) rather than the box, so the fold it is already sitting
     * in carries through the exit. */
    const shut = clamp(Math.abs(s) - 1, 0, 1);
    const fillRad = (settings.foldDeg * Math.PI / 180) * k;
    const deg = settings.foldDeg * k + SHUT_TURN * shut;
    const paneW = fillRad > 1e-4
      ? flapWidth(hinge, outer, side, fillRad, W, settings.depth * W)
      : sectionW;

    /* An unfolded pane stays a plain 2D layer: Chrome will not composite an
     * iframe (a YouTube player, say) inside a 3D-transformed ancestor, and the
     * centre pane is exactly where those live. */
    const turn = -side * deg;
    const tx = (s < 0 ? hinge - paneW : hinge).toFixed(2);
    el.style.transform = Math.abs(turn) < 0.01
      ? `translate(${tx}px,0)`
      : `translate3d(${tx}px,0,0) rotateY(${turn.toFixed(3)}deg)`;
    el.style.transformOrigin = s < 0 ? '100% 50%' : '0% 50%';
    // the last sliver of an almost-edge-on panel is a hairline; fade it out
    el.style.opacity = Math.abs(s) <= 1.5 ? '1'
      : Math.max(0, 1 - (Math.abs(s) - 1.5) / 0.5).toFixed(3);
    el.style.width = paneW + 'px';
    el.style.height = H + 'px';
    el.classList.toggle('clickable', j !== slideIdx || hasNext());

    const inner = el.firstElementChild;
    const isCentre = j === slideIdx;
    // the outer edge is the anchor: for a left wing that is its left side
    inner.style.transformOrigin = s < 0 ? '0% 50%' : '100% 50%';
    inner.style.transform =
      `scaleX(${(1 - shut).toFixed(4)}) translateY(${isCentre ? lift.toFixed(2) : 0}px)`;

    /* Light falls off toward the far edge of a folded panel — a flat dim
     * would read as a faded thumbnail rather than a panel turned away. */
    const shade = inner.querySelector('.shade');
    shade.style.opacity = (settings.dim * k).toFixed(3);
    shade.style.background =
      `linear-gradient(to ${s < 0 ? 'right' : 'left'}, rgba(0,0,0,0), rgba(0,0,0,1))`;

    const p = pages[pageOfSlide(j)];

    /* In the centre the whole slide is visible; out on a wing it keeps the
     * centre's height and lets the divider cut it off, anchored to the edge it
     * shares with the centre. `k` carries it continuously between the two, so
     * a rotation grows a clipped wing into a whole slide. */
    const aspect = p.w / p.h;
    const shrink = lerp(1, settings.sideScale, k);
    const centreH = Math.min(H, Math.max(1, anchor(0, W).w - g) / aspect);

    /* 'fill': the slide covers the whole flap, at the centre's height — so the
     *   two meet at the same height along the hinge and the fold reads as one
     *   continuous surface. Narrow flaps squeeze the slide horizontally; that
     *   distortion is the price of a seam that lines up.
     * 'contain': the flap holds the whole slide undistorted instead, floating
     *   inside its section. */
    const fill = settings.fit === 'fill';
    const imgH = (fill ? centreH : Math.min(H, paneW / aspect)) * shrink;
    const imgW = fill ? lerp(centreH * aspect, paneW, k) : imgH * aspect;
    const ax = lerp(0.5, s < 0 ? 1 : 0, fill ? k : 0);   // 0 = flush left, 1 = flush right
    const box = { left: (paneW - imgW) * ax, top: (H - imgH) / 2, w: imgW, h: imgH };
    const imgStyle =
      `width:${imgW.toFixed(1)}px;height:${imgH.toFixed(1)}px;` +
      `left:${box.left.toFixed(1)}px;top:${box.top.toFixed(1)}px;`;

    const over = inner.querySelector('.over');
    const under = inner.querySelector('.under');
    over.style.cssText = imgStyle;
    under.style.cssText = imgStyle;
    setLayer(over, p.url);

    syncMedia(inner, pageOfSlide(j), box, isCentre && !(anim && anim.kind === 'rotate'));

    // the embed layer sits outside the panes, so it needs stage coordinates
    if (isCentre) {
      centreBox = { ...box, left: box.left + (s < 0 ? hinge - paneW : hinge) };
    }
    if (isCentre && underUrl) {
      setLayer(under, underUrl);
      under.style.opacity = '1';
      over.style.opacity = String(blend);
    } else {
      under.style.opacity = '0';
      over.style.opacity = '1';
    }
  }

  for (const [j, el] of paneFor) {
    if (!live.has(j)) { el.remove(); paneFor.delete(j); }
  }

  positionDividers(W);
  if (centreBox) syncEmbedLayer(centreBox);

  const sig = `${slideIdx}:${stepIdx}:${live.size}:${settings.videoAuto}:${soundOn}`;
  if (sig !== playbackSig) { playbackSig = sig; syncPlayback(); }
}

function positionDividers(W) {
  const [fl, fc] = settings.fracs;
  $('divider0').style.left = (fl * W) + 'px';
  $('divider1').style.left = ((fl + fc) * W) + 'px';
}

function frame() {
  if (dirty || anim) { dirty = false; render(); }
  requestAnimationFrame(frame);
}
requestAnimationFrame(frame);

/* ═══════════════════════════ media ═══════════════════════════ */

/* One <video> per media item per pane, created on demand and re-pointed when
 * the pane changes page. Geometry comes from the slide image's box, so the
 * video travels with the slide through a rotation and gets clipped by the
 * divider exactly like the picture under it. */
function syncMedia(inner, pageIdx, box, live) {
  // anything in an iframe is hoisted to its own layer, to load ahead of time
  const want = (mediaByPage[pageIdx] || []).filter(m => !HOSTED.has(m.kind));
  let holder = inner.querySelector('.vids');

  if (!want.length) { if (holder) holder.remove(); return; }
  if (!holder) {
    holder = document.createElement('div');
    holder.className = 'vids';
    inner.appendChild(holder);
  }

  // A hosted player only exists while its slide is the centre one — removing
  // the iframe is the only reliable way to stop YouTube playing.
  const hosted = want.some(m => m.kind === 'youtube' || m.kind === 'vimeo');
  const stamp = `${pageIdx}|${live ? 1 : 0}|${hosted ? (soundOn ? 1 : 0) : 0}|${settings.videoLoop ? 1 : 0}`;

  if (holder.dataset.stamp !== stamp) {
    holder.dataset.stamp = stamp;
    holder.dataset.page = String(pageIdx);
    holder.textContent = '';
    for (const m of want) holder.appendChild(makeMediaEl(m, live));
    playbackSig = '';
  }

  [...holder.children].forEach((el, i) => {
    const m = want[i];
    const left = box.left + m.x * box.w;
    const top = box.top + m.y * box.h;
    let w = m.w * box.w;
    let h = m.h * box.h;
    let extra = '';

    el.style.cssText =
      `left:${left.toFixed(1)}px;top:${top.toFixed(1)}px;` +
      `width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;` + extra +
      (m.kind === 'audio' ? 'opacity:0;pointer-events:none;' : '');
  });
}

function makeMediaEl(m, live) {
  const v = document.createElement('video');
  v.src = m.url;
  v.playsInline = true;
  v.preload = 'auto';
  v.loop = settings.videoLoop;
  v.dataset.kind = m.kind;
  return v;
}

/* ───────── the embed layer ───────── */

/* A live page is not drawn inside a pane. Two reasons: a pane is recycled as
 * you move, which would tear the page down and reload it on every pass; and a
 * folded pane is 3D-transformed, which Chrome refuses to composite an iframe
 * inside. So embeds live in their own flat layer over the stage, load while
 * their slide is still a couple of slides away, and are simply revealed when
 * you arrive. */
const embedFrames = new Map();     // page index → iframe

/* Everything that is an iframe: a live page, and the hosted players — which
 * belong here too, so they load ahead instead of on arrival. */
const HOSTED = new Set(['site', 'youtube', 'vimeo']);
const siteOn = (pageIdx) => (mediaByPage[pageIdx] || []).find(m => HOSTED.has(m.kind));

function frameSrc(m) {
  if (m.kind === 'site') return m.url;

  /* autoplay is off on purpose: the frame is created while its slide is still
   * a couple away, and is told to play only once you arrive. */
  const sound = soundOn && (navigator.userActivation?.hasBeenActive ?? true);
  if (m.kind === 'youtube') {
    return `https://www.youtube-nocookie.com/embed/${m.url}?` + new URLSearchParams({
      autoplay: '0', mute: sound ? '0' : '1', enablejsapi: '1',
      controls: '0', disablekb: '1', fs: '0', rel: '0',
      modestbranding: '1', iv_load_policy: '3', playsinline: '1',
      ...(settings.videoLoop ? { loop: '1', playlist: m.url } : {}),
    });
  }
  return `https://player.vimeo.com/video/${m.url}?` + new URLSearchParams({
    autoplay: '0', muted: sound ? '0' : '1', title: '0', byline: '0',
    portrait: '0', controls: '0', ...(settings.videoLoop ? { loop: '1' } : {}),
  });
}

/** Drive a hosted player over postMessage — no third-party script needed. */
function command(frame, what) {
  const target = frame.contentWindow;
  if (!target) return;
  if (frame.dataset.kind === 'vimeo') {
    target.postMessage(JSON.stringify({ method: what === 'play' ? 'play' : 'pause' }), '*');
    return;
  }
  target.postMessage(JSON.stringify({
    event: 'command', func: what === 'play' ? 'playVideo' : 'pauseVideo', args: [],
  }), '*');
}

/* A player ignores commands until it is ready, and there is no reliable ready
 * signal without loading its API script — so ask a few times over a second. */
function startHosted(frame) {
  let tries = 0;
  const nudge = () => {
    if (frame.dataset.playing !== '1' || !frame.isConnected) return;
    command(frame, 'play');
    if (++tries < 8) setTimeout(nudge, 150);
  };
  nudge();
}

function syncEmbedLayer(box) {
  const radius = clamp(Number(settings.embedPreload) || 0, 0, 99);
  const current = currentPage();
  const rotating = !!(anim && anim.kind === 'rotate');

  for (let page = 0; page < mediaByPage.length; page++) {
    const site = siteOn(page);
    if (!site) continue;

    const near = Math.abs(slideOfPage(page) - slideIdx) <= radius || page === current;
    let frame = embedFrames.get(page);

    if (near && !frame) {
      frame = document.createElement('iframe');
      frame.className = site.kind === 'site' ? 'site' : 'player';
      frame.dataset.kind = site.kind;
      frame.src = frameSrc(site);
      frame.allow = 'autoplay; encrypted-media; clipboard-write; picture-in-picture';
      frame.setAttribute('frameborder', '0');
      frame.setAttribute('referrerpolicy', 'strict-origin-when-cross-origin');
      embedLayer.appendChild(frame);
      embedFrames.set(page, frame);
    } else if (!near && frame) {
      // out of range: drop it, so a page that counts its viewers stops counting
      frame.remove();
      embedFrames.delete(page);
      continue;
    }
    if (!frame) continue;

    const src = frameSrc(site);
    if (frame.dataset.url !== src) { frame.dataset.url = src; frame.src = src; }

    /* Every frame is laid out at the size it will have in the centre, whether
     * or not it is showing — a page that loads at 0×0 reflows the moment you
     * reveal it, which is the flash we are trying to avoid. */
    const zoom = site.kind === 'site' ? clamp(Number(settings.embedZoom) || 1, 1, 4) : 1;
    const live = page === current && !rotating;

    // a player is loaded well before its slide but must not start early
    if (site.kind !== 'site' && live !== (frame.dataset.playing === '1')) {
      frame.dataset.playing = live ? '1' : '0';
      if (live && settings.videoAuto) startHosted(frame);
      else command(frame, 'pause');
    }
    frame.style.cssText =
      `left:${(box.left + site.x * box.w).toFixed(1)}px;` +
      `top:${(box.top + site.y * box.h).toFixed(1)}px;` +
      `width:${(site.w * box.w * zoom).toFixed(1)}px;` +
      `height:${(site.h * box.h * zoom).toFixed(1)}px;` +
      `transform:scale(${(1 / zoom).toFixed(4)});transform-origin:0 0;` +
      (live ? '' : 'visibility:hidden;opacity:0;pointer-events:none;');
  }
}

/** Which slide a page belongs to — rebuilt with the grouping. */
let pageToSlide = [];
const slideOfPage = (page) => pageToSlide[page] ?? -1;

let playbackSig = '';

/** Only the centre pane's current page plays; everything else rewinds. */
function syncPlayback() {
  const p = currentPage();
  for (const [j, el] of paneFor) {
    const holder = el.querySelector('.vids');
    if (!holder) continue;
    const live = j === slideIdx && Number(holder.dataset.page) === p;
    for (const v of holder.children) {
      if (v.tagName !== 'VIDEO') continue;   // hosted players live and die with the pane
      v.loop = settings.videoLoop;
      if (live && settings.videoAuto) startVideo(v);
      else {
        v.pause();
        if (!live) { try { v.currentTime = 0; } catch { /* not seekable yet */ } }
      }
    }
  }
}

/** Sound needs a user gesture. Presenting is all gestures, so try with sound
 *  and fall back to muted rather than not playing at all. */
async function startVideo(v) {
  v.muted = !soundOn;
  if (!v.paused) return;
  try {
    await v.play();
  } catch {
    v.muted = true;
    try { await v.play(); } catch { /* codec or source problem; leave the poster */ }
  }
}

function toggleSound() {
  soundOn = !soundOn;
  for (const el of panesEl.querySelectorAll('video')) el.muted = !soundOn;
  // hosted players carry mute in their URL; drop them and let them rebuild
  for (const f of panesEl.querySelectorAll('.vids')) f.dataset.stamp = '';
  for (const [page, frame] of embedFrames) {
    if (frame.dataset.kind !== 'site') { frame.remove(); embedFrames.delete(page); }
  }
  dirty = true;
}

/* Slides left deliberately empty are placeholders for a live page. The deck
 * says which are empty; this drops the configured site onto them. A slide with
 * `embed: <url>` in its speaker notes has already been handled and is left be. */
let lastBlanks = null;

let mediaCacheBase = '/';

/** embeds.json, if it is there: per-slide pages and a default for blanks. */
let embedConfig = { blank: '', slides: {} };

async function loadEmbedConfig() {
  try {
    const res = await fetch('/embeds.json', { cache: 'no-store' });
    if (!res.ok) return;
    const cfg = await res.json();
    embedConfig = {
      blank: typeof cfg.blank === 'string' ? cfg.blank : '',
      slides: cfg.slides && typeof cfg.slides === 'object' ? cfg.slides : {},
    };
  } catch (err) {
    console.warn('[triptych] embeds.json is not valid JSON —', err.message);
  }
}

/* Where a live page comes from, most specific first:
 *   1. `embed: <url>` in the slide's own speaker notes (handled in pptx.js)
 *   2. that slide's number in embeds.json
 *   3. the blank-slide default — embeds.json, else the settings field  */
function fillBlankSlides(blanks) {
  lastBlanks = blanks || lastBlanks;
  blanks = lastBlanks;
  if (!blanks) return;

  const fallback = settings.webEmbedBlank
    ? (embedConfig.blank || settings.webEmbed || '').trim()
    : '';
  const inset = { x: 0.03, y: 0.04, w: 0.94, h: 0.92 };

  blanks.forEach((isBlank, i) => {
    const url = (embedConfig.slides[String(i + 1)] || (isBlank ? fallback : '')).trim();
    if (!url) return;
    const page = pageForSlide(i);
    if (page === null || mediaByPage[page].some(m => m.kind === 'site')) return;
    mediaByPage[page].push({ url, kind: 'site', name: url, fromBlank: true, ...inset });
  });
}

/** Which PDF page a PowerPoint slide index landed on. */
let slideToPage = null;
const pageForSlide = (i) => (slideToPage ? slideToPage[i] ?? null : null);

/* A hosted player brings its own chrome and its own rules. If serve.py has a
 * local copy of the clip (see "cache videos" in settings), play that instead:
 * a plain <video>, no branding, no controls, no third party. */
async function useCachedVideos() {
  const checks = [];
  for (const list of mediaByPage) {
    for (const m of list) {
      if (m.kind !== 'youtube' && m.kind !== 'vimeo') continue;
      const path = `${mediaCacheBase}media/${m.kind}-${m.url}.mp4`;
      checks.push(fetch(path, { method: 'HEAD' }).then(r => {
        if (r.ok) { m.kind = 'video'; m.url = path; }
      }).catch(() => { /* no cache, keep the embed */ }));
    }
  }
  await Promise.all(checks);
}

/** Ask serve.py to pull the clips down with yt-dlp, then switch to them. */
async function cacheVideos(button) {
  const wanted = mediaByPage.flat().filter(m => m.kind === 'youtube' || m.kind === 'vimeo');
  if (!wanted.length) { button.textContent = 'no hosted clips in this deck'; return; }

  const label = button.textContent;
  for (const [i, m] of wanted.entries()) {
    button.textContent = `fetching ${i + 1} / ${wanted.length}…`;
    try {
      const r = await fetch(`/fetch-video?host=${m.kind}&id=${encodeURIComponent(m.url)}`);
      if (!r.ok) throw new Error(await r.text());
    } catch (err) {
      button.textContent = String(err.message || err).slice(0, 60);
      return;
    }
  }
  fillBlankSlides(deck.blanks);
  await useCachedVideos();
  playbackSig = '';
  for (const h of panesEl.querySelectorAll('.vids')) h.dataset.stamp = '';
  dirty = true;
  button.textContent = `cached ${wanted.length} clip${wanted.length === 1 ? '' : 's'}`;
  setTimeout(() => { button.textContent = label; }, 4000);
}

/* ───────── pptx: the video a PDF export leaves behind ───────── */

async function attachPptx(buffer, name) {
  await configReady;
  return attachDeck(await readPptx(buffer), name);
}

/* Both routes end here: a .pptx parsed in the browser, and a talk.json built
 * by publish.py (same shape, so a published talk ships no .pptx at all). */
async function attachDeck(deck, name) {
  const byIndex = deck.media;
  mediaByPage = pages.map(() => []);

  if (deck.slideCount === pages.length) {
    // one PDF page per PowerPoint slide
    slideToPage = byIndex.map((_, i) => i);
  } else if (deck.slideCount === slides.length) {
    // the PDF exported build steps as extra pages; media belongs to the slide
    slideToPage = byIndex.map((_, i) => slides[i][0]);
  } else {
    // no honest alignment — line them up from the front and say so
    slideToPage = byIndex.map((_, i) => (i < pages.length ? i : null));
    pptxNote = `${name}: ${deck.slideCount} slides vs ${pages.length} PDF pages ` +
      `(${slides.length} after grouping) — media aligned from the first slide, ` +
      `which may be off.`;
    console.warn('[triptych] ' + pptxNote);
  }
  slideToPage.forEach((page, i) => { if (page !== null) mediaByPage[page] = byIndex[i]; });

  fillBlankSlides(deck.blanks);
  await useCachedVideos();

  const total = mediaByPage.reduce((n, m) => n + m.length, 0);
  console.log(`[triptych] ${name}: ${total} media item${total === 1 ? '' : 's'} ` +
    `across ${deck.slideCount} slides`);
  playbackSig = '';
  dirty = true;
  buildInspector();       // the summary counts video, which only exists now
  if (DEBUG) {
    window.__triptych = {
      get media() { return mediaByPage.map((m, i) => ({ page: i + 1, items: m })); },
      get slides() { return slides; },
      get config() { return embedConfig; },
    };
  }
  return total;
}

/* ═══════════════════════════ navigation ═══════════════════════════ */

const stepsIn = (j) => (slides[j] || []).length;
const hasNext = () => stepIdx < stepsIn(slideIdx) - 1 || slideIdx < slides.length - 1;

function commitAnim() {
  if (!anim) return;
  if (anim.kind === 'rotate') {
    slideIdx = clamp(slideIdx + anim.dir, 0, slides.length - 1);
    stepIdx = anim.dir > 0 ? 0 : stepsIn(slideIdx) - 1;
  }
  anim = null;
  markCurrentInInspector();
}

function settle() { if (anim) commitAnim(); }

function startRotate(dir) {
  settle();
  const target = slideIdx + dir;
  if (target < 0 || target >= slides.length) return bump(dir);
  anim = { kind: 'rotate', dir, t0: performance.now(), dur: settings.rotateMs };
  dirty = true;
}

function startStep(dir) {
  settle();
  const from = pages[currentPage()].url;
  stepIdx += dir;
  if (settings.stepMs <= 0) { anim = null; markCurrentInInspector(); dirty = true; return; }
  anim = { kind: 'step', from, dir, t0: performance.now(), dur: settings.stepMs };
  markCurrentInInspector();
  dirty = true;
}

function next() {
  settle();
  if (stepIdx < stepsIn(slideIdx) - 1) startStep(1);
  else startRotate(1);
}

function prev() {
  settle();
  if (stepIdx > 0) startStep(-1);
  else startRotate(-1);
}

const nextSlide = () => startRotate(1);
const prevSlide = () => startRotate(-1);

function goToPage(p) {
  settle();
  for (let i = 0; i < slides.length; i++) {
    const j = slides[i].indexOf(p);
    if (j >= 0) { slideIdx = i; stepIdx = j; break; }
  }
  markCurrentInInspector();
  dirty = true;
}

/** Nudge the stage when you run off either end of the deck. */
function bump(dir) {
  panesEl.animate(
    [{ transform: 'translateX(0)' }, { transform: `translateX(${-dir * 10}px)` }, { transform: 'translateX(0)' }],
    { duration: 220, easing: 'ease-out' });
}

/* ═══════════════════════════ dividers ═══════════════════════════ */

const MIN_FRAC = 0.03;

for (const id of ['divider0', 'divider1']) {
  const el = $(id);
  const index = Number(el.dataset.index);

  el.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    el.setPointerCapture(e.pointerId);
    el.classList.add('dragging');
    document.body.classList.add('dragging-divider');
  });

  el.addEventListener('pointermove', (e) => {
    if (!el.hasPointerCapture(e.pointerId)) return;
    const W = stage.clientWidth;
    const at = clamp(e.clientX / W, 0, 1);
    let [fl, fc, fr] = settings.fracs;
    if (index === 0) {
      const span = fl + fc;
      fl = clamp(at, MIN_FRAC, span - MIN_FRAC);
      fc = span - fl;
    } else {
      const sum = clamp(at, fl + MIN_FRAC, 1 - MIN_FRAC);
      fc = sum - fl;
      fr = 1 - sum;
    }
    settings.fracs = [fl, fc, fr];
    saveSettings();
    dirty = true;
  });

  const end = (e) => {
    if (el.hasPointerCapture(e.pointerId)) el.releasePointerCapture(e.pointerId);
    el.classList.remove('dragging');
    document.body.classList.remove('dragging-divider');
  };
  el.addEventListener('pointerup', end);
  el.addEventListener('pointercancel', end);
  el.addEventListener('dblclick', resetPanes);
}

function resetPanes() {
  settings.fracs = DEFAULTS.fracs.slice();
  saveSettings();
  dirty = true;
}

/* ═══════════════════════════ inspector ═══════════════════════════ */

function buildInspector() {
  const strip = $('filmstrip');
  strip.innerHTML = '';
  if (!pages.length) return;

  slides.forEach((group, i) => {
    const box = document.createElement('div');
    box.className = 'fs-slide';
    box.dataset.slide = String(i);
    box.innerHTML =
      `<div class="fs-slide-head"><b>Slide ${i + 1}</b>` +
      `<span>${group.length} step${group.length > 1 ? 's' : ''} · ` +
      `page${group.length > 1 ? 's' : ''} ${group[0] + 1}${group.length > 1 ? '–' + (group[group.length - 1] + 1) : ''}</span></div>` +
      `<div class="fs-steps"></div>`;
    const steps = box.querySelector('.fs-steps');
    group.forEach(p => {
      const t = document.createElement('div');
      t.className = 'fs-thumb';
      t.dataset.page = String(p);
      t.innerHTML = `<img src="${pages[p].thumb}" alt=""><span>${p + 1}</span>`;
      t.addEventListener('click', () => goToPage(p));
      steps.appendChild(t);
    });
    strip.appendChild(box);

    const last = group[group.length - 1];
    if (last + 1 < pages.length) strip.appendChild(junctionRow(last));
  });

  const all = mediaByPage.flat();
  const clips = all.filter(m => m.kind !== 'site').length;
  const sites = all.length - clips;
  $('inspectorSummary').textContent =
    `${slides.length} slides · ${pages.length} pages · ` +
    `${junctions.filter((_, k) => isBuild(k)).length} builds` +
    (clips ? ` · ${clips} video` : '') +
    (sites ? ` · ${sites} embed` : '');
  $('inspectorNote').textContent = pptxNote;
  $('inspectorNote').classList.toggle('hidden', !pptxNote);
  markCurrentInInspector();
}

function junctionRow(k) {
  const j = junctions[k];
  const build = isBuild(k);
  const row = document.createElement('div');
  row.className = 'junction ' + (build ? 'build' : 'rotate');
  row.innerHTML =
    `<button>${build ? 'build' : 'rotate'}</button>` +
    `<span class="score">p${k + 1}→p${k + 2} · score ${j.score.toFixed(3)}` +
    ` · vis ${j.visual.toFixed(2)}${j.subset ? ' · text-subset' : ''}${j.titleSame ? ' · same-title' : ''}</span>` +
    (overrides[k] ? '<span class="manual">manual</span>' : '');
  row.querySelector('button').addEventListener('click', () => {
    overrides[k] = build ? 'rotate' : 'build';
    if (overrides[k] === (autoIsBuild(k) ? 'build' : 'rotate')) delete overrides[k];
    if (deckKey) writeJSON(deckKey, overrides);
    regroup();
    buildInspector();
  });
  return row;
}

function markCurrentInInspector() {
  const strip = $('filmstrip');
  if (!strip.children.length) return;
  strip.querySelectorAll('.fs-slide').forEach(el =>
    el.classList.toggle('current', Number(el.dataset.slide) === slideIdx));
  const p = currentPage();
  strip.querySelectorAll('.fs-thumb').forEach(el =>
    el.classList.toggle('current', Number(el.dataset.page) === p));
  const cur = strip.querySelector('.fs-slide.current');
  if (cur && !$('inspector').classList.contains('hidden')) {
    cur.scrollIntoView({ block: 'nearest' });
  }
}

/* ═══════════════════════════ settings ui ═══════════════════════════ */

const bind = (id, key, out, fmt = (v) => v) => {
  const el = $(id);
  const isCheck = el.type === 'checkbox';
  const apply = () => {
    const v = isCheck ? el.checked : (el.type === 'range' ? Number(el.value) : el.value);
    settings[key] = v;
    if (out) $(out).textContent = fmt(v);
    saveSettings();
    afterSettingChange(key);
    dirty = true;
  };
  const sync = () => {
    if (isCheck) el.checked = settings[key];
    else el.value = settings[key];
    if (out) $(out).textContent = fmt(settings[key]);
  };
  el.addEventListener('input', apply);
  sync();
  return sync;
};

function afterSettingChange(key) {
  if (key === 'bg') document.documentElement.style.setProperty('--bg', settings.bg);
  if (key === 'threshold') { regroup(); buildInspector(); }
  if (key === 'videoAuto' || key === 'videoLoop') playbackSig = '';
  if (key === 'webEmbed' || key === 'webEmbedBlank') refreshEmbedsSoon();
  if (key === 'embedZoom' || key === 'embedPreload') dirty = true;
}

const syncers = [
  bind('setSideScale', 'sideScale', 'outSideScale', v => Math.round(v * 100) + '%'),
  bind('setDim', 'dim', 'outDim', v => Math.round(v * 100) + '%'),
  bind('setFit', 'fit'),
  bind('setFold', 'foldDeg', 'outFold', v => v + '°'),
  bind('setDepth', 'depth', 'outDepth', v => Number(v).toFixed(1) + '×'),
  bind('setRotateMs', 'rotateMs', 'outRotateMs', v => v + 'ms'),
  bind('setStepMs', 'stepMs', 'outStepMs', v => v + 'ms'),
  bind('setStepStyle', 'stepStyle'),
  bind('setGutter', 'gutter', 'outGutter', v => v + 'px'),
  bind('setBg', 'bg'),
  bind('setThreshold', 'threshold', 'outThreshold', v => Number(v).toFixed(3)),
  bind('setVideoAuto', 'videoAuto'),
  bind('setVideoLoop', 'videoLoop'),
  bind('setWebEmbed', 'webEmbed'),
  bind('setWebEmbedBlank', 'webEmbedBlank'),
  bind('setEmbedZoom', 'embedZoom', 'outEmbedZoom', v => Math.round(v * 100) + '%'),
  bind('setEmbedPreload', 'embedPreload', 'outEmbedPreload',
    v => Number(v) >= 99 ? 'all' : `${v} ahead`),
];
const syncSettingsUI = () => syncers.forEach(fn => fn());
document.documentElement.style.setProperty('--bg', settings.bg);

$('btnResetPanes').addEventListener('click', resetPanes);
if (/^(localhost|127\.0\.0\.1|\[::1\])$/.test(location.hostname)) {
  $('btnCacheVideos').addEventListener('click', (e) => cacheVideos(e.currentTarget));
} else {
  // a static host has no yt-dlp to call; publish --cache-clips does it instead
  $('btnCacheVideos').remove();
}
$('btnClearOverrides').addEventListener('click', () => {
  overrides = {};
  if (deckKey) writeJSON(deckKey, overrides);
  regroup();
  buildInspector();
});

/* panels */
function togglePanel(id) {
  const el = $(id);
  const other = id === 'settings' ? 'inspector' : 'settings';
  $(other).classList.add('hidden');
  el.classList.toggle('hidden');
  if (id === 'inspector' && !el.classList.contains('hidden')) markCurrentInInspector();
  if (id === 'settings' && !el.classList.contains('hidden')) syncSettingsUI();
}
document.querySelectorAll('[data-close]').forEach(b =>
  b.addEventListener('click', () => $(b.dataset.close).classList.add('hidden')));

/* ═══════════════════════════ file input ═══════════════════════════ */

const readFile = (file) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(r.result);
  r.onerror = () => reject(r.error || new Error('could not read ' + file.name));
  r.readAsArrayBuffer(file);
});

/** Takes a PDF, a PPTX, or both — the PDF draws, the PPTX supplies the video. */
async function openFiles(list) {
  const files = [...(list || [])];
  const pdf = files.find(f => /\.pdf$/i.test(f.name));
  const pptx = files.find(f => /\.pptx$/i.test(f.name));
  if (!pdf && !pptx) return;

  try {
    if (pdf) {
      mediaByPage = [];
      pptxNote = '';
      await loadPdf({ data: new Uint8Array(await readFile(pdf)) }, pdf.name);
    }
    if (pptx) {
      if (!pages.length) {
        return reportLoadError(new Error(
          'a .pptx on its own has no slides to draw — export the deck to PDF ' +
          'and drop both files together'));
      }
      await attachPptx(await readFile(pptx), pptx.name);
    }
  } catch (err) {
    reportLoadError(err);
  }
}

$('fileInput').addEventListener('change', (e) => openFiles(e.target.files));
dropzone.addEventListener('click', (e) => { if (e.target === dropzone) $('fileInput').click(); });

['dragenter', 'dragover'].forEach(t => window.addEventListener(t, (e) => {
  e.preventDefault();
  dropzone.classList.remove('hidden');
  dropzone.classList.add('dragover');
}));
window.addEventListener('dragleave', (e) => {
  if (e.relatedTarget) return;
  dropzone.classList.remove('dragover');
  if (pages.length) dropzone.classList.add('hidden');
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const files = [...(e.dataTransfer.files || [])]
    .filter(f => /\.(pdf|pptx)$/i.test(f.name));
  if (files.length) openFiles(files);
  else if (pages.length) dropzone.classList.add('hidden');
});

function reportLoadError(err) {
  console.error(err);
  $('loading').classList.add('hidden');
  dropzone.classList.remove('hidden');
  dropzone.querySelector('.dz-sub').innerHTML =
    `<span style="color:#e06c6c">${escapeHtml(String(err.message || err))}</span> ` +
    `<label class="link" for="fileInput">Try again</label>.`;
}

const escapeHtml = (s) => s.replace(/[&<>"]/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

const refreshEmbedsSoon = debounce(() => refreshEmbeds(), 700);

/** Re-apply the blank-slide embed after the setting changes. */
function refreshEmbeds() {
  for (const list of mediaByPage) {
    const keep = list.filter(m => !(m.kind === 'site' && m.fromBlank));
    list.length = 0;
    list.push(...keep);
  }
  fillBlankSlides(lastBlanks);
  for (const frame of embedFrames.values()) frame.remove();
  embedFrames.clear();
  playbackSig = '';
  for (const h of panesEl.querySelectorAll('.vids')) h.dataset.stamp = '';
  dirty = true;
}

/* ───────── a published talk ───────── */

/* On a static host there is no server to ask, so publish.py has already done
 * the work: the deck is a PDF beside a talk.json holding everything the .pptx
 * was needed for — which slides carry video, which were left empty, and which
 * live pages go where. */
async function openTalk(base) {
  const at = (file) => base.replace(/\/?$/, '/') + file;

  dropzone.classList.add('hidden');
  $('loading').classList.remove('hidden');
  $('loadTitle').textContent = 'Loading the talk…';

  try {
    const res = await withTimeout(fetch(at('talk.json'), { cache: 'no-cache' }), 30000,
      'the talk did not load in 30s');
    if (!res.ok) throw new Error(`no talk.json at ${at('talk.json')} (${res.status})`);
    const talk = await res.json();

    embedConfig = {
      blank: talk.embeds?.blank || '',
      slides: talk.embeds?.slides || {},
    };
    mediaCacheBase = base.replace(/\/?$/, '/');
    mediaByPage = [];
    pptxNote = '';

    // the version is the deck's content hash, so a re-published deck can never
    // come out of a browser cache still holding the old slides
    const deckUrl = at(talk.pdf || 'deck.pdf') + (talk.version ? `?v=${talk.version}` : '');
    const fresh = params.get('fresh') === '0' ? null : await freshFromGoogle(talk);

    if (fresh) {
      await loadPdf({ data: new Uint8Array(fresh.pdf) }, talk.title || 'talk');
      await attachPptx(fresh.pptx, talk.title || 'talk');
    } else {
      const pdf = await fetch(deckUrl);
      if (!pdf.ok) throw new Error(`the deck is missing (${pdf.status})`);
      await loadPdf({ data: new Uint8Array(await pdf.arrayBuffer()) }, talk.title || 'talk');
      await attachDeck({
        slideCount: talk.slideCount ?? (talk.media || []).length,
        media: (talk.media || []).map(list => list.map(m => ({ ...m }))),
        blanks: talk.blanks || [],
      }, talk.title || 'talk');
    }

    if (talk.title) document.title = talk.title + ' — Triptych';
  } catch (err) {
    reportLoadError(err);
  }
}

/* Google's export endpoints send CORS headers, so a talk published from Google
 * Slides can check for a newer deck as it opens — no server, no re-publish. It
 * is strictly best-effort: on a slow or hostile network the published snapshot
 * is used instead, which is the copy that has to work when the wifi doesn't.
 * `?fresh=0` skips the check outright. */
const FRESH_TIMEOUT = 9000;

async function freshFromGoogle(talk) {
  if (talk.source?.kind !== 'gslides' || !talk.source.id) return null;
  const base = `https://docs.google.com/presentation/d/${talk.source.id}/export/`;
  $('loadTitle').textContent = 'Checking Google Slides…';

  const grab = (fmt) => fetch(base + fmt, { cache: 'no-store', mode: 'cors' })
    .then(r => (r.ok ? r.arrayBuffer() : Promise.reject(new Error(String(r.status)))));

  try {
    const [pdf, pptx] = await withTimeout(
      Promise.all([grab('pdf'), grab('pptx')]), FRESH_TIMEOUT, 'slow');
    if (!pdf.byteLength) return null;
    console.log('[triptych] using the current deck from Google Slides');
    return { pdf, pptx };
  } catch (err) {
    console.log(`[triptych] keeping the published deck (${err.message})`);
    return null;
  }
}

/* ───────── google slides ───────── */

/* The browser can't fetch docs.google.com itself, so serve.py hands both
 * exports back from this origin: the PDF draws the slides, the PPTX carries
 * the video links. One live deck, so the two always agree on slide count. */
async function openGoogleSlides(link) {
  const id = encodeURIComponent(String(link || '').trim());
  if (!id) return;
  const api = (fmt) => `/gslides?fmt=${fmt}&id=${id}`;

  dropzone.classList.add('hidden');
  $('loading').classList.remove('hidden');
  $('loadTitle').textContent = 'Fetching from Google Slides…';

  try {
    const res = await withTimeout(fetch(api('pdf')), 90000,
      'Google Slides did not answer in 90s — check the share setting and your connection');
    if (!res.ok) throw new Error(await res.text() || `Google Slides fetch failed (${res.status})`);
    $('loadTitle').textContent = 'Reading the deck…';

    mediaByPage = [];
    pptxNote = '';
    await loadPdf({ data: new Uint8Array(await res.arrayBuffer()) }, 'Google Slides deck');

    const deck = await fetch(api('pptx'));
    if (deck.ok) await attachPptx(await deck.arrayBuffer(), 'Google Slides deck');
  } catch (err) {
    reportLoadError(err);
  }
}

/** A fetch that never settles reads as a frozen app; make it say so instead. */
function withTimeout(promise, ms, message) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(message)), ms)),
  ]);
}

$('gsForm').addEventListener('submit', (e) => {
  e.preventDefault();
  openGoogleSlides($('gsInput').value);
});

/* ?src=deck.pdf — skip the drop step when the deck sits next to index.html.
 * A .pptx of the same name beside it is picked up automatically.
 * ?gslides=<link or id> does the same for a Google Slides deck. */
const params = new URLSearchParams(location.search);

const configReady = loadEmbedConfig();

const webParam = params.get('web');
if (webParam) { settings.webEmbed = webParam; saveSettings(); }

/* A published talk: ?talk=<slug>, or a page that names one directly. */
const talkParam = params.get('talk') || window.TRIPTYCH_TALK;
if (talkParam) {
  openTalk(/^https?:|^\.|^\//.test(talkParam) ? talkParam : `talks/${talkParam}`);
}

const gslides = params.get('gslides');
if (gslides) openGoogleSlides(gslides);
const src = params.get('src');
if (src) {
  loadPdf({ url: src }, src.split('/').pop())
    .then(() => {
      const sibling = params.get('pptx') || src.replace(/\.pdf$/i, '.pptx');
      if (sibling === src) return;
      return fetch(sibling)
        .then(r => r.ok ? r.arrayBuffer() : null)
        .then(buf => buf && attachPptx(buf, sibling.split('/').pop()))
        .catch(() => { /* no sibling deck; images only */ });
    })
    .catch(reportLoadError);
}

/* ═══════════════════════════ input ═══════════════════════════ */

window.addEventListener('keydown', (e) => {
  if (e.target instanceof Element && e.target.matches('input, select, textarea')) return;
  const k = e.key;

  if (k === 'Escape') {
    $('settings').classList.add('hidden');
    $('inspector').classList.add('hidden');
    return;
  }
  if (!pages.length) return;

  switch (k) {
    case 'ArrowRight': case ' ': case 'PageDown': case 'n':
      e.preventDefault(); e.shiftKey ? nextSlide() : next(); break;
    case 'ArrowLeft': case 'PageUp': case 'p': case 'Backspace':
      e.preventDefault(); e.shiftKey ? prevSlide() : prev(); break;
    case 'ArrowDown': e.preventDefault(); nextSlide(); break;
    case 'ArrowUp': e.preventDefault(); prevSlide(); break;
    case 'Home': e.preventDefault(); goToPage(0); break;
    case 'End': e.preventDefault(); goToPage(pages.length - 1); break;
    case 'g': togglePanel('inspector'); break;
    case 's': togglePanel('settings'); break;
    case 'o': $('fileInput').click(); break;
    case 'm': toggleSound(); break;
    // flip who gets the mouse: a live page takes it by default so you can
    // drive the site, a video player doesn't so its chrome stays hidden
    case 'v': document.body.classList.toggle('embeds-flipped'); break;
    case 'r': resetPanes(); break;
    case 'f':
      if (document.fullscreenElement) document.exitFullscreen();
      else document.documentElement.requestFullscreen?.();
      break;
  }
});

let wheelLock = 0;
stage.addEventListener('wheel', (e) => {
  if (Math.abs(e.deltaX) < Math.abs(e.deltaY) || Math.abs(e.deltaX) < 24) return;
  const now = performance.now();
  if (now - wheelLock < 420) return;
  wheelLock = now;
  e.deltaX > 0 ? nextSlide() : prevSlide();
}, { passive: true });

window.addEventListener('resize', () => { dirty = true; });
document.addEventListener('fullscreenchange', () => { dirty = true; });

/* ═══════════════════════════ helpers ═══════════════════════════ */

function clamp(v, lo, hi) { return v < lo ? lo : v > hi ? hi : v; }
function lerp(a, b, t) { return a + (b - a) * t; }
function easeInOutCubic(t) { return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2; }

function normalizeFracs(f) {
  if (!Array.isArray(f) || f.length !== 3 || f.some(v => !(v > 0))) return DEFAULTS.fracs.slice();
  const sum = f[0] + f[1] + f[2];
  return f.map(v => v / sum);
}

function hashString(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return (h >>> 0).toString(36);
}

/** Give the browser a turn between pages. A MessageChannel task isn't clamped
 *  to 1 Hz the way setTimeout is in a hidden tab, so ingest keeps its pace. */
function breathe() {
  return new Promise((resolve) => {
    const ch = new MessageChannel();
    ch.port1.onmessage = () => { ch.port1.close(); resolve(); };
    ch.port2.postMessage(0);
  });
}

function debounce(fn, ms) {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}
