/* Triptych — a PDF presenter with two kinds of transition.
 *
 *   rotation    the panes move across the triptych: next slide swings into the
 *               centre, the centre shrinks into the left pane.
 *   within-slide the centre pane advances one build step; nothing moves.
 *
 * Which junction is which is inferred from the deck itself (see the grouping
 * section below); drag a card onto its neighbour in the timeline to say
 * otherwise.
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
/** @type {Record<number,'build'|'rotate'>} stacking set by hand, per junction */
let overrides = {};
/** @type {Record<number,string>} a web page laid over a page, keyed by page index */
let deckEmbeds = {};
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

  deckKey = 'triptych:deck:' + hashString(n + ':' + pages.map(p => p.tokens.size).join(','));
  const saved = readJSON(deckKey) || {};
  overrides = saved.overrides || (saved.embeds ? {} : saved);   // older decks stored bare
  deckEmbeds = saved.embeds || {};

  for (const frame of embedFrames.values()) frame.remove();
  embedFrames.clear();

  junctions = [];
  for (let k = 0; k + 1 < pages.length; k++) junctions.push(similarity(pages[k], pages[k + 1]));

  regroup(0);
  $('loading').classList.add('hidden');
  document.title = deckName + ' — Triptych';
  dirty = true;
  buildTimeline();
  showSetup();                                 // arrange first, then present
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

const saveDeck = () => { if (deckKey) writeJSON(deckKey, { overrides, embeds: deckEmbeds }); };

/* ═══════════════════════════ geometry ═══════════════════════════ */

const SHUT_TURN = 22;     // extra degrees a closing panel turns out toward the viewer

/* Anchor rects for the slots a pane can occupy. Slot 0 is the centre;
 * ±1 are the wings; ±2 sit just off-stage, which is what makes a rotation
 * read as panes sliding across the triptych rather than a crossfade. */
function anchor(slot, W) {
  const [fl, fc, fr] = settings.fracs;
  if (slot <= -1) return { x: 0, w: fl * W };
  if (slot >= 1) return { x: (fl + fc) * W, w: fr * W };
  return { x: fl * W, w: fc * W };
}

/* The panels are a chain, not three loose flaps. Past a wing, a panel keeps its
 * inner edge stitched to the outer edge of the panel ahead of it — the one
 * taking its place — while its own outer edge stays pinned to the edge of the
 * stage. So it closes as its neighbour slides across, and reaches nothing
 * exactly as that neighbour arrives. */
function geomAt(s, W) {
  const c = clamp(s, -2, 2);
  if (c <= -1) {
    return { x: 0, w: Math.max(0, geomAt(c + 1, W).x) };
  }
  if (c >= 1) {
    const ahead = geomAt(c - 1, W);
    const left = ahead.x + ahead.w;
    return { x: left, w: Math.max(0, W - left) };
  }
  const lo = Math.floor(c), hi = Math.ceil(c), t = c - lo;
  const a = anchor(lo, W), b = anchor(hi, W);
  return { x: lerp(a.x, b.x, t), w: lerp(a.w, b.w, t) };
}

/* The panels are one folding screen, not three loose flaps.
 *
 * Each is hinged to its neighbours, and a joint is a single point in space —
 * so the two panels meeting there share an edge exactly: same place, and the
 * same height once perspective has had its way with it. That only works if the
 * chain is built in 3D, carrying depth from joint to joint, which is what this
 * does. The dividers still decide the apportionment: each panel's length is
 * solved so that its far edge lands on its section boundary once projected.
 *
 * The fold opens toward the viewer, so the outer edges are the near, tall ones,
 * and a panel on its way out closes as its neighbour arrives in its place.
 */

const tiltAt = (s) =>
  settings.foldDeg * clamp(Math.abs(s), 0, 1) + SHUT_TURN * clamp(Math.abs(s) - 1, 0, 1);

/** How long must a panel be for its far edge to project onto `target`? */
function solveLength(hingeX, hingeZ, dir, rad, target, W, P) {
  const cx = W / 2, cos = Math.cos(rad), sin = Math.sin(rad);
  const at = (d) => {
    const z = hingeZ + d * sin;
    return cx + (hingeX + dir * d * cos - cx) * P / (P - z);
  };
  let lo = 0, hi = sin > 1e-6 ? 0.82 * (P - hingeZ) / sin : W * 6;
  for (let i = 0; i < 32; i++) {
    const mid = (lo + hi) / 2;
    const short = dir < 0 ? at(mid) > target : at(mid) < target;
    if (short) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

/* Walk the chain out from one joint, which stays at the back of the fold.
 *
 * It has to be a joint that sits at the back in the resting pose — one of the
 * two dividers — and that moves one seat over the course of a rotation, in
 * whichever direction you are going. Going forward it is the trailing edge of
 * the panel becoming the left wing, travelling from the right divider to the
 * left. Going back it is that panel's leading edge, travelling the other way.
 * The two rules meet at rest, where the centre panel is flat and both dividers
 * are at the back, which is what makes a rotation and its reverse the same
 * animation played each way. */
function buildChain(W, offset) {
  const P = settings.depth * W;
  const slotOf = (n) => n - offset;
  const base = geomAt(slotOf(0), W);
  const back = offset < 0;
  const joint = back ? base.x : base.x + base.w;
  const firstRight = back ? 0 : 1;      // the lowest panel that extends rightward
  const panels = new Map();

  let x = joint, z = 0;
  for (let n = firstRight - 1; n >= -2; n--) {
    const s = slotOf(n), deg = tiltAt(s), rad = deg * Math.PI / 180;
    const L = solveLength(x, z, -1, rad, geomAt(s, W).x, W, P);
    panels.set(n, { x, z, L, deg, dir: -1, s });
    x -= L * Math.cos(rad);
    z += L * Math.sin(rad);
  }

  x = joint; z = 0;
  for (let n = firstRight; n <= 2; n++) {
    const s = slotOf(n), deg = tiltAt(s), rad = deg * Math.PI / 180;
    const g = geomAt(s, W);
    const L = solveLength(x, z, 1, rad, g.x + g.w, W, P);
    panels.set(n, { x, z, L, deg: -deg, dir: 1, s });
    x += L * Math.cos(rad);
    z += L * Math.sin(rad);
  }
  return panels;
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

  const chain = buildChain(W, offset);

  let centreBox = null;
  const live = new Set();
  for (let j = slideIdx - 2; j <= slideIdx + 2; j++) {
    if (j < 0 || j >= slides.length) continue;
    const s = (j - slideIdx) - offset;
    if (Math.abs(s) > 2) continue;
    live.add(j);

    const el = acquirePane(j);
    const panel = chain.get(j - slideIdx);
    const k = clamp(Math.abs(s), 0, 1);
    const paneW = panel.L;

    /* The panel hangs off its hinge, at that joint's depth. Its neighbour hangs
     * off the very same point, which is what makes the two read as attached. */
    const tx = panel.dir < 0 ? panel.x - paneW : panel.x;
    el.style.transformOrigin = panel.dir < 0 ? '100% 50%' : '0% 50%';
    el.style.transform =
      `translate3d(${tx.toFixed(2)}px,0,${panel.z.toFixed(2)}px) ` +
      `rotateY(${panel.deg.toFixed(3)}deg)`;
    el.style.width = paneW + 'px';
    el.style.height = H + 'px';
    el.classList.toggle('clickable', j !== slideIdx || hasNext());

    const inner = el.firstElementChild;
    const isCentre = j === slideIdx;
    inner.style.transform = `translateY(${isCentre ? lift.toFixed(2) : 0}px)`;

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
    const centreH = Math.min(H, Math.max(1, anchor(0, W).w) / aspect);

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
    if (isCentre) centreBox = { ...box, left: box.left + tx };
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

/* Where a live page comes from, most specific first:
 *   1. dropped on that slide in the timeline
 *   2. `embed: <url>` in the slide's speaker notes (handled in pptx.js)
 *   3. that slide's number in embeds.json, or a published talk's talk.json
 *   4. the blank-slide default  */
function applyEmbeds(blanks) {
  lastBlanks = blanks || lastBlanks;
  if (!mediaByPage.length) return;

  for (const list of mediaByPage) {
    const keep = list.filter(m => !(m.kind === 'site' && m.placed));
    list.length = 0;
    list.push(...keep);
  }

  const inset = { x: 0.03, y: 0.04, w: 0.94, h: 0.92 };
  const put = (page, url) => {
    if (page == null || page < 0 || page >= mediaByPage.length) return;
    if (mediaByPage[page].some(m => m.kind === 'site')) return;
    mediaByPage[page].push({ url, kind: 'site', name: url, placed: true, ...inset });
  };

  for (const [page, url] of Object.entries(deckEmbeds)) put(Number(page), String(url).trim());

  const fallback = settings.webEmbedBlank
    ? (embedConfig.blank || settings.webEmbed || '').trim()
    : '';
  (lastBlanks || []).forEach((isBlank, i) => {
    const url = (embedConfig.slides[String(i + 1)] || (isBlank ? fallback : '')).trim();
    if (url) put(pageForSlide(i), url);
  });

  for (const [page, frame] of embedFrames) { frame.remove(); embedFrames.delete(page); }
  playbackSig = '';
}

/** Which PDF page a PowerPoint slide index landed on. */
let slideToPage = null;
const pageForSlide = (i) => (slideToPage ? slideToPage[i] ?? null : null);

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
  applyEmbeds(deck.blanks);
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

  applyEmbeds(deck.blanks);
  await useCachedVideos();

  const total = mediaByPage.reduce((n, m) => n + m.length, 0);
  console.log(`[triptych] ${name}: ${total} media item${total === 1 ? '' : 's'} ` +
    `across ${deck.slideCount} slides`);
  playbackSig = '';
  dirty = true;
  buildTimeline();        // media can add an embed marker to a slide
  if (DEBUG) {
    window.__triptych = {
      chain: (offset = 0) => [...buildChain(stage.clientWidth, offset).entries()]
        .sort((a, b) => a[0] - b[0]),
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
  markCurrentInTimeline();
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
  if (settings.stepMs <= 0) { anim = null; markCurrentInTimeline(); dirty = true; return; }
  anim = { kind: 'step', from, dir, t0: performance.now(), dur: settings.stepMs };
  markCurrentInTimeline();
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
  markCurrentInTimeline();
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

/* ═══════════════════════════ timeline ═══════════════════════════ */

/* Every slide as a card, in order, left to right. Slides that share a stack are
 * build steps: they hold the triptych still and change in place. Drag a card
 * onto its neighbour to stack it; drag it into a gap to pull it back out. */

const stripEl = () => $('strip');
let dragging = null;

function buildTimeline() {
  const strip = stripEl();
  strip.textContent = '';
  if (!pages.length) return;

  slides.forEach((group, i) => {
    strip.appendChild(gapEl(group[0]));

    const slot = document.createElement('div');
    slot.className = 'slot';
    slot.dataset.slide = String(i);
    slot.dataset.first = String(group[0]);
    slot.dataset.last = String(group[group.length - 1]);

    const cards = document.createElement('div');
    cards.className = 'cards';
    for (const page of group) cards.appendChild(cardEl(page, group.length));
    slot.appendChild(cards);

    const embed = document.createElement('button');
    embed.className = 'embed';
    embed.textContent = '⌘';
    embed.title = deckEmbeds[group[0]] ? 'Change the web page on this slide'
                                       : 'Put a web page on this slide';
    embed.addEventListener('click', (e) => { e.stopPropagation(); askForEmbed(slot, group[0]); });
    slot.appendChild(embed);

    slot.classList.toggle('has-embed', !!deckEmbeds[group[0]]);
    strip.appendChild(slot);
  });

  strip.appendChild(gapEl(pages.length));
  markCurrentInTimeline();
}

function cardEl(page, stackSize) {
  const card = document.createElement('div');
  card.className = 'card';
  card.dataset.page = String(page);
  card.title = stackSize > 1
    ? 'Click to present from here, drag into a gap to unstack'
    : 'Click to present from here, drag onto the next slide to stack';
  card.innerHTML = `<img src="${pages[page].thumb}" alt=""><b>${page + 1}</b>`;
  card.addEventListener('click', () => { if (!dragging) present(page); });
  card.addEventListener('pointerdown', (e) => beginDrag(e, card, page));
  return card;
}

function gapEl(before) {
  const gap = document.createElement('div');
  gap.className = 'gap';
  gap.dataset.before = String(before);
  return gap;
}

/* ───────── stacking ───────── */

const setJunction = (k, how) => {
  if (k < 0 || k >= junctions.length) return;
  if ((autoIsBuild(k) ? 'build' : 'rotate') === how) delete overrides[k];
  else overrides[k] = how;
};

function stackPages(page, onto) {
  setJunction(Math.min(page, onto), 'build');
  commitStacking(page);
}

function unstackPage(page) {
  setJunction(page - 1, 'rotate');
  setJunction(page, 'rotate');
  commitStacking(page);
}

function commitStacking(keepPage) {
  saveDeck();
  regroup(keepPage);
  applyEmbeds();
  buildTimeline();
  dirty = true;
}

/* ───────── dragging ───────── */

function beginDrag(event, card, page) {
  if (event.button !== 0) return;
  event.preventDefault();
  const startX = event.clientX, startY = event.clientY;
  card.setPointerCapture(event.pointerId);

  const move = (e) => {
    if (!dragging) {
      if (Math.hypot(e.clientX - startX, e.clientY - startY) < 5) return;
      const ghost = card.cloneNode(true);
      ghost.className = 'ghost';
      document.body.appendChild(ghost);
      document.body.classList.add('dragging-card');
      card.classList.add('dragging');
      dragging = { page, card, ghost };
    }
    dragging.ghost.style.left = (e.clientX - 66) + 'px';
    dragging.ghost.style.top = (e.clientY - 40) + 'px';
    highlight(dropTarget(e.clientX, e.clientY, page));
  };

  const end = (e) => {
    card.removeEventListener('pointermove', move);
    card.removeEventListener('pointerup', end);
    card.removeEventListener('pointercancel', end);
    if (!dragging) return;

    const target = dropTarget(e.clientX, e.clientY, page);
    dragging.ghost.remove();
    card.classList.remove('dragging');
    document.body.classList.remove('dragging-card');
    highlight(null);
    dragging = null;

    if (target?.kind === 'slot') stackPages(page, target.page);
    else if (target?.kind === 'gap') unstackPage(page);
  };

  card.addEventListener('pointermove', move);
  card.addEventListener('pointerup', end);
  card.addEventListener('pointercancel', end);
}

/* A page can only join the slide before or after it, because the deck's order
 * is the deck's order. Anything else is not a target. */
function dropTarget(x, y, page) {
  const under = document.elementFromPoint(x, y);
  if (!under) return null;

  const slot = under.closest('.slot');
  if (slot) {
    const first = Number(slot.dataset.first), last = Number(slot.dataset.last);
    if (page >= first && page <= last) return null;              // already here
    if (page === last + 1 || page === first - 1) {
      return { kind: 'slot', page: page === last + 1 ? last : first, el: slot };
    }
    return null;
  }

  const gap = under.closest('.gap');
  if (gap && slides[slideOfPage(page)]?.length > 1) {
    const before = Number(gap.dataset.before);
    if (before === page || before === page + 1) return { kind: 'gap', el: gap };
  }
  return null;
}

function highlight(target) {
  for (const el of stripEl().querySelectorAll('.target')) el.classList.remove('target');
  target?.el?.classList.add('target');
}

function markCurrentInTimeline() {
  const strip = stripEl();
  if (!strip.children.length) return;
  const page = currentPage();
  let active = null;
  for (const card of strip.querySelectorAll('.card')) {
    const on = Number(card.dataset.page) === page;
    card.classList.toggle('current', on);
    if (on) active = card;
  }
  if (active && inSetup()) active.scrollIntoView({ block: 'nearest', inline: 'center' });
}

/* ───────── embeds ───────── */

function askForEmbed(slot, page) {
  slot.querySelector('.embed-field')?.remove();

  const field = document.createElement('div');
  field.className = 'embed-field';
  const input = document.createElement('input');
  input.type = 'text';
  input.spellcheck = false;
  input.placeholder = 'https://';
  input.value = deckEmbeds[page] || '';
  input.title = 'Enter to set, empty to remove';
  field.appendChild(input);
  slot.appendChild(field);
  input.focus();
  input.select();

  const close = () => field.remove();
  input.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') close();
    if (e.key !== 'Enter') return;
    const url = input.value.trim();
    if (url) deckEmbeds[page] = /^https?:\/\//.test(url) ? url : 'https://' + url;
    else delete deckEmbeds[page];
    saveDeck();
    applyEmbeds();
    buildTimeline();
    dirty = true;
  });
  input.addEventListener('blur', close);
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
  if (key === 'threshold') { regroup(); applyEmbeds(); buildTimeline(); }
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
  commitStacking(currentPage());
});

/* panels */
function togglePanel(id) {
  const el = $(id);
  el.classList.toggle('hidden');
  if (id === 'settings' && !el.classList.contains('hidden')) syncSettingsUI();
}

/* Two states, and only two: setting the deck up, or presenting it. The timeline
 * belongs to the first — it is where the deck is arranged, not something to
 * consult mid-talk. */
function showSetup() {
  $('timeline').classList.toggle('hidden', !pages.length);
  dropzone.classList.remove('hidden');
  $('settings').classList.add('hidden');
  if (pages.length) markCurrentInTimeline();
  dirty = true;
}

function present(page) {
  if (!pages.length) return;
  if (page !== undefined) goToPage(page);
  dropzone.classList.add('hidden');
  dirty = true;
}

const inSetup = () => !dropzone.classList.contains('hidden');
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
$('gsInput').addEventListener('input', clearLoadError);

['dragenter', 'dragover'].forEach(t => window.addEventListener(t, (e) => {
  e.preventDefault();
  document.body.classList.add('dropping');
}));
window.addEventListener('dragleave', (e) => {
  if (!e.relatedTarget) document.body.classList.remove('dropping');
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  document.body.classList.remove('dropping');
  const files = [...(e.dataTransfer.files || [])].filter(f => /\.(pdf|pptx)$/i.test(f.name));
  if (files.length) openFiles(files);
});

function reportLoadError(err) {
  console.error(err);
  $('loading').classList.add('hidden');
  showSetup();
  const note = $('askError');
  note.textContent = String(err.message || err);
  note.classList.remove('hidden');
  $('gsInput').focus();
  $('gsInput').select();
}

function clearLoadError() { $('askError').classList.add('hidden'); }

const refreshEmbedsSoon = debounce(() => refreshEmbeds(), 700);

const refreshEmbeds = () => { applyEmbeds(); buildTimeline(); dirty = true; };

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
/** A deck id out of anything Google might hand you. */
function deckIdFrom(value) {
  const text = String(value || '').trim();
  const inUrl = /\/presentation\/d\/(?:e\/)?([A-Za-z0-9_-]{10,120})/.exec(text);
  if (inUrl) return inUrl[1];
  return /^[A-Za-z0-9_-]{10,120}$/.test(text) ? text : null;
}

/* Google's export endpoints send CORS headers, so the browser can fetch them
 * itself — which is what makes this work on a static host, where serve.py and
 * its /gslides is not there to ask. The proxy stays as the fallback for a
 * network that blocks the direct call. */
async function openGoogleSlides(link) {
  const id = deckIdFrom(link);
  if (!id) return reportLoadError(new Error('That is not a Google Slides link.'));

  dropzone.classList.add('hidden');
  $('loading').classList.remove('hidden');
  $('loadTitle').textContent = 'Fetching from Google Slides';
  setLoadProgress(0, 0);

  const routes = [
    (fmt) => `https://docs.google.com/presentation/d/${id}/export/${fmt}`,
    (fmt) => `/gslides?fmt=${fmt}&id=${encodeURIComponent(id)}`,
  ];

  const grab = async (fmt, optional) => {
    let last = null;
    for (const route of routes) {
      try {
        const res = await fetch(route(fmt), { cache: 'no-store' });
        if (res.ok) return await res.arrayBuffer();
        last = new Error(res.status === 404 || res.status === 403
          ? 'Google would not hand over that deck. Share it with anyone who has the link.'
          : `Google Slides returned ${res.status}.`);
      } catch (err) {
        last = err;
      }
    }
    if (optional) return null;
    throw last || new Error('Could not reach Google Slides.');
  };

  try {
    const pdf = await grab('pdf');
    mediaByPage = [];
    pptxNote = '';
    mediaCacheBase = '/';
    await loadPdf({ data: new Uint8Array(pdf) }, 'Google Slides deck');

    const pptx = await grab('pptx', true);
    if (pptx) await attachPptx(pptx, 'Google Slides deck');
    buildTimeline();
  } catch (err) {
    reportLoadError(err);
  }
}

$('gsForm').addEventListener('submit', (e) => {
  e.preventDefault();
  clearLoadError();
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
    if (!$('settings').classList.contains('hidden')) $('settings').classList.add('hidden');
    else if (inSetup()) present();
    else showSetup();
    dirty = true;
    return;
  }
  if (!pages.length) return;

  // in setup, the deck is arranged rather than driven; one key starts the talk
  if (inSetup()) {
    if (k === 'ArrowRight' || k === ' ' || k === 'Enter') { e.preventDefault(); present(); }
    return;
  }
  switch (k) {
    case 'ArrowRight': case ' ': case 'PageDown': case 'n':
      e.preventDefault(); e.shiftKey ? nextSlide() : next(); break;
    case 'ArrowLeft': case 'PageUp': case 'p': case 'Backspace':
      e.preventDefault(); e.shiftKey ? prevSlide() : prev(); break;
    case 'ArrowDown': e.preventDefault(); nextSlide(); break;
    case 'ArrowUp': e.preventDefault(); prevSlide(); break;
    case 'Home': e.preventDefault(); goToPage(0); break;
    case 'End': e.preventDefault(); goToPage(pages.length - 1); break;

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
