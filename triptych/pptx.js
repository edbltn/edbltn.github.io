/* Reading a .pptx for the things a PDF export throws away — chiefly video.
 *
 * A PowerPoint file is a zip of XML. Chrome can inflate deflate streams
 * natively, so there is no library here: about a hundred lines of central
 * directory walking gets us the parts, and DOMParser does the rest.
 *
 * We deliberately do *not* try to render slides from the XML — the PDF export
 * already does that perfectly. All we want is: which slides carry a video, what
 * file it is, and where on the slide it sits.
 */

/* ───────────────────────── zip ───────────────────────── */

const SIG_EOCD = 0x06054b50;
const SIG_CDIR = 0x02014b50;

export async function openZip(buffer) {
  const dv = new DataView(buffer);
  const size = buffer.byteLength;

  let eocd = -1;
  for (let i = size - 22; i >= Math.max(0, size - 66000); i--) {
    if (dv.getUint32(i, true) === SIG_EOCD) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('not a zip file');

  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = new Map();
  const dec = new TextDecoder();

  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== SIG_CDIR) break;
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = dec.decode(new Uint8Array(buffer, p + 46, nameLen));
    entries.set(name, { method, csize, local });
    p += 46 + nameLen + extraLen + commentLen;
  }

  const bytesOf = async (name) => {
    const e = entries.get(name);
    if (!e) return null;
    // the local header repeats the name/extra lengths, and they can differ
    const nameLen = dv.getUint16(e.local + 26, true);
    const extraLen = dv.getUint16(e.local + 28, true);
    const start = e.local + 30 + nameLen + extraLen;
    const raw = buffer.slice(start, start + e.csize);
    if (e.method === 0) return raw;
    if (e.method !== 8) throw new Error(`unsupported zip method ${e.method} for ${name}`);
    const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return await new Response(stream).arrayBuffer();
  };

  return {
    names: () => [...entries.keys()],
    has: (n) => entries.has(n),
    bytes: bytesOf,
    text: async (n) => {
      const b = await bytesOf(n);
      return b && new TextDecoder().decode(b);
    },
    blob: async (n, type) => {
      const b = await bytesOf(n);
      return b && new Blob([b], type ? { type } : undefined);
    },
  };
}

/* ───────────────────────── xml helpers ───────────────────────── */

const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml');

/** Search by local name so we never depend on a namespace prefix. */
const descendants = (root, localName) =>
  [...root.getElementsByTagName('*')].filter(n => n.localName === localName);
const firstDescendant = (root, localName) =>
  [...root.getElementsByTagName('*')].find(n => n.localName === localName) || null;

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Relationship attributes (r:id, r:embed, r:link). Must go through the
 *  namespace: <p:sldId id="256" r:id="rId2"/> carries two `id` attributes and
 *  only one of them names a relationship. */
function relAttr(el, ...localNames) {
  for (const n of localNames) {
    const v = el.getAttributeNS(R_NS, n);
    if (v) return v;
  }
  // last resort for documents that declare the prefix unconventionally
  for (const a of el.attributes) {
    if (a.name.startsWith('r:') && localNames.includes(a.localName)) return a.value;
  }
  return null;
}

/** Resolve "../media/media1.mp4" against "ppt/slides/slide3.xml".
 *  A leading slash means the target is already package-absolute — some
 *  exporters write "/ppt/media/image.png" instead of a relative path. */
function resolvePart(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '.' || seg === '') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

async function relsFor(zip, partName) {
  const dir = partName.split('/').slice(0, -1).join('/');
  const file = partName.split('/').pop();
  const xml = await zip.text(`${dir}/_rels/${file}.rels`);
  const map = new Map();
  if (!xml) return map;
  for (const r of descendants(parseXml(xml), 'Relationship')) {
    map.set(r.getAttribute('Id'), {
      target: r.getAttribute('Target'),
      external: r.getAttribute('TargetMode') === 'External',
    });
  }
  return map;
}

/** Speaker-note text for a slide, or '' if it has none. */
async function notesFor(zip, slidePart, rels) {
  for (const rel of rels.values()) {
    if (!rel.target || !/notesSlide\d*\.xml$/.test(rel.target)) continue;
    const xml = await zip.text(resolvePart(slidePart, rel.target));
    if (!xml) continue;
    return descendants(parseXml(xml), 't').map(n => n.textContent).join(' ');
  }
  return '';
}

/* ───────────────────────── deck ───────────────────────── */

const MIME = {
  mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime',
  webm: 'video/webm', ogv: 'video/ogg', avi: 'video/x-msvideo',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', aac: 'audio/aac',
};
const mimeFor = (name) => MIME[name.split('.').pop().toLowerCase()] || '';

/** A hyperlink to a video host, turned into something embeddable. */
export function hostedVideo(href) {
  if (!href) return null;
  const yt = /(?:youtube\.com\/(?:watch\?(?:.*&)?v=|embed\/|shorts\/|live\/)|youtu\.be\/)([A-Za-z0-9_-]{6,})/.exec(href);
  if (yt) return { kind: 'youtube', id: yt[1], embed: yt[1] };
  const vim = /vimeo\.com\/(?:video\/)?(\d{6,})/.exec(href);
  if (vim) return { kind: 'vimeo', id: vim[1], embed: vim[1] };
  return null;
}

/**
 * @returns {Promise<{slideCount:number, media:Array<Array<{url:string,kind:'video'|'audio',
 *          x:number,y:number,w:number,h:number,name:string,external:boolean}>>}>}
 *          `media[i]` is the media on slide i, in slide-relative fractions.
 */
export async function readPptx(buffer) {
  const zip = await openZip(buffer);

  const presXml = await zip.text('ppt/presentation.xml');
  if (!presXml) throw new Error('no ppt/presentation.xml — is this a .pptx?');
  const pres = parseXml(presXml);

  const sz = firstDescendant(pres, 'sldSz');
  const slideW = sz ? Number(sz.getAttribute('cx')) : 12192000;
  const slideH = sz ? Number(sz.getAttribute('cy')) : 6858000;

  const presRels = await relsFor(zip, 'ppt/presentation.xml');
  const slideParts = descendants(pres, 'sldId')
    .map(el => presRels.get(relAttr(el, 'id')))
    .filter(Boolean)
    .map(rel => resolvePart('ppt/presentation.xml', rel.target));

  const media = [];
  const blanks = [];
  const urlCache = new Map();

  for (const part of slideParts) {
    const xml = await zip.text(part);
    const found = [];
    if (xml) {
      const rels = await relsFor(zip, part);
      const doc = parseXml(xml);

      /* A slide holding nothing at all is a placeholder — somewhere to put a
       * live page. Ask the deck rather than guessing from the raster: a themed
       * background would defeat any "is it uniform?" test. */
      const tree = firstDescendant(doc, 'spTree');
      blanks.push(!tree || !['sp', 'pic', 'graphicFrame', 'grpSp'].some(
        name => descendants(tree, name).length));

      // an explicit `embed: <url>` in the speaker notes wins over everything
      const note = await notesFor(zip, part, rels);
      const marked = /^[ \t]*(?:embed|iframe|site|web)[ \t]*[:=][ \t]*(https?:\/\/\S+)/im.exec(note);
      if (marked) found.push({ url: marked[1], kind: 'site', name: marked[1], x: 0, y: 0, w: 1, h: 1 });

      const shapes = [...descendants(doc, 'pic'), ...descendants(doc, 'graphicFrame'),
                      ...descendants(doc, 'sp')];

      for (const pic of shapes) {
        // Two ways a slide carries video:
        //   1. an embedded (or linked) media part named by nvPr
        //   2. a poster picture hyperlinked to YouTube — which is what both
        //      PowerPoint's and Google Slides' "online video" become on export
        const ref = descendants(pic, 'videoFile')[0]
          || descendants(pic, 'media')[0]
          || descendants(pic, 'audioFile')[0];

        let kind, url, name;

        if (ref) {
          const relId = relAttr(ref, 'embed', 'link');
          const rel = relId && rels.get(relId);
          if (!rel) continue;

          kind = ref.localName === 'audioFile' ? 'audio' : 'video';
          if (rel.external) {
            url = rel.target;                    // a link out to a file or a URL
            name = rel.target.split('/').pop();
          } else {
            const target = resolvePart(part, rel.target);
            name = target.split('/').pop();
            if (!urlCache.has(target)) {
              const blob = await zip.blob(target, mimeFor(name));
              urlCache.set(target, blob ? URL.createObjectURL(blob) : null);
            }
            url = urlCache.get(target);
          }
        } else {
          // Shape-level hyperlink only — a link inside a text run lives under
          // a:rPr and is just a link, not a video.
          const link = descendants(pic, 'hlinkClick')
            .find(el => el.parentNode && el.parentNode.localName === 'cNvPr');
          const rel = link && rels.get(relAttr(link, 'id'));
          if (!rel || !rel.external) continue;

          const hosted = hostedVideo(rel.target);
          if (!hosted) continue;
          kind = hosted.kind;
          url = hosted.embed;
          name = hosted.id;
        }
        if (!url) continue;

        // placement; a shape with no xfrm inherits it — fall back to full bleed
        const off = firstDescendant(pic, 'off');
        const ext = firstDescendant(pic, 'ext');
        const box = off && ext ? {
          x: Number(off.getAttribute('x')) / slideW,
          y: Number(off.getAttribute('y')) / slideH,
          w: Number(ext.getAttribute('cx')) / slideW,
          h: Number(ext.getAttribute('cy')) / slideH,
        } : { x: 0, y: 0, w: 1, h: 1 };

        found.push({ url, kind, name, ...box });
      }
    }
    media.push(found);
    if (blanks.length < media.length) blanks.push(false);
  }

  return { slideCount: slideParts.length, media, blanks, slideAspect: slideW / slideH };
}
