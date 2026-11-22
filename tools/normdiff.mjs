/* Comparador de normativas FCN para Zona Nadadores CN Salou.
   Se ejecuta en una pestaña de https://www.natacio.cat (Claude in Chrome).
   Guarda en localStorage de natacio.cat el texto de cada normativa y, cuando la FCN la actualiza,
   devuelve las frases añadidas y eliminadas para que Claude escriba un resumen breve.
   Uso:
     const m = await import('https://cdn.jsdelivr.net/gh/CNSalou/zona-nadadores@<commit>/tools/normdiff.mjs');
     const r = await m.check();          // {nuevas, cambiadas:[{id,title,date,sec,added,removed}], baseline:[...]}
     m.part(r, 0)                         // trozo de ~900 caracteres del informe (para leerlo por partes)
     m.accept(id)                         // tras resumir: guarda la versión nueva como referencia
*/
const FEED = 'https://esb.optimalwayconsulting.com/fcnat/1/Se6zwx4E4nd8facPWIaM1knT2fh11tC9/DocumentUpdates/getSectionUpdates/';
const PDF = id => 'https://www.natacio.cat/wp-json/files/ow-getFile?file=' + id;
const K = id => 'cnsalou_snap_' + id, KP = id => 'cnsalou_snap_pending_' + id;
let lib = null;
async function pdfjs() {
  if (lib) return lib;
  lib = await import('https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs');
  lib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';
  return lib;
}
export async function feed(sections = [2, 7], n = 40, since = '2026-08-01') {
  const out = [];
  for (const s of sections) {
    const t = await fetch(FEED + s + '/' + n, { cache: 'no-store' }).then(r => r.text());
    const j = JSON.parse(decodeURIComponent(escape(atob(t.trim()))));
    (j.messageData || []).forEach(x => {
      if (x.docId && /normativa|calendari|reglament/i.test(x.title || '') && String(x.datePublished || '') >= since)
        out.push({ id: String(x.docId), title: String(x.title).trim(), ts: String(x.datePublished), date: String(x.datePublished).slice(0, 10), sec: s === 7 ? 'Màsters' : 'Natació' });
    });
  }
  const seen = {};
  return out.sort((a, b) => b.ts.localeCompare(a.ts)).filter(x => seen[x.id] ? false : (seen[x.id] = 1));
}
export async function text(id) {
  const L = await pdfjs();
  const buf = await fetch(PDF(id)).then(r => r.arrayBuffer());
  const doc = await L.getDocument({ data: buf }).promise;
  let s = '';
  for (let p = 1; p <= doc.numPages; p++) {
    const c = await (await doc.getPage(p)).getTextContent();
    s += '\n' + c.items.map(i => i.str + (i.hasEOL ? '\n' : '')).join(' ');
  }
  return s.replace(/[ \t]+/g, ' ')
    .replace(/(COMITÈ DE [A-ZÀ-Ú ]+)?\s*Normativa[ _A-Z0-9\-]*\s*Temporada\s*20\d\d\s*-?\s*20?\d\d\s*Act\.?\s*[\d./ ]+\s*pàg\s*\d+\s*de\s*\d+/gi, ' ')
    .trim();
}
const sentences = t => t.replace(/\s+/g, ' ').split(/(?<=[.;:])\s+(?=[A-ZÀ-Ú0-9(])|\s(?=\d{1,2}\.\d{1,2}\.\s)/).map(x => x.trim()).filter(x => x.length > 3);
const norm = x => x.toLowerCase().replace(/\s+/g, ' ').replace(/[’'´]/g, "'").trim();
export function diff(oldT, newT) {
  const A = sentences(oldT), B = sentences(newT), sa = new Set(A.map(norm)), sb = new Set(B.map(norm));
  return { added: B.filter(x => !sa.has(norm(x))), removed: A.filter(x => !sb.has(norm(x))) };
}
function get(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } }
function put(k, v) { localStorage.setItem(k, JSON.stringify(v)); }
/* Revisa todas las normativas del listado. Las que no tienen referencia guardada se guardan (baseline).
   Las que tienen una fecha de publicación más nueva que la guardada se comparan. */
export async function check(opts = {}) {
  const items = await feed(opts.sections, opts.n, opts.since);
  const r = { revisadas: items.length, cambiadas: [], baseline: [], sinCambiosTexto: [] };
  for (const it of items) {
    const snap = get(K(it.id));
    if (snap && snap.ts >= it.ts) continue;
    const t = await text(it.id);
    if (!snap) { put(K(it.id), { ts: it.ts, title: it.title, t }); r.baseline.push(it.id + ' ' + it.title); continue; }
    const d = diff(snap.t, t);
    if (!d.added.length && !d.removed.length) { put(K(it.id), { ts: it.ts, title: it.title, t }); r.sinCambiosTexto.push(it); continue; }
    put(KP(it.id), { ts: it.ts, title: it.title, t });
    r.cambiadas.push(Object.assign({}, it, { added: d.added.slice(0, 40), removed: d.removed.slice(0, 40) }));
  }
  window.__normdiff = r;
  return { revisadas: r.revisadas, cambiadas: r.cambiadas.map(c => ({ id: c.id, date: c.date, title: c.title, añadidas: c.added.length, eliminadas: c.removed.length })), baseline: r.baseline.length, sinCambiosTexto: r.sinCambiosTexto.map(x => x.id + ' ' + x.date) };
}
/* Informe de una normativa cambiada, en trozos legibles */
export function report(id) {
  const c = (window.__normdiff || { cambiadas: [] }).cambiadas.find(x => x.id === String(id));
  if (!c) return '';
  const cut = s => s.length > 400 ? s.slice(0, 400) + '…' : s;
  return c.title + ' (' + c.date + ')\n+ AÑADIDO:\n' + c.added.map(s => '+ ' + cut(s)).join('\n') + '\n- ELIMINADO:\n' + c.removed.map(s => '- ' + cut(s)).join('\n');
}
export function part(id, i = 0, size = 900) { return report(id).slice(i * size, (i + 1) * size); }
export function accept(id) { const p = get(KP(id)); if (p) { put(K(id), p); localStorage.removeItem(KP(id)); return true; } return false; }
export function status() {
  return Object.keys(localStorage).filter(k => k.startsWith('cnsalou_snap_') && !k.includes('pending')).map(k => { const v = get(k); return k.slice(13) + ' ' + (v && v.ts); });
}
