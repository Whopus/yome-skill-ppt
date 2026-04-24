// skills/yome-skill-ppt/viewer/handler.mjs
//
// Pure trace-replay engine for the @yome/ppt viewer.
// This is the SHARED source of truth used by both:
//   1. viewer/index.html  — loaded as <script type="module" src="handler.mjs">
//   2. tests              — imported by Server/agent/compress/viewerHandler.test.ts
//
// The viewer never executes real commands; it only walks a pre-recorded
// trace and visualises what the PPT state SHOULD look like at each step.
// To stay aligned with the authoritative sandbox backend
// (skills/yome-skill-ppt/backends/sandbox/src/handlers.ts), we expose the
// same mutation surface: parsed-command in → state mutated.
//
// API:
//   parseCommand(cmd: string) -> { domain, action, positionals, flags }
//   applyCommand(state, parsed) -> { ok, summary?, err? }
//   makeInitialState() -> PptWorldState
//   getActiveFile(state) -> PptFile | undefined
//
// This module has NO dom dependency and runs in both browser and node.

export function makeInitialState() {
  return { openFiles: [] };
}

export function getActiveFile(state) {
  return state.openFiles.find(f => f.active) || state.openFiles[0];
}

function setActiveFile(state, path) {
  state.openFiles.forEach(f => { f.active = f.path === path; });
}

function findFile(state, path) {
  return state.openFiles.find(f => f.path === path);
}

function reindex(file) {
  file.slides.forEach((s, i) => { s.index = i + 1; });
}

function newBlankFile(path) {
  return {
    path,
    dirty: true,
    active: true,
    slides: [
      {
        index: 1,
        layout: 1,
        shapes: [
          { index: 1, name: 'Title Placeholder' },
          { index: 2, name: 'Subtitle Placeholder' },
        ],
      },
    ],
  };
}

function parseIntOrNull(s) {
  if (s === undefined || s === null || s === '') return null;
  const n = parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}

function parseBool(s) {
  return s === 'true' || s === '1' || s === '';
}

function trunc(s, n) {
  n = n || 30;
  return s && s.length > n ? s.slice(0, n) + '…' : (s || '');
}

export function parseCommand(cmd) {
  const stripped = (cmd || '').replace(/^@\S+\s+/, '').trim();
  const tokens = [];
  let current = '';
  let inDQ = false;
  let inSQ = false;
  for (let i = 0; i < stripped.length; i++) {
    const ch = stripped[i];
    if (ch === '"' && !inSQ) { inDQ = !inDQ; current += ch; continue; }
    if (ch === "'" && !inDQ) { inSQ = !inSQ; current += ch; continue; }
    if (/\s/.test(ch) && !inDQ && !inSQ) {
      if (current) { tokens.push(current); current = ''; }
      continue;
    }
    current += ch;
  }
  if (current) tokens.push(current);

  const domain = tokens.shift() || '';
  const action = tokens.shift() || '';
  const positionals = [];
  const flags = {};
  for (const tok of tokens) {
    if (tok.startsWith('--')) {
      const eq = tok.indexOf('=');
      if (eq > 0) {
        const k = tok.slice(2, eq);
        let v = tok.slice(eq + 1);
        if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
          v = v.slice(1, -1);
        }
        flags[k] = v;
      } else {
        flags[tok.slice(2)] = '';
      }
    } else {
      let t = tok;
      if ((t.startsWith('"') && t.endsWith('"')) || (t.startsWith("'") && t.endsWith("'"))) {
        t = t.slice(1, -1);
      }
      positionals.push(t);
    }
  }
  return { domain, action, positionals, flags };
}

export function applyCommand(state, parsed) {
  const action = parsed.action;
  const pos = parsed.positionals || [];
  const fl  = parsed.flags || {};

  if (action === 'open') {
    const path = pos[0] || fl.path;
    if (!path) return { ok: false, err: 'open: missing path' };
    let f = findFile(state, path);
    if (!f) {
      f = { path, slides: [], active: true };
      state.openFiles.push(f);
    }
    setActiveFile(state, path);
    return { ok: true, summary: 'opened ' + path };
  }
  if (action === 'new') {
    const p = pos[0] || fl.path || ('~/Desktop/Untitled-' + (state.openFiles.length + 1) + '.pptx');
    state.openFiles = state.openFiles.filter(f => f.path !== p);
    state.openFiles.push(newBlankFile(p));
    setActiveFile(state, p);
    return { ok: true, summary: 'new ' + p };
  }
  if (action === 'save') {
    const sf = getActiveFile(state);
    if (!sf) return { ok: false, err: 'save: no active' };
    if (fl.path) sf.path = fl.path;
    sf.dirty = false;
    return { ok: true, summary: 'saved ' + sf.path };
  }
  if (action === 'close') {
    const cf = getActiveFile(state);
    if (!cf) return { ok: false, err: 'close: no active' };
    state.openFiles = state.openFiles.filter(f => f.path !== cf.path);
    const nx = state.openFiles[0];
    if (nx) nx.active = true;
    return { ok: true, summary: 'closed ' + cf.path };
  }
  if (action === 'files') return { ok: true, summary: state.openFiles.length + ' open' };

  const af = getActiveFile(state);
  if (!af) return { ok: false, err: action + ': no active presentation' };

  if (action === 'slides') return { ok: true, summary: af.slides.length + ' slides' };

  if (action === 'slide.add') {
    const layout = parseIntOrNull(fl.layout);
    const lay = layout || 6;
    const at = parseIntOrNull(fl.index);
    const slide = { index: 0, layout: lay, shapes: [] };
    if (at && at >= 1 && at <= af.slides.length + 1) af.slides.splice(at - 1, 0, slide);
    else af.slides.push(slide);
    reindex(af);
    af.dirty = true;
    return { ok: true, summary: '+slide @' + slide.index };
  }
  if (action === 'slide.delete') {
    const di = parseIntOrNull(pos[0]);
    if (!di) return { ok: false, err: 'slide.delete: missing index' };
    const s = af.slides[di - 1];
    if (!s) return { ok: false, err: 'slide ' + di + ' not found' };
    if (s.shapes.length > 0 && fl.force !== 'true') return { ok: false, err: 'use --force' };
    af.slides.splice(di - 1, 1);
    reindex(af);
    af.dirty = true;
    return { ok: true, summary: '-slide @' + di };
  }
  if (action === 'slide.move') {
    const from = parseIntOrNull(pos[0]);
    const to = parseIntOrNull(fl.to);
    if (!from || !to) return { ok: false, err: 'slide.move: need <index> --to' };
    const moved = af.slides.splice(from - 1, 1)[0];
    af.slides.splice(Math.max(0, Math.min(af.slides.length, to - 1)), 0, moved);
    reindex(af);
    af.dirty = true;
    return { ok: true, summary: 'mv ' + from + '→' + to };
  }
  if (action === 'set') {
    const si = parseIntOrNull(pos[0]);
    const sh = parseIntOrNull(fl.shape);
    const t = fl.text;
    if (!si || !sh || t === undefined) return { ok: false, err: 'set: need <slide> --shape --text' };
    const sl = af.slides[si - 1];
    const sp = sl && sl.shapes[sh - 1];
    if (!sp) return { ok: false, err: 'shape ' + si + '/' + sh + ' not found' };
    sp.text = t;
    af.dirty = true;
    return { ok: true, summary: 'set ' + si + '/' + sh };
  }
  if (action === 'addtext') {
    const asi = parseIntOrNull(pos[0]);
    const atx = fl.text;
    if (!asi || atx === undefined) return { ok: false, err: 'addtext: need <slide> --text' };
    const asl = af.slides[asi - 1];
    if (!asl) return { ok: false, err: 'slide ' + asi + ' not found' };
    const asp = {
      index: asl.shapes.length + 1,
      name: 'TextBox',
      text: atx,
      left: parseIntOrNull(fl.left) || 100,
      top: parseIntOrNull(fl.top) || 200,
      width: parseIntOrNull(fl.width) || 400,
      height: parseIntOrNull(fl.height) || 50,
      fontSize: parseIntOrNull(fl.size) || undefined,
      color: fl.color,
      bold: fl.bold !== undefined && parseBool(fl.bold),
      italic: fl.italic !== undefined && parseBool(fl.italic),
    };
    asl.shapes.push(asp);
    af.dirty = true;
    return { ok: true, summary: '+text ' + asi + '/' + asp.index };
  }
  if (action === 'title') {
    const tsi = parseIntOrNull(pos[0]);
    const tt = fl.text;
    if (!tsi || tt === undefined) return { ok: false, err: 'title: need <slide> --text' };
    const tsl = af.slides[tsi - 1];
    if (!tsl) return { ok: false, err: 'slide ' + tsi + ' not found' };
    let th = tsl.shapes[0];
    if (!th) { th = { index: 1, name: 'Title Placeholder' }; tsl.shapes.push(th); }
    th.text = tt;
    tsl.title = tt;
    af.dirty = true;
    return { ok: true, summary: 'title ' + tsi + ': ' + trunc(tt, 24) };
  }
  if (action === 'fmt') {
    const fsi = parseIntOrNull(pos[0]);
    const fsh = parseIntOrNull(fl.shape);
    if (!fsi || !fsh) return { ok: false, err: 'fmt: need <slide> --shape' };
    const fsl = af.slides[fsi - 1];
    const fsp = fsl && fsl.shapes[fsh - 1];
    if (!fsp) return { ok: false, err: 'shape ' + fsi + '/' + fsh + ' not found' };
    if (fl.bold !== undefined) fsp.bold = parseBool(fl.bold);
    if (fl.italic !== undefined) fsp.italic = parseBool(fl.italic);
    const sz = parseIntOrNull(fl.size);
    if (sz !== null) fsp.fontSize = sz;
    if (fl.color !== undefined) fsp.color = fl.color;
    if (fl.bg !== undefined) fsp.bg = fl.bg;
    if (fl.align !== undefined) fsp.align = fl.align;
    af.dirty = true;
    return { ok: true, summary: 'fmt ' + fsi + '/' + fsh };
  }
  if (action === 'export') return { ok: !!(fl.format && fl.path), summary: 'export ' + fl.format };
  if (action === 'get' || action === 'read') return { ok: true, summary: action };

  return { ok: false, err: 'unsupported action: ' + action };
}
