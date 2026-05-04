// skills/yome-skill-ppt/backends/sandbox/src/handlers.ts
//
// Sandbox backend for the ppt domain. Implements a subset of the ppt
// signature against an in-memory PptWorldState. Real macOS / iOS / node
// backends live in sibling backends/ directories and are not loaded by the
// hub sandbox runtime.
//
// Contract (matches existing handlers in web/lib/sandbox/handlers/*):
//   handlePpt(world, cmd) -> { stdout, stderr, exitCode }
// where `world` exposes `scratch.skillStates.ppt: PptWorldState` and `cmd`
// is the parsed command (positionals + flags).

import {
  createInitialPptState,
  getActiveFile,
  newBlankFile,
  setActiveFile,
  type PptAutoShapeType,
  type PptChart,
  type PptChartType,
  type PptFile,
  type PptShape,
  type PptSlide,
  type PptTable,
  type PptTableCell,
  type PptWorldState,
} from './state';

const VALID_AUTO_SHAPE_TYPES: ReadonlyArray<PptAutoShapeType> = [
  'rectangle', 'oval', 'roundedRect', 'triangle', 'rightArrow',
  'star5', 'pentagon', 'diamond', 'hexagon', 'cloud', 'lightningBolt', 'heart',
];

function normalizeAutoShapeType(raw: string | undefined): PptAutoShapeType {
  const t = (raw ?? 'rectangle').toLowerCase();
  const canonicalMap: Record<string, PptAutoShapeType> = {
    rectangle: 'rectangle',
    oval: 'oval',
    roundedrect: 'roundedRect',
    roundedrectangle: 'roundedRect',
    triangle: 'triangle',
    rightarrow: 'rightArrow',
    star5: 'star5',
    star: 'star5',
    pentagon: 'pentagon',
    diamond: 'diamond',
    hexagon: 'hexagon',
    cloud: 'cloud',
    lightningbolt: 'lightningBolt',
    heart: 'heart',
  };
  return canonicalMap[t] ?? 'rectangle';
}

const VALID_CHART_TYPES: ReadonlyArray<PptChartType> = ['column', 'bar', 'line', 'pie', 'area', 'scatter'];

function normalizeChartType(raw: string | undefined): PptChartType {
  const t = (raw ?? 'column').toLowerCase();
  return (VALID_CHART_TYPES as readonly string[]).includes(t) ? (t as PptChartType) : 'column';
}

/** Parse CSV; returns null if header row + at least 1 data row not present. */
function parseCsvChartData(raw: string): PptChart | null {
  const text = raw
    .replace(/\\r\\n/g, '\n')
    .replace(/\\n/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();
  if (!text) return null;
  const lines = text.split('\n').filter(l => l.length > 0);
  if (lines.length < 2) return null;
  const header = lines[0].split(',').map(s => s.trim());
  const dataRows = lines.slice(1).map(l => l.split(',').map(s => s.trim()));
  const seriesNames = header.slice(1);
  const categories = dataRows.map(r => r[0] ?? '');
  const series = seriesNames.map((name, si) => ({
    name,
    values: dataRows.map(r => {
      const v = Number.parseFloat(r[si + 1] ?? '0');
      return Number.isFinite(v) ? v : 0;
    }),
  }));
  return { type: 'column', categories, series };
}

export interface ParsedCommand {
  domain: string;
  action: string;
  positionals: string[];
  flags: Record<string, string>;
}

export interface HandlerResult {
  stdout: string;
  stderr: string;
  exitCode: number;
}

const ok = (stdout: string): HandlerResult => ({ stdout, stderr: '', exitCode: 0 });
const err = (stderr: string, code = 1): HandlerResult => ({ stdout: '', stderr, exitCode: code });

interface PptCarrier {
  scratch: { skillStates?: { ppt?: PptWorldState } & Record<string, unknown> } & Record<string, unknown>;
}

function getState(world: PptCarrier): PptWorldState {
  if (!world.scratch.skillStates) world.scratch.skillStates = {};
  if (!world.scratch.skillStates.ppt) {
    world.scratch.skillStates.ppt = createInitialPptState();
  }
  return world.scratch.skillStates.ppt!;
}

function findFile(state: PptWorldState, path: string): PptFile | undefined {
  return state.openFiles.find(f => f.path === path);
}

function nextSlideIndex(file: PptFile): number {
  return file.slides.length + 1;
}

function nextShapeIndex(slide: PptSlide): number {
  return slide.shapes.length + 1;
}

function reindexSlides(file: PptFile): void {
  file.slides.forEach((s, i) => { s.index = i + 1; });
}

function parseIntOrNull(s: string | undefined): number | null {
  if (s === undefined || s === '') return null;
  const n = Number.parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}

function parseBool(s: string | undefined): boolean {
  return s === 'true' || s === '1' || s === '';
}

export function handlePpt(world: PptCarrier, cmd: ParsedCommand): HandlerResult {
  const state = getState(world);

  switch (cmd.action) {
    case 'open': {
      const path = cmd.positionals[0] ?? cmd.flags.path;
      if (!path) return err('ppt open: missing <path>');
      let file = findFile(state, path);
      if (!file) {
        // Sandbox approximation: opening an unknown path creates an empty deck.
        file = { path, slides: [], active: true };
        state.openFiles.push(file);
      }
      setActiveFile(state, path);
      return ok(JSON.stringify({ ok: true, opened: path, slides: file.slides.length }));
    }

    case 'new': {
      const path = cmd.positionals[0] ?? cmd.flags.path ?? `~/Desktop/Untitled-${state.openFiles.length + 1}.pptx`;
      if (findFile(state, path) && cmd.flags.force !== 'true') {
        return err(`ppt new: ${path} already open (use --force)`);
      }
      const file = newBlankFile(path);
      state.openFiles = state.openFiles.filter(f => f.path !== path);
      state.openFiles.push(file);
      setActiveFile(state, path);
      return ok(JSON.stringify({ ok: true, created: path, slides: 1 }));
    }

    case 'save': {
      const file = getActiveFile(state);
      if (!file) return err('ppt save: no active presentation');
      const newPath = cmd.flags.path;
      if (newPath) {
        if (state.openFiles.some(f => f.path === newPath) && cmd.flags.force !== 'true') {
          return err(`ppt save: ${newPath} already exists (use --force)`);
        }
        file.path = newPath;
      }
      file.dirty = false;
      return ok(JSON.stringify({ ok: true, path: file.path }));
    }

    case 'close': {
      const file = getActiveFile(state);
      if (!file) return err('ppt close: no active presentation');
      const shouldSave = cmd.flags.save !== 'false';
      if (shouldSave) file.dirty = false;
      state.openFiles = state.openFiles.filter(f => f.path !== file.path);
      const next = state.openFiles[0];
      if (next) next.active = true;
      return ok(JSON.stringify({ ok: true, closed: file.path, saved: shouldSave }));
    }

    case 'files': {
      if (state.openFiles.length === 0) return ok('name\tslideCount');
      const lines = ['name\tslideCount'];
      for (const f of state.openFiles) lines.push(`${f.path}\t${f.slides.length}`);
      return ok(lines.join('\n'));
    }

    case 'slides': {
      const file = getActiveFile(state);
      if (!file) return err('ppt slides: no active presentation');
      const lines = ['index\ttitle\tshapes\tlayout'];
      for (const s of file.slides) {
        lines.push([s.index, s.title ?? '', s.shapes.length, s.layout ?? ''].join('\t'));
      }
      return ok(lines.join('\n'));
    }

    case 'slide.add': {
      const file = getActiveFile(state);
      if (!file) return err('ppt slide.add: no active presentation');
      const layoutN = parseIntOrNull(cmd.flags.layout);
      const layout = layoutN ?? 6;
      const at = parseIntOrNull(cmd.flags.index);
      const slide: PptSlide = { index: 0, layout, shapes: [] };
      if (at && at >= 1 && at <= file.slides.length + 1) {
        file.slides.splice(at - 1, 0, slide);
      } else {
        file.slides.push(slide);
      }
      reindexSlides(file);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, index: slide.index, layout, totalSlides: file.slides.length }));
    }

    case 'slide.delete': {
      const file = getActiveFile(state);
      if (!file) return err('ppt slide.delete: no active presentation');
      const idx = parseIntOrNull(cmd.positionals[0]);
      if (!idx) return err('ppt slide.delete: missing <index>');
      const slide = file.slides[idx - 1];
      if (!slide) return err(`ppt slide.delete: slide ${idx} not found`);
      if (slide.shapes.length > 0 && cmd.flags.force !== 'true') {
        return err(`ppt slide.delete: slide ${idx} has ${slide.shapes.length} shapes (use --force)`);
      }
      file.slides.splice(idx - 1, 1);
      reindexSlides(file);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, deleted: idx, totalSlides: file.slides.length }));
    }

    case 'slide.move': {
      const file = getActiveFile(state);
      if (!file) return err('ppt slide.move: no active presentation');
      const from = parseIntOrNull(cmd.positionals[0]);
      const to = parseIntOrNull(cmd.flags.to);
      if (!from || !to) return err('ppt slide.move: missing <index> or --to');
      if (from < 1 || from > file.slides.length) return err(`ppt slide.move: invalid <index> ${from}`);
      const [moved] = file.slides.splice(from - 1, 1);
      const insertAt = Math.max(0, Math.min(file.slides.length, to - 1));
      file.slides.splice(insertAt, 0, moved);
      reindexSlides(file);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, from, to: moved.index }));
    }

    case 'slide.duplicate': {
      const file = getActiveFile(state);
      if (!file) return err('ppt slide.duplicate: no active presentation');
      const from = parseIntOrNull(cmd.positionals[0]);
      if (!from || from < 1 || from > file.slides.length) {
        return err(`ppt slide.duplicate: invalid <index> ${from ?? ''}`);
      }
      const src = file.slides[from - 1];
      const clone: PptSlide = {
        index: 0,
        title: src.title,
        layout: src.layout,
        notes: src.notes,
        shapes: src.shapes.map(s => ({ ...s, chart: s.chart ? { ...s.chart, series: s.chart.series.map(x => ({ ...x, values: [...x.values] })) } : undefined })),
      };
      const toTarget = parseIntOrNull(cmd.flags.to);
      const insertAt = toTarget && toTarget >= 1
        ? Math.min(file.slides.length, toTarget - 1)
        : from;
      file.slides.splice(insertAt, 0, clone);
      reindexSlides(file);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, duplicated: from, at: clone.index }));
    }

    case 'get': {
      const file = getActiveFile(state);
      if (!file) return err('ppt get: no active presentation');
      const idx = parseIntOrNull(cmd.positionals[0]);
      if (!idx) return err('ppt get: missing <slide_index>');
      const slide = file.slides[idx - 1];
      if (!slide) return err(`ppt get: slide ${idx} not found`);
      const lines = ['index\tname\thasText\tcontent'];
      for (const sh of slide.shapes) {
        lines.push([sh.index, sh.name ?? '', sh.text ? 'true' : 'false', sh.text ?? ''].join('\t'));
      }
      return ok(lines.join('\n'));
    }

    case 'read': {
      const file = getActiveFile(state);
      if (!file) return err('ppt read: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt read: need <slide_index> and --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt read: shape ${slideIdx}/${shapeIdx} not found`);
      const obj: Record<string, unknown> = { text: shape.text ?? '' };
      if (shape.bold) obj.bold = 'true';
      if (shape.italic) obj.italic = 'true';
      if (shape.fontSize) obj.fontSize = shape.fontSize;
      return ok(JSON.stringify(obj));
    }

    case 'set': {
      const file = getActiveFile(state);
      if (!file) return err('ppt set: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      const text = cmd.flags.text;
      if (!slideIdx || !shapeIdx || text === undefined) return err('ppt set: need <slide_index> --shape --text');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt set: shape ${slideIdx}/${shapeIdx} not found`);
      shape.text = text;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx }));
    }

    case 'addtext': {
      const file = getActiveFile(state);
      if (!file) return err('ppt addtext: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const text = cmd.flags.text;
      if (!slideIdx || text === undefined) return err('ppt addtext: need <slide_index> --text');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt addtext: slide ${slideIdx} not found`);
      const shape = {
        index: nextShapeIndex(slide),
        name: 'TextBox',
        text,
        left: parseIntOrNull(cmd.flags.left) ?? 100,
        top: parseIntOrNull(cmd.flags.top) ?? 200,
        width: parseIntOrNull(cmd.flags.width) ?? 400,
        height: parseIntOrNull(cmd.flags.height) ?? 50,
        fontSize: parseIntOrNull(cmd.flags.size) ?? undefined,
        color: cmd.flags.color,
        bold: parseBool(cmd.flags.bold) && cmd.flags.bold !== undefined,
        italic: parseBool(cmd.flags.italic) && cmd.flags.italic !== undefined,
      };
      slide.shapes.push(shape);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shape.index }));
    }

    case 'title': {
      const file = getActiveFile(state);
      if (!file) return err('ppt title: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const text = cmd.flags.text;
      if (!slideIdx || text === undefined) return err('ppt title: need <slide_index> --text');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt title: slide ${slideIdx} not found`);
      let titleShape = slide.shapes[0];
      if (!titleShape) {
        titleShape = { index: 1, name: 'Title Placeholder' };
        slide.shapes.push(titleShape);
      }
      titleShape.text = text;
      slide.title = text;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, title: text }));
    }

    case 'fmt': {
      const file = getActiveFile(state);
      if (!file) return err('ppt fmt: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt fmt: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt fmt: shape ${slideIdx}/${shapeIdx} not found`);
      const applied: Record<string, unknown> = {};
      if (cmd.flags.bold !== undefined) { shape.bold = parseBool(cmd.flags.bold); applied.bold = shape.bold; }
      if (cmd.flags.italic !== undefined) { shape.italic = parseBool(cmd.flags.italic); applied.italic = shape.italic; }
      const sz = parseIntOrNull(cmd.flags.size);
      if (sz !== null) { shape.fontSize = sz; applied.size = sz; }
      if (cmd.flags.color !== undefined) { shape.color = cmd.flags.color; applied.color = shape.color; }
      if (cmd.flags.bg !== undefined) { shape.bg = cmd.flags.bg; applied.bg = shape.bg; }
      if (cmd.flags.align !== undefined) {
        const a = cmd.flags.align as 'left' | 'center' | 'right';
        shape.align = a;
        applied.align = a;
      }
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx, applied }));
    }

    case 'export': {
      const file = getActiveFile(state);
      if (!file) return err('ppt export: no active presentation');
      const format = cmd.flags.format;
      const path = cmd.flags.path;
      if (!format || !path) return err('ppt export: need --format and --path');
      return ok(JSON.stringify({ ok: true, exported: path, format, slides: file.slides.length }));
    }

    case 'shape.delete': {
      const file = getActiveFile(state);
      if (!file) return err('ppt shape.delete: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt shape.delete: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt shape.delete: slide ${slideIdx} not found`);
      const shape = slide.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt shape.delete: shape ${slideIdx}/${shapeIdx} not found`);
      slide.shapes.splice(shapeIdx - 1, 1);
      slide.shapes.forEach((s, i) => { s.index = i + 1; });
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx }));
    }

    case 'shape.move': {
      const file = getActiveFile(state);
      if (!file) return err('ppt shape.move: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt shape.move: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt shape.move: shape ${slideIdx}/${shapeIdx} not found`);
      const left = parseIntOrNull(cmd.flags.left);
      const top = parseIntOrNull(cmd.flags.top);
      if (left === null && top === null) return err('ppt shape.move: need --left and/or --top');
      if (left !== null) shape.left = left;
      if (top !== null) shape.top = top;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx, left: shape.left, top: shape.top }));
    }

    case 'shape.resize': {
      const file = getActiveFile(state);
      if (!file) return err('ppt shape.resize: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt shape.resize: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt shape.resize: shape ${slideIdx}/${shapeIdx} not found`);
      const w = parseIntOrNull(cmd.flags.width);
      const h = parseIntOrNull(cmd.flags.height);
      if (w === null && h === null) return err('ppt shape.resize: need --width and/or --height');
      if (w !== null) shape.width = w;
      if (h !== null) shape.height = h;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx, width: shape.width, height: shape.height }));
    }

    case 'shape.add': {
      const file = getActiveFile(state);
      if (!file) return err('ppt shape.add: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const rawType = cmd.flags.type;
      if (!slideIdx || !rawType) return err('ppt shape.add: need <slide_index> --type');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt shape.add: slide ${slideIdx} not found`);
      const autoShapeType = normalizeAutoShapeType(rawType);
      const shape: PptShape = {
        index: nextShapeIndex(slide),
        name: `AutoShape:${autoShapeType}`,
        autoShapeType,
        left: parseIntOrNull(cmd.flags.left) ?? 100,
        top: parseIntOrNull(cmd.flags.top) ?? 100,
        width: parseIntOrNull(cmd.flags.width) ?? 200,
        height: parseIntOrNull(cmd.flags.height) ?? 100,
      };
      if (cmd.flags.text) shape.text = cmd.flags.text;
      slide.shapes.push(shape);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shape.index, type: autoShapeType }));
    }

    case 'picture.add': {
      const file = getActiveFile(state);
      if (!file) return err('ppt picture.add: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const path = cmd.flags.path;
      if (!slideIdx || !path) return err('ppt picture.add: need <slide_index> --path');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt picture.add: slide ${slideIdx} not found`);
      const shape: PptShape = {
        index: nextShapeIndex(slide),
        name: 'Picture',
        picturePath: path,
        left: parseIntOrNull(cmd.flags.left) ?? 100,
        top: parseIntOrNull(cmd.flags.top) ?? 100,
        width: parseIntOrNull(cmd.flags.width) ?? undefined,
        height: parseIntOrNull(cmd.flags.height) ?? undefined,
      };
      slide.shapes.push(shape);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shape.index, path }));
    }

    case 'notes.set': {
      const file = getActiveFile(state);
      if (!file) return err('ppt notes.set: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const text = cmd.flags.text;
      if (!slideIdx || text === undefined) return err('ppt notes.set: need <slide_index> --text');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt notes.set: slide ${slideIdx} not found`);
      slide.notes = text;
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx }));
    }

    case 'notes.get': {
      const file = getActiveFile(state);
      if (!file) return err('ppt notes.get: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      if (!slideIdx) return err('ppt notes.get: need <slide_index>');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt notes.get: slide ${slideIdx} not found`);
      return ok(slide.notes ?? '');
    }

    case 'table.add': {
      const file = getActiveFile(state);
      if (!file) return err('ppt table.add: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const rows = parseIntOrNull(cmd.flags.rows);
      const cols = parseIntOrNull(cmd.flags.cols);
      if (!slideIdx || !rows || !cols) return err('ppt table.add: need <slide_index> --rows --cols');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt table.add: slide ${slideIdx} not found`);
      const table: PptTable = { rows, cols, cells: [] };
      const shape: PptShape = {
        index: nextShapeIndex(slide),
        name: 'Table',
        table,
        left: parseIntOrNull(cmd.flags.left) ?? 100,
        top: parseIntOrNull(cmd.flags.top) ?? 100,
        width: parseIntOrNull(cmd.flags.width) ?? 600,
        height: parseIntOrNull(cmd.flags.height) ?? 200,
      };
      slide.shapes.push(shape);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shape.index, rows, cols }));
    }

    case 'table.set': {
      const file = getActiveFile(state);
      if (!file) return err('ppt table.set: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      const r = parseIntOrNull(cmd.flags.row);
      const c = parseIntOrNull(cmd.flags.col);
      const text = cmd.flags.text;
      if (!slideIdx || !shapeIdx || !r || !c || text === undefined) {
        return err('ppt table.set: need <slide_index> --shape --row --col --text');
      }
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt table.set: shape ${slideIdx}/${shapeIdx} not found`);
      if (!shape.table) return err(`ppt table.set: shape ${slideIdx}/${shapeIdx} is not a table`);
      if (r < 1 || r > shape.table.rows) return err(`ppt table.set: row ${r} out of range`);
      if (c < 1 || c > shape.table.cols) return err(`ppt table.set: col ${c} out of range`);
      const existing = shape.table.cells.find(x => x.row === r && x.col === c);
      if (existing) existing.text = text;
      else shape.table.cells.push({ row: r, col: c, text });
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx, row: r, col: c }));
    }

    case 'charts': {
      const file = getActiveFile(state);
      if (!file) return err('ppt charts: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      if (!slideIdx) return err('ppt charts: missing <slide_index>');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt charts: slide ${slideIdx} not found`);
      const lines = ['shape\ttype\ttitle\tseries'];
      for (const sh of slide.shapes) {
        if (sh.chart) {
          lines.push([sh.index, sh.chart.type, sh.chart.title ?? '', sh.chart.series.length].join('\t'));
        }
      }
      return ok(lines.join('\n'));
    }

    case 'chart.add': {
      const file = getActiveFile(state);
      if (!file) return err('ppt chart.add: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      if (!slideIdx) return err('ppt chart.add: missing <slide_index>');
      const slide = file.slides[slideIdx - 1];
      if (!slide) return err(`ppt chart.add: slide ${slideIdx} not found`);
      const rawData = cmd.flags.data;
      if (!rawData) return err('ppt chart.add: missing --data');
      const chart = parseCsvChartData(rawData);
      if (!chart) return err('ppt chart.add: invalid --data CSV (need header + >=1 data row)');
      chart.type = normalizeChartType(cmd.flags.type);
      if (cmd.flags.title !== undefined && cmd.flags.title !== '') chart.title = cmd.flags.title;
      const shape = {
        index: nextShapeIndex(slide),
        name: 'Chart',
        left: parseIntOrNull(cmd.flags.left) ?? 80,
        top: parseIntOrNull(cmd.flags.top) ?? 120,
        width: parseIntOrNull(cmd.flags.width) ?? 480,
        height: parseIntOrNull(cmd.flags.height) ?? 300,
        chart,
      };
      slide.shapes.push(shape);
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shape.index, type: chart.type, series: chart.series.length }));
    }

    case 'chart.update': {
      const file = getActiveFile(state);
      if (!file) return err('ppt chart.update: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt chart.update: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt chart.update: shape ${slideIdx}/${shapeIdx} not found`);
      if (!shape.chart) return err(`ppt chart.update: shape ${slideIdx}/${shapeIdx} is not a chart`);
      const rawData = cmd.flags.data;
      if (!rawData) return err('ppt chart.update: missing --data');
      const parsed = parseCsvChartData(rawData);
      if (!parsed) return err('ppt chart.update: invalid --data CSV');
      shape.chart = { type: shape.chart.type, title: shape.chart.title, categories: parsed.categories, series: parsed.series };
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx, series: parsed.series.length }));
    }

    case 'chart.delete': {
      const file = getActiveFile(state);
      if (!file) return err('ppt chart.delete: no active presentation');
      const slideIdx = parseIntOrNull(cmd.positionals[0]);
      const shapeIdx = parseIntOrNull(cmd.flags.shape);
      if (!slideIdx || !shapeIdx) return err('ppt chart.delete: need <slide_index> --shape');
      const slide = file.slides[slideIdx - 1];
      const shape = slide?.shapes[shapeIdx - 1];
      if (!shape) return err(`ppt chart.delete: shape ${slideIdx}/${shapeIdx} not found`);
      if (!shape.chart) return err(`ppt chart.delete: shape ${slideIdx}/${shapeIdx} is not a chart`);
      slide.shapes.splice(shapeIdx - 1, 1);
      slide.shapes.forEach((s, i) => { s.index = i + 1; });
      file.dirty = true;
      return ok(JSON.stringify({ ok: true, slide: slideIdx, shape: shapeIdx }));
    }

    default:
      return err(`ppt ${cmd.action}: not supported in sandbox`);
  }
}
