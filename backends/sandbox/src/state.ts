// skills/yome-skill-ppt/backends/sandbox/src/state.ts
//
// Virtual world state for the ppt domain. Drives sandbox replays / benchmarks
// in the Hub, and feeds the viewer (viewer/index.html) via postMessage.
//
// Per spec 3.6, this state lives in BenchmarkFixtures.skillStates.ppt for a
// case's t=0 snapshot, and the sandbox handler mutates a working copy as the
// trace plays out.

export interface PptShape {
  /** 1-based shape index within the slide */
  index: number;
  name?: string;
  text?: string;
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: string;
  bg?: string;
  /** Optional layout box in points */
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  align?: 'left' | 'center' | 'right';
}

export interface PptSlide {
  /** 1-based slide index */
  index: number;
  title?: string;
  layout?: number; // 1=title, 2=title+content, 5=title only, 6=blank
  shapes: PptShape[];
}

export interface PptFile {
  path: string;          // e.g. '~/Desktop/Q3-Review.pptx'
  slides: PptSlide[];
  dirty?: boolean;       // unsaved changes
  /** True if this file is the currently active presentation */
  active?: boolean;
}

export interface PptWorldState {
  openFiles: PptFile[];
}

export function createInitialPptState(seed?: Partial<PptWorldState>): PptWorldState {
  return {
    openFiles: seed?.openFiles ?? [],
  };
}

export function getActiveFile(state: PptWorldState): PptFile | undefined {
  return state.openFiles.find(f => f.active) ?? state.openFiles[0];
}

export function setActiveFile(state: PptWorldState, path: string): void {
  for (const f of state.openFiles) f.active = f.path === path;
}

/** Default new-presentation seed: 1 blank slide with title + subtitle placeholders. */
export function newBlankFile(path: string): PptFile {
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
