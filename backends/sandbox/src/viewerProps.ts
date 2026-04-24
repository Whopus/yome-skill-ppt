// skills/yome-skill-ppt/backends/sandbox/src/viewerProps.ts
//
// Project a PptWorldState snapshot into the props the viewer (viewer/index.html)
// expects. Stays a pure function so the same projection can be applied at
// every trace step to drive seek().

import type { PptShape, PptSlide, PptWorldState } from './state';

export interface ViewerShape {
  index: number;
  text: string;
  bold: boolean;
  italic: boolean;
  fontSize?: number;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface ViewerSlide {
  index: number;
  title?: string;
  layout?: number;
  shapes: ViewerShape[];
}

export interface ViewerFile {
  path: string;
  active: boolean;
  dirty: boolean;
  slides: ViewerSlide[];
}

export interface ViewerProps {
  files: ViewerFile[];
  /** Path of the active file, or undefined if no file open. */
  activePath?: string;
}

function projectShape(s: PptShape): ViewerShape {
  return {
    index: s.index,
    text: s.text ?? '',
    bold: !!s.bold,
    italic: !!s.italic,
    fontSize: s.fontSize,
    color: s.color,
    bg: s.bg,
    align: s.align,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  };
}

function projectSlide(s: PptSlide): ViewerSlide {
  return {
    index: s.index,
    title: s.title,
    layout: s.layout,
    shapes: s.shapes.map(projectShape),
  };
}

export function toViewerProps(state: PptWorldState): ViewerProps {
  const files: ViewerFile[] = state.openFiles.map(f => ({
    path: f.path,
    active: !!f.active,
    dirty: !!f.dirty,
    slides: f.slides.map(projectSlide),
  }));
  const active = files.find(f => f.active);
  return { files, activePath: active?.path };
}
