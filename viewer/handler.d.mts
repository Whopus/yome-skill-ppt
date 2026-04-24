// Type declarations for the viewer's pure trace-replay engine.
// The implementation is plain JS (handler.mjs) so it loads in both the
// browser <script type="module"> tag and node-side vitest without any
// build step.

import type { PptWorldState } from '../backends/sandbox/src/state';

export interface ParsedCommand {
  domain: string;
  action: string;
  positionals: string[];
  flags: Record<string, string>;
}

export interface ApplyResult {
  ok: boolean;
  summary?: string;
  err?: string;
}

export function makeInitialState(): PptWorldState;
export function getActiveFile(state: PptWorldState): PptWorldState['openFiles'][number] | undefined;
export function parseCommand(cmd: string): ParsedCommand;
export function applyCommand(state: PptWorldState, parsed: ParsedCommand): ApplyResult;
