// skills/yome-skill-ppt/backends/sandbox/src/index.ts
//
// Public entry for the @yome/ppt sandbox backend. The hub MiniRuntime
// imports `register(api)` to wire this skill's command dispatcher into the
// sandbox; tests / CLI replay tools can also import { handlePpt, toViewerProps }
// directly.

export { handlePpt, type HandlerResult, type ParsedCommand } from './handlers';
export { toViewerProps, type ViewerProps } from './viewerProps';
export {
  createInitialPptState,
  newBlankFile,
  type PptFile,
  type PptShape,
  type PptSlide,
  type PptWorldState,
} from './state';

import { handlePpt, type HandlerResult, type ParsedCommand } from './handlers';
import { toViewerProps, type ViewerProps } from './viewerProps';
import type { PptWorldState } from './state';

/** Minimal API the hub MiniRuntime hands to skill backends at registration. */
export interface SkillSandboxApi {
  registerCommandHandler: (
    domain: string,
    handler: (carrier: { scratch: { skillStates?: Record<string, unknown> } & Record<string, unknown> }, cmd: ParsedCommand) => HandlerResult,
  ) => void;
  registerViewerProjector?: (
    domain: string,
    projector: (state: unknown) => unknown,
  ) => void;
}

export const SKILL_DOMAIN = 'ppt';
export const SKILL_VERSION = '1.0.0';
export const IMPLEMENTS_SIGNATURE = '>=1.0.0 <2.0.0';

/** Register the ppt skill with a hub-side MiniRuntime. */
export function register(api: SkillSandboxApi): void {
  api.registerCommandHandler(SKILL_DOMAIN, (carrier, cmd) => handlePpt(carrier, cmd));
  api.registerViewerProjector?.(SKILL_DOMAIN, (state: unknown): ViewerProps => {
    return toViewerProps((state ?? { openFiles: [] }) as PptWorldState);
  });
}
