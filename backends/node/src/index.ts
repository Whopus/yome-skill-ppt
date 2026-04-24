// skills/yome-skill-ppt/backends/node/src/index.ts
//
// Node backend for @yome/ppt. In the v0.1 monorepo phase the real
// implementation lives in Server/agent/commands/pptCommands.ts (executed by
// DeviceAgentRuntime). This entry point exists so the spec 4 layout is
// faithful and so the future CLI installer (Phase 2) can `import()` it via
// the manifest's delivery.node.entry pointer.
//
// TODO(spec-v0.1 / Phase 2): once cli/src/skills/runner/nodeBackend.ts is in
// place, re-export the real ppt dispatch here and remove the duplication
// from Server.

export const SKILL_DOMAIN = 'ppt';
export const SKILL_VERSION = '1.0.0';
export const IMPLEMENTS_SIGNATURE = '>=1.0.0 <2.0.0';

export interface NodeBackendDispatch {
  action: string;
  positionals: string[];
  flags: Record<string, string>;
}

export async function dispatch(_req: NodeBackendDispatch): Promise<{ ok: boolean; stdout?: string; stderr?: string }> {
  return {
    ok: false,
    stderr: 'ppt node backend: real implementation lives in Server/agent/commands/pptCommands.ts during v0.1 monorepo phase.',
  };
}
