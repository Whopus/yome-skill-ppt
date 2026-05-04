// skills/yome-skill-ppt/backends/sandbox/src/compress.ts
//
// Tool-result compression functions for the ppt domain. Spec 4.4 says the
// signature.json `compress` field's `module` should point at this file
// (backends/sandbox/ or backends/node/).
//
// During the v0.1 monorepo phase, the runtime registry in
// Server/agent/compress/registry.ts statically imports from
// Server/agent/compress/ppt.ts. That file currently holds the same logic;
// once the registry can resolve into skill-repo paths (planned with the
// CLI's skill installer in Phase 2), this file becomes the authoritative
// source and Server/agent/compress/ppt.ts will re-export from here.

function trunc(s: string, n = 30): string {
  return s.length > n ? s.slice(0, n) + '...[truncated]' : s;
}

export function compressFiles(content: string): string {
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('name');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无打开文稿';
  return `[compressed] ${rows.length}个文稿:\n` + rows.map(r => {
    const c = r.split('\t');
    return `${trunc(c[0] || '')} ${c[1] || ''}张`;
  }).join('\n');
}

export function compressSlides(content: string): string {
  // TSV: index\ttitle\tshapes\tlayout
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('index');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无幻灯片';
  return `[compressed] ${rows.length}张:\n` + rows.map(r => {
    const c = r.split('\t');
    const parts = [`${c[0]}.`, trunc(c[1] || '(无标题)', 20)];
    if (c[3]) parts.push(c[3]);
    return parts.join(' ');
  }).join('\n');
}

export function compressGet(content: string): string {
  // TSV: index\tname\thasText\tcontent
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('index');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无形状';
  return `[compressed] ${rows.length}个形状:\n` + rows.map(r => {
    const c = r.split('\t');
    return `${c[0]}. ${trunc(c[3] || c[1] || '')}`;
  }).join('\n');
}

export function compressCharts(content: string): string {
  // TSV: shape\ttype\ttitle\tseries
  const lines = content.split('\n').filter(l => l.trim() && /\t/.test(l));
  const isHeader = (lines[0] || '').startsWith('shape');
  const rows = isHeader ? lines.slice(1) : lines;
  if (rows.length === 0) return '[compressed] 无图表';
  return `[compressed] ${rows.length}个图表:\n` + rows.map(r => {
    const c = r.split('\t');
    const head = `${c[0]}. ${c[1] || ''}`;
    const title = c[2] ? ` ${trunc(c[2], 20)}` : '';
    const series = c[3] ? ` ×${c[3]}` : '';
    return head + title + series;
  }).join('\n');
}

export function compressRead(content: string): string {
  try {
    const o = JSON.parse(content) as Record<string, unknown>;
    const parts: string[] = [];
    if (o.text) parts.push(trunc(String(o.text)));
    if (o.bold === 'true') parts.push('bold');
    if (o.italic === 'true') parts.push('italic');
    if (o.fontSize) parts.push(`${o.fontSize}pt`);
    return `[compressed] ${parts.join(' ')}`;
  } catch {
    return content.length > 30 ? `[compressed] ${content.slice(0, 30)}...[truncated]` : `[compressed] ${content}`;
  }
}
