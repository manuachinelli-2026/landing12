#!/usr/bin/env node
// MCP (stdio) para consultar los issues de /organizacion.
// Lee la misma tabla `issues` de Supabase que usa la página, con la misma anon key pública.
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';

const SUPABASE_URL = process.env.ORG_SUPABASE_URL || 'https://fmnabcmqdanlitzyoyma.supabase.co';
const SUPABASE_ANON_KEY = process.env.ORG_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZtbmFiY21xZGFubGl0enlveW1hIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTM1NDY1NDAsImV4cCI6MjA2OTEyMjU0MH0.hdW4wn29xaHSM6npBWDXY0hRc6g85_r4lkA_eq8zXnU';

const STATUSES = ['backlog', 'review', 'todo', 'in_progress', 'done'];
const STATUS_LABELS = { backlog: 'Backlog', review: 'Review', todo: 'Todo', in_progress: 'In progress', done: 'Done' };
// Mismo esquema que la página: 0 = sin prioridad, 1 = urgente ... 4 = baja
const PRIORITY_LABELS = { 0: 'Sin prioridad', 1: 'Urgente', 2: 'Alta', 3: 'Media', 4: 'Baja' };
const PRIORITY_BY_NAME = { sin_prioridad: 0, urgente: 1, alta: 2, media: 3, baja: 4 };
const ASSIGNEE_LABELS = { Manu: 'Manu', Flor: 'Flor', Ambos: 'Manu y Flor' };

async function fetchIssues() {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/issues?select=*`, {
    headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` },
  });
  if (!res.ok) throw new Error(`Supabase respondió ${res.status}: ${await res.text()}`);
  return res.json();
}

const prioRank = issue => issue.priority || 5;

function sortLikePage(issues) {
  return [...issues].sort((a, b) =>
    STATUSES.indexOf(a.status) - STATUSES.indexOf(b.status)
    || prioRank(a) - prioRank(b)
    || (a.position || 0) - (b.position || 0));
}

function daysSince(iso) {
  return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
}

function updatedText(issue) {
  const days = daysSince(issue.updated_at || issue.created_at);
  if (days <= 0) return 'actualizado hoy';
  return days === 1 ? 'actualizado hace 1 día' : `actualizado hace ${days} días`;
}

// "Manu" incluye los issues asignados a los dos
function matchesAssignee(issue, assignee) {
  if (assignee === 'sin_asignar') return !issue.assignee;
  if (assignee === 'Ambos') return issue.assignee === 'Ambos';
  return issue.assignee === assignee || issue.assignee === 'Ambos';
}

function issueLine(issue) {
  const tags = [
    STATUS_LABELS[issue.status],
    ASSIGNEE_LABELS[issue.assignee] || 'sin asignar',
    updatedText(issue),
  ];
  if (issue.long_term) tags.push('long term');
  if (daysSince(issue.created_at) < 7) tags.push('nuevo');
  return `- [${PRIORITY_LABELS[issue.priority || 0]}] ${issue.title} — ${tags.join(' · ')} (id: ${issue.id})`;
}

function text(t) {
  return { content: [{ type: 'text', text: t }] };
}

const server = new McpServer({ name: 'organizacion-l12', version: '1.0.0' });

server.registerTool('list_issues', {
  title: 'Listar issues',
  description: 'Lista los issues del tablero de organización de la fiesta 12·12 (Manu y Flor). Por defecto excluye los que están en Done. Ordena igual que la página: por status, después prioridad, después orden manual.',
  inputSchema: {
    status: z.array(z.enum(STATUSES)).optional().describe('Filtrar por uno o más status'),
    assignee: z.enum(['Manu', 'Flor', 'Ambos', 'sin_asignar']).optional().describe('Filtrar por asignado. "Manu" o "Flor" incluyen los asignados a los dos'),
    priority: z.array(z.enum(Object.keys(PRIORITY_BY_NAME))).optional().describe('Filtrar por prioridad'),
    include_done: z.boolean().optional().describe('Incluir los issues en Done (por defecto no)'),
    long_term: z.boolean().optional().describe('true: solo long term; false: sin long term'),
    search: z.string().optional().describe('Texto a buscar en título y descripción'),
    stale_days: z.number().int().optional().describe('Solo issues sin actualizar hace al menos esta cantidad de días'),
  },
}, async ({ status, assignee, priority, include_done, long_term, search, stale_days }) => {
  let issues = await fetchIssues();
  if (status) issues = issues.filter(i => status.includes(i.status));
  else if (!include_done) issues = issues.filter(i => i.status !== 'done');
  if (assignee) issues = issues.filter(i => matchesAssignee(i, assignee));
  if (priority) {
    const wanted = priority.map(p => PRIORITY_BY_NAME[p]);
    issues = issues.filter(i => wanted.includes(i.priority || 0));
  }
  if (long_term !== undefined) issues = issues.filter(i => !!i.long_term === long_term);
  if (search) {
    const q = search.toLowerCase();
    issues = issues.filter(i => `${i.title} ${i.description || ''}`.toLowerCase().includes(q));
  }
  if (stale_days !== undefined) issues = issues.filter(i => daysSince(i.updated_at || i.created_at) >= stale_days);
  if (!issues.length) return text('No hay issues con esos filtros.');
  return text(`${issues.length} issue(s):\n${sortLikePage(issues).map(issueLine).join('\n')}`);
});

server.registerTool('get_issue', {
  title: 'Ver issue',
  description: 'Muestra el detalle completo de un issue (descripción y links). Acepta el id o parte del título.',
  inputSchema: {
    id_or_title: z.string().describe('El id del issue o parte de su título'),
  },
}, async ({ id_or_title }) => {
  const issues = await fetchIssues();
  const q = id_or_title.toLowerCase();
  const byId = issues.find(i => i.id === id_or_title);
  const matches = byId ? [byId] : issues.filter(i => i.title.toLowerCase().includes(q));
  if (!matches.length) return text(`No encontré ningún issue con "${id_or_title}".`);
  if (matches.length > 1) {
    return text(`Hay ${matches.length} issues que coinciden, especificá cuál:\n${sortLikePage(matches).map(issueLine).join('\n')}`);
  }
  const i = matches[0];
  return text([
    `# ${i.title}`,
    `Status: ${STATUS_LABELS[i.status]}`,
    `Prioridad: ${PRIORITY_LABELS[i.priority || 0]}`,
    `Asignado: ${ASSIGNEE_LABELS[i.assignee] || 'sin asignar'}`,
    `Long term: ${i.long_term ? 'sí' : 'no'}`,
    `Creado: ${i.created_at.slice(0, 10)}${i.created_by ? ` por ${i.created_by}` : ''} · ${updatedText(i)}`,
    `id: ${i.id}`,
    '',
    i.description ? `Descripción:\n${i.description}` : 'Sin descripción.',
    i.links ? `\nLinks:\n${i.links}` : '',
  ].join('\n'));
});

server.registerTool('pending_summary', {
  title: 'Resumen de pendientes',
  description: 'Resumen de lo que falta hacer para la fiesta: cuántos issues abiertos hay por status y por persona, los urgentes y de prioridad alta, los que hace mucho no se tocan y los nuevos de la semana. Opcionalmente solo para una persona.',
  inputSchema: {
    assignee: z.enum(['Manu', 'Flor']).optional().describe('Resumen solo de una persona (incluye los asignados a los dos)'),
  },
}, async ({ assignee }) => {
  let open = (await fetchIssues()).filter(i => i.status !== 'done');
  if (assignee) open = open.filter(i => matchesAssignee(i, assignee));
  open = sortLikePage(open);

  const byStatus = STATUSES.filter(s => s !== 'done')
    .map(s => `${STATUS_LABELS[s]}: ${open.filter(i => i.status === s).length}`).join(' · ');
  const byPerson = ['Manu', 'Flor']
    .map(p => `${p}: ${open.filter(i => matchesAssignee(i, p)).length}`)
    .concat(`sin asignar: ${open.filter(i => !i.assignee).length}`).join(' · ');
  const section = (title, list) => `\n## ${title} (${list.length})\n${list.length ? list.map(issueLine).join('\n') : '- ninguno'}`;

  return text([
    `# Pendientes${assignee ? ` de ${assignee}` : ''}: ${open.length} issues abiertos`,
    `Por status: ${byStatus}`,
    assignee ? '' : `Por persona: ${byPerson}`,
    section('Urgentes y prioridad alta', open.filter(i => i.priority === 1 || i.priority === 2)),
    section('Sin tocar hace 14 días o más', open.filter(i => daysSince(i.updated_at || i.created_at) >= 14 && !i.long_term)),
    section('Nuevos esta semana', open.filter(i => daysSince(i.created_at) < 7)),
  ].filter(Boolean).join('\n'));
});

await server.connect(new StdioServerTransport());
