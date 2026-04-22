#!/usr/bin/env node
/*
 * Manufacturing dashboard LAN server.
 *
 * Node built-ins only (http, fs, path, url, crypto). Serves the existing
 * bom-viewer.html dashboard, a tech time-logger view, and a small JSON/SSE
 * API. Runtime state (queues, assignments, event log) persists to disk on
 * every mutation so restarts are safe.
 *
 * Start:
 *   node server.js                 # default port 3737
 *   PORT=4000 node server.js
 *
 * The dashboard's "Import JSON" flow is mirrored to POST /api/data so the
 * server and dashboard always agree on the master dataset. BOM/doc/tech
 * shapes match the Apps Script export verbatim — nothing is reshaped here.
 */

const http   = require('http');
const fs     = require('fs');
const path   = require('path');
const url    = require('url');
const crypto = require('crypto');

const ROOT      = __dirname;
const DATA_DIR  = path.join(ROOT, 'data');
const CURRENT_JSON_PATH = path.join(DATA_DIR, 'current.json');   // latest dashboard export
const RUNTIME_PATH      = path.join(DATA_DIR, 'runtime.json');   // queues + assignments + completions
const EVENTS_PATH       = path.join(DATA_DIR, 'events.ndjson');  // append-only event log
const PORT = Number(process.env.PORT) || 3737;

const DOWNTIME_REASONS = [
  'Supplier Fault',
  'Assembly Fault',
  'Priority Change',
  'Equipment Fault'
];

// ---------------------------------------------------------------------------
// Persistence helpers
// ---------------------------------------------------------------------------

function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
}

function readJsonIfExists(p, fallback) {
  try {
    if (!fs.existsSync(p)) return fallback;
    const txt = fs.readFileSync(p, 'utf8');
    return txt ? JSON.parse(txt) : fallback;
  } catch (err) {
    console.error(`[persist] failed to read ${p}:`, err.message);
    return fallback;
  }
}

// Write-then-rename so a crash mid-write can't leave a half-written file.
function writeJsonAtomic(p, obj) {
  const tmp = p + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(obj, null, 2));
  fs.renameSync(tmp, p);
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

ensureDataDir();

let dashboardData = readJsonIfExists(CURRENT_JSON_PATH, null); // { nodes, docNodes, supplierRegistry } | null

const defaultRuntime = () => ({
  version: 1,
  dataImportedAt: null,
  queues: {},        // initials -> [assignmentId]
  assignments: {},   // id -> assignment
  completions: {
    docs: {},        // docId -> { ts, by, techInitials, assignmentId }
    bomNodes: {}     // bomNodeId -> { ts, by }
  }
});

let runtime = readJsonIfExists(RUNTIME_PATH, defaultRuntime());
// Backfill any missing top-level keys from older runtime files.
{
  const d = defaultRuntime();
  for (const k of Object.keys(d)) if (!(k in runtime)) runtime[k] = d[k];
  if (!runtime.completions.docs) runtime.completions.docs = {};
  if (!runtime.completions.bomNodes) runtime.completions.bomNodes = {};
}

let events = loadEvents();

function loadEvents() {
  if (!fs.existsSync(EVENTS_PATH)) return [];
  const lines = fs.readFileSync(EVENTS_PATH, 'utf8').split('\n').filter(Boolean);
  return lines.map(l => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
}

function saveRuntime()  { writeJsonAtomic(RUNTIME_PATH, runtime); }
function saveDataset()  { if (dashboardData) writeJsonAtomic(CURRENT_JSON_PATH, dashboardData); }
function appendEvent(e) {
  events.push(e);
  fs.appendFileSync(EVENTS_PATH, JSON.stringify(e) + '\n');
}

// ---------------------------------------------------------------------------
// Dataset derivations — all match the Apps Script export shape exactly
// ---------------------------------------------------------------------------

// Canonical roster comes from dashboardData.techRegistry when present:
//   [{ id: 'AC', name: 'Alice' }, ...]
// Older datasets may not have that top-level registry, so we fall back to
// deriving a best-effort roster from technician entries embedded on docs/nodes.
function techRegistry() {
  if (!dashboardData) return [];
  if (Array.isArray(dashboardData.techRegistry) && dashboardData.techRegistry.length) {
    return dashboardData.techRegistry
      .map(t => {
        if (t && typeof t === 'object') {
          const id = String(t.id || '').trim();
          const name = String(t.name || t.id || '').trim();
          return id ? { id, name: name || id } : null;
        }
        const id = String(t || '').trim();
        return id ? { id, name: id } : null;
      })
      .filter(Boolean)
      .sort((a, b) => a.id.localeCompare(b.id));
  }
  const set = new Set();
  for (const d of (dashboardData.docNodes || [])) {
    for (const t of (d.technicians || [])) {
      const id = String((t && (t.id || t.name)) || '').trim();
      if (id) set.add(id);
    }
  }
  for (const n of (dashboardData.nodes || [])) {
    for (const t of (n.technicians || [])) {
      const id = String((t && (t.id || t.name)) || '').trim();
      if (id) set.add(id);
    }
  }
  return Array.from(set).sort().map(id => ({ id, name: id }));
}

function nodeById(id) {
  if (!dashboardData) return null;
  return (dashboardData.nodes || []).find(n => n.id === id) || null;
}

function docById(id) {
  if (!dashboardData) return null;
  return (dashboardData.docNodes || []).find(d => d.id === id) || null;
}

// Walk leads_to forward from a doc, stopping at a cycle or a missing link.
// Used to enqueue downstream docs when a mid-chain doc is assigned.
function walkLeadsToChain(startDocId) {
  const seen = new Set();
  const chain = [];
  let cur = startDocId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const d = docById(cur);
    if (!d) break;
    chain.push(d.id);
    cur = d.leads_to || null;
  }
  return chain;
}

// "Head" docs of a BOM = docs in that BOM not referenced by any other
// same-BOM doc's leads_to. Usually one; parallel branches yield several.
function headDocsForBom(bomNodeId) {
  if (!dashboardData) return [];
  const sameBom = (dashboardData.docNodes || []).filter(d => d.bomNodeId === bomNodeId);
  const ledTo = new Set(sameBom.map(d => d.leads_to).filter(Boolean));
  return sameBom.filter(d => !ledTo.has(d.id)).map(d => d.id);
}

// All BOM leaf nodes under a given BOM (post multi-parent expansion, since
// that's the shape the export already has).
function bomDescendants(bomNodeId) {
  if (!dashboardData) return [];
  const nodes = dashboardData.nodes || [];
  const children = new Map();
  for (const n of nodes) {
    if (!children.has(n.parent)) children.set(n.parent, []);
    children.get(n.parent).push(n.id);
  }
  const out = [];
  const stack = [bomNodeId];
  while (stack.length) {
    const cur = stack.pop();
    out.push(cur);
    for (const c of (children.get(cur) || [])) stack.push(c);
  }
  return out;
}

function bomLeafDescendants(bomNodeId) {
  if (!dashboardData) return [];
  const nodes = dashboardData.nodes || [];
  const hasChild = new Set(nodes.map(n => n.parent).filter(Boolean));
  return bomDescendants(bomNodeId).filter(id => !hasChild.has(id));
}

// ---------------------------------------------------------------------------
// Elapsed-time math — a single source of truth used by both API responses
// and log event duration fields.
// ---------------------------------------------------------------------------
function liveElapsed(a, now = Date.now()) {
  let prod = a.productiveMs || 0;
  let down = a.downtimeMs || 0;
  if (a.state === 'active' && a.currentSegmentStartedAt) {
    prod += now - new Date(a.currentSegmentStartedAt).getTime();
  } else if (a.state === 'downtime' && a.currentDowntime) {
    down += now - new Date(a.currentDowntime.startedAt).getTime();
  }
  return { productiveMs: prod, downtimeMs: down };
}

// ---------------------------------------------------------------------------
// Assignment creation
// ---------------------------------------------------------------------------
// An assignment is one (tech × doc) pair with its own timer. Assigning a BOM,
// or a doc to multiple techs, fans out into multiple assignments.

function newId() { return crypto.randomBytes(6).toString('hex'); }

function makeAssignment({ techInitials, docId, bomNodeId, context }) {
  return {
    id: newId(),
    techInitials,
    docId,
    bomNodeId,
    context: context || null, // { kind: 'build'|'batch'|null, buildId?, sn?, batchId?, qty? }
    state: 'queued',
    startedAt: null,
    completedAt: null,
    productiveMs: 0,
    downtimeMs: 0,
    currentSegmentStartedAt: null,
    currentDowntime: null,
    coAssigneeAutoclosed: false,
    createdAt: new Date().toISOString()
  };
}

// Resolve an assign request into a list of (techInitials, docId, bomNodeId)
// triples. Honors:
//   - kind='doc' → enqueue the doc + its leads_to chain
//   - kind='bom' → enqueue every head doc of the BOM and each chain
// techInitials can be a single string or an array (multi-tech).
function resolveAssignTargets({ kind, nodeId, docId }) {
  if (kind === 'doc') {
    const d = docById(docId);
    if (!d) return [];
    return walkLeadsToChain(docId).map(id => ({
      docId: id,
      bomNodeId: docById(id).bomNodeId
    }));
  }
  if (kind === 'bom') {
    const heads = headDocsForBom(nodeId);
    const seen = new Set();
    const out = [];
    for (const h of heads) {
      for (const id of walkLeadsToChain(h)) {
        if (seen.has(id)) continue;
        seen.add(id);
        out.push({ docId: id, bomNodeId: docById(id).bomNodeId });
      }
    }
    return out;
  }
  return [];
}

function ensureQueue(initials) {
  if (!runtime.queues[initials]) runtime.queues[initials] = [];
  return runtime.queues[initials];
}

// ---------------------------------------------------------------------------
// Mutations — every path that changes runtime goes through one of these so
// persistence and SSE broadcasting are uniform.
// ---------------------------------------------------------------------------

function logEvent(type, fields = {}) {
  const e = {
    id: newId(),
    ts: new Date().toISOString(),
    type,
    ...fields
  };
  appendEvent(e);
  broadcast({ type: 'event', event: e });
  return e;
}

function commit() {
  saveRuntime();
  broadcast({ type: 'runtime', runtime });
}

function assign({ techInitials, kind, nodeId, docId, context }) {
  const techs = Array.isArray(techInitials) ? techInitials : [techInitials];
  const targets = resolveAssignTargets({ kind, nodeId, docId });
  if (!targets.length) throw new Error('no docs resolved for assignment');

  const created = [];
  for (const tech of techs) {
    for (const t of targets) {
      // Skip if this tech already has an open assignment for this doc.
      const existing = Object.values(runtime.assignments).find(a =>
        a.techInitials === tech && a.docId === t.docId && a.state !== 'complete'
      );
      if (existing) continue;
      const a = makeAssignment({
        techInitials: tech,
        docId: t.docId,
        bomNodeId: t.bomNodeId,
        context
      });
      runtime.assignments[a.id] = a;
      ensureQueue(tech).push(a.id);
      created.push(a);
      logEvent('assign', {
        techInitials: tech,
        assignmentId: a.id,
        docId: a.docId,
        bomNodeId: a.bomNodeId,
        context: a.context
      });
    }
  }
  commit();
  return created;
}

function unassign(assignmentId) {
  const a = runtime.assignments[assignmentId];
  if (!a) throw new Error('unknown assignment');
  if (a.state !== 'queued') throw new Error('can only unassign queued tasks');
  const q = runtime.queues[a.techInitials] || [];
  const i = q.indexOf(assignmentId);
  if (i >= 0) q.splice(i, 1);
  delete runtime.assignments[assignmentId];
  logEvent('unassign', {
    techInitials: a.techInitials,
    assignmentId: a.id,
    docId: a.docId,
    bomNodeId: a.bomNodeId
  });
  commit();
}

function reorder(techInitials, assignmentIds) {
  const q = runtime.queues[techInitials] || [];
  const before = new Set(q);
  if (assignmentIds.length !== q.length || !assignmentIds.every(id => before.has(id))) {
    throw new Error('reorder must be a permutation of the existing queue');
  }
  // Can't move a started (active/paused/downtime) task out of position 0.
  const active = assignmentIds.findIndex(id => {
    const a = runtime.assignments[id];
    return a && a.state !== 'queued' && a.state !== 'complete';
  });
  if (active > 0) throw new Error('active task must remain at head of queue');
  runtime.queues[techInitials] = assignmentIds;
  commit();
}

// Timer actions. All state transitions pass through here so productive /
// downtime ms are accumulated consistently.
function timerAction({ assignmentId, techInitials, action, reason }) {
  const a = runtime.assignments[assignmentId];
  if (!a) throw new Error('unknown assignment');
  if (a.techInitials !== techInitials) throw new Error('tech mismatch');
  if (a.state === 'complete') throw new Error('assignment already complete');

  const now = Date.now();
  const nowIso = new Date(now).toISOString();

  // Close whatever segment is open right now and fold its duration in.
  function closeOpenSegment() {
    if (a.state === 'active' && a.currentSegmentStartedAt) {
      const ms = now - new Date(a.currentSegmentStartedAt).getTime();
      a.productiveMs += ms;
      a.currentSegmentStartedAt = null;
      return { kind: 'productive', ms };
    }
    if (a.state === 'downtime' && a.currentDowntime) {
      const ms = now - new Date(a.currentDowntime.startedAt).getTime();
      a.downtimeMs += ms;
      const ended = a.currentDowntime;
      a.currentDowntime = null;
      return { kind: 'downtime', ms, reason: ended.reason };
    }
    return null;
  }

  if (action === 'start' || action === 'resume') {
    // Promote head-of-queue assignment to running. Only one running task
    // per tech at a time — pause whatever they had going first.
    for (const [id, other] of Object.entries(runtime.assignments)) {
      if (id !== assignmentId &&
          other.techInitials === techInitials &&
          (other.state === 'active' || other.state === 'downtime')) {
        throw new Error(`tech ${techInitials} already has a running task (${id}); pause or stop it first`);
      }
    }
    if (a.state === 'active') return a;            // no-op
    if (a.state === 'queued') a.startedAt = nowIso;
    // If resuming from downtime, close the downtime segment and emit its
    // duration+reason so the log shows the downtime as a discrete event.
    const closed = closeOpenSegment();
    if (closed && closed.kind === 'downtime') {
      logEvent('downtime_end', {
        techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
        context: a.context, reason: closed.reason, durationMs: closed.ms
      });
    }
    a.state = 'active';
    a.currentSegmentStartedAt = nowIso;
    logEvent(action === 'start' ? 'start' : 'resume', {
      techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
      context: a.context
    });
  }

  else if (action === 'pause') {
    // Plain pause — no reason. Downtime uses a dedicated 'downtime_start'.
    if (a.state !== 'active') throw new Error(`cannot pause from ${a.state}`);
    const closed = closeOpenSegment();
    a.state = 'paused';
    logEvent('pause', {
      techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
      context: a.context, productiveMs: closed ? closed.ms : 0
    });
  }

  else if (action === 'downtime_start') {
    if (!reason || !DOWNTIME_REASONS.includes(reason)) {
      throw new Error('valid downtime reason required');
    }
    // Allowed from active OR paused; if active, close the productive segment.
    if (a.state === 'active') {
      const closed = closeOpenSegment();
      logEvent('pause', {
        techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
        context: a.context, productiveMs: closed ? closed.ms : 0, auto: true
      });
    } else if (a.state !== 'paused') {
      throw new Error(`cannot enter downtime from ${a.state}`);
    }
    a.state = 'downtime';
    a.currentDowntime = { reason, startedAt: nowIso };
    logEvent('downtime_start', {
      techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
      context: a.context, reason
    });
  }

  else if (action === 'stop') {
    // Complete the doc. Any other open assignments for the same doc get
    // auto-closed as co-assignees — their segments close at the same ts and
    // they're marked so supervisor review is easy.
    closeOpenSegment();
    a.state = 'complete';
    a.completedAt = nowIso;
    logEvent('stop', {
      techInitials, assignmentId, docId: a.docId, bomNodeId: a.bomNodeId,
      context: a.context,
      productiveMs: a.productiveMs, downtimeMs: a.downtimeMs
    });
    completeDoc(a.docId, { by: 'tech', techInitials, assignmentId, ts: nowIso });
    closeCoAssignees(a.docId, { keepId: a.id, ts: now, nowIso });
  }

  else {
    throw new Error(`unknown action: ${action}`);
  }

  // Pop completed tasks off the head of the queue so "up next" stays correct.
  const q = runtime.queues[techInitials] || [];
  while (q.length && runtime.assignments[q[0]] && runtime.assignments[q[0]].state === 'complete') {
    q.shift();
  }

  commit();
  return a;
}

function completeDoc(docId, meta) {
  if (runtime.completions.docs[docId]) return; // idempotent
  runtime.completions.docs[docId] = {
    ts: meta.ts,
    by: meta.by,
    techInitials: meta.techInitials || null,
    assignmentId: meta.assignmentId || null
  };
  logEvent('doc_complete', {
    docId,
    techInitials: meta.techInitials || null,
    assignmentId: meta.assignmentId || null,
    by: meta.by
  });
  // A doc's BOM may now be fully covered → walk up for rollup.
  const d = docById(docId);
  if (d && d.bomNodeId) checkBomComplete(d.bomNodeId, meta.ts);
}

// Close sibling assignments when one tech stops a multi-tech doc.
function closeCoAssignees(docId, { keepId, ts, nowIso }) {
  for (const a of Object.values(runtime.assignments)) {
    if (a.id === keepId || a.docId !== docId) continue;
    if (a.state === 'complete') continue;
    // Close whatever segment they had open.
    if (a.state === 'active' && a.currentSegmentStartedAt) {
      a.productiveMs += ts - new Date(a.currentSegmentStartedAt).getTime();
      a.currentSegmentStartedAt = null;
    } else if (a.state === 'downtime' && a.currentDowntime) {
      a.downtimeMs += ts - new Date(a.currentDowntime.startedAt).getTime();
      a.currentDowntime = null;
    }
    a.state = 'complete';
    a.completedAt = nowIso;
    a.coAssigneeAutoclosed = true;
    const q = runtime.queues[a.techInitials] || [];
    const i = q.indexOf(a.id);
    if (i === 0) q.shift();
    else if (i > 0) q.splice(i, 1);
    logEvent('stop', {
      techInitials: a.techInitials,
      assignmentId: a.id,
      docId: a.docId,
      bomNodeId: a.bomNodeId,
      context: a.context,
      productiveMs: a.productiveMs,
      downtimeMs: a.downtimeMs,
      coAssigneeAutoclosed: true
    });
  }
}

// A BOM is complete once every doc that lives on it is complete.
// Rolls up ancestors: a parent BOM is rollup-complete once every leaf
// descendant is complete.
function checkBomComplete(bomNodeId, ts) {
  if (runtime.completions.bomNodes[bomNodeId]) return;
  const docs = (dashboardData.docNodes || []).filter(d => d.bomNodeId === bomNodeId);
  if (!docs.length) return;
  const allDone = docs.every(d => runtime.completions.docs[d.id]);
  if (!allDone) return;
  runtime.completions.bomNodes[bomNodeId] = { ts, by: 'tech' };
  logEvent('bom_complete', { bomNodeId, by: 'tech' });
  rollupAncestors(bomNodeId, ts);
}

function rollupAncestors(completedBomId, ts) {
  const node = nodeById(completedBomId);
  if (!node || !node.parent) return;
  // Walk every ancestor; mark rollup_complete when all leaf descendants done.
  let cur = node.parent;
  while (cur) {
    if (runtime.completions.bomNodes[cur]) break;
    const leaves = bomLeafDescendants(cur).filter(id => id !== cur);
    if (!leaves.length) break;
    const allDone = leaves.every(id => runtime.completions.bomNodes[id]);
    if (!allDone) break;
    runtime.completions.bomNodes[cur] = { ts, by: 'rollup' };
    logEvent('bom_rollup_complete', { bomNodeId: cur, by: 'rollup' });
    const parentNode = nodeById(cur);
    cur = parentNode ? parentNode.parent : null;
  }
}

function importDataset(next) {
  if (!next || !Array.isArray(next.nodes) || !Array.isArray(next.docNodes)) {
    throw new Error('dataset must include { nodes, docNodes }');
  }
  dashboardData = next;
  runtime.dataImportedAt = new Date().toISOString();
  saveDataset();
  saveRuntime();
  logEvent('data_import', {
    nodeCount: dashboardData.nodes.length,
    docCount: dashboardData.docNodes.length,
    techCount: techRegistry().length
  });
  broadcast({ type: 'data', dataImportedAt: runtime.dataImportedAt });
  broadcast({ type: 'runtime', runtime });
}

// ---------------------------------------------------------------------------
// Live status snapshot — derived, not stored. Returned by /api/state so
// the dashboard overlay can paint without re-deriving.
// ---------------------------------------------------------------------------
function liveStateSnapshot() {
  const now = Date.now();
  const perAssignment = {};
  for (const [id, a] of Object.entries(runtime.assignments)) {
    perAssignment[id] = { ...a, live: liveElapsed(a, now) };
  }
  // Rolled up per BOM node for the overlay. Priority for the state label is
  // active > downtime > paused > queued > idle. "complete" is only applied
  // when the BOM itself is fully done (see completions.bomNodes below) — a
  // partial doc completion inside an unfinished BOM shouldn't read as done.
  const perBom = {};
  const rank = { idle: 0, queued: 1, paused: 2, downtime: 3, active: 4 };
  for (const a of Object.values(perAssignment)) {
    const b = a.bomNodeId;
    if (!b) continue;
    if (!perBom[b]) perBom[b] = {
      state: 'idle', activeAssignments: [], productiveMs: 0, downtimeMs: 0, techs: new Set()
    };
    const slot = perBom[b];
    slot.productiveMs += a.live.productiveMs;
    slot.downtimeMs   += a.live.downtimeMs;
    if (a.state !== 'queued' && a.state !== 'complete') {
      slot.activeAssignments.push(a.id);
      slot.techs.add(a.techInitials);
    }
    const mapped = a.state === 'complete' ? 'idle' : a.state;
    if ((rank[mapped] ?? 0) > (rank[slot.state] ?? 0)) slot.state = mapped;
  }
  // Walk up the BOM tree so ancestors reflect their subtree's activity.
  // Overview/Builds/Batch overlays paint every node, not just leaves.
  for (const bomId of Object.keys(perBom)) {
    const direct = perBom[bomId];
    let cur = nodeById(bomId);
    cur = cur ? cur.parent : null;
    while (cur) {
      if (!perBom[cur]) perBom[cur] = {
        state: 'idle', activeAssignments: [], productiveMs: 0, downtimeMs: 0, techs: new Set()
      };
      const anc = perBom[cur];
      // Roll up techs & active assignments; don't double-count ms (would
      // overweight long chains) — ms stays per-node.
      for (const t of (direct.techs instanceof Set ? direct.techs : direct.techs)) anc.techs.add(t);
      for (const aid of direct.activeAssignments) anc.activeAssignments.push(aid);
      if ((rank[direct.state] ?? 0) > (rank[anc.state] ?? 0)) anc.state = direct.state;
      const n = nodeById(cur);
      cur = n ? n.parent : null;
    }
  }
  for (const slot of Object.values(perBom)) {
    if (slot.techs instanceof Set) slot.techs = Array.from(slot.techs);
  }
  // Completions override — only a true BOM completion flips the overlay.
  for (const [bomId, c] of Object.entries(runtime.completions.bomNodes)) {
    if (!perBom[bomId]) perBom[bomId] = {
      state: 'idle', activeAssignments: [], productiveMs: 0, downtimeMs: 0, techs: []
    };
    perBom[bomId].state = c.by === 'rollup' ? 'rollup_complete' : 'complete';
  }
  return {
    dataImportedAt: runtime.dataImportedAt,
    queues: runtime.queues,
    assignments: perAssignment,
    completions: runtime.completions,
    perBom,
    techRegistry: techRegistry(),
    downtimeReasons: DOWNTIME_REASONS
  };
}

// ---------------------------------------------------------------------------
// SSE hub
// ---------------------------------------------------------------------------
const sseClients = new Set();

function broadcast(msg) {
  const line = `data: ${JSON.stringify(msg)}\n\n`;
  for (const res of sseClients) {
    try { res.write(line); } catch { /* client gone; cleaned up on 'close' */ }
  }
}

// Heartbeat keeps idle proxies and browsers from dropping the SSE connection.
setInterval(() => {
  const ping = `: ping ${Date.now()}\n\n`;
  for (const res of sseClients) { try { res.write(ping); } catch {} }
}, 20000).unref();

// ---------------------------------------------------------------------------
// HTTP plumbing
// ---------------------------------------------------------------------------

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', c => {
      data += c;
      if (data.length > 50 * 1024 * 1024) { // 50 MB cap — dashboards stay well under
        reject(new Error('payload too large'));
        req.destroy();
      }
    });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
    'Access-Control-Allow-Origin': '*'
  });
  res.end(body);
}

function sendText(res, status, text, type = 'text/plain; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type });
  res.end(text);
}

function sendFile(res, filePath, type) {
  fs.readFile(filePath, (err, buf) => {
    if (err) return sendText(res, 404, 'not found');
    res.writeHead(200, {
      'Content-Type': type,
      'Cache-Control': 'no-cache'
    });
    res.end(buf);
  });
}

const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  // CORS preflight — permissive on a LAN-only tool.
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  try {
    // ---- Static ----
    if (req.method === 'GET' && (pathname === '/' || pathname === '/index.html')) {
      res.writeHead(302, { Location: '/bom-viewer.html' });
      return res.end();
    }
    if (req.method === 'GET' && pathname === '/bom-viewer.html') {
      return sendFile(res, path.join(ROOT, 'bom-viewer.html'), 'text/html; charset=utf-8');
    }
    if (req.method === 'GET' && (pathname === '/tech' || pathname === '/tech.html')) {
      const p = path.join(ROOT, 'tech.html');
      if (!fs.existsSync(p)) return sendText(res, 503, 'tech view not yet installed');
      return sendFile(res, p, 'text/html; charset=utf-8');
    }
    if (req.method === 'GET' && pathname === '/integration.js') {
      const p = path.join(ROOT, 'integration.js');
      if (!fs.existsSync(p)) return sendText(res, 200, '/* integration.js not yet installed */', 'application/javascript');
      return sendFile(res, p, 'application/javascript');
    }
    if (req.method === 'GET' && pathname === '/favicon.ico') {
      const p = path.join(ROOT, 'public', 'favicon.ico');
      if (fs.existsSync(p)) return sendFile(res, p, 'image/x-icon');
      return sendText(res, 204, '');
    }

    // ---- Data API ----
    if (req.method === 'GET' && pathname === '/api/data') {
      return sendJson(res, 200, dashboardData || { nodes: [], docNodes: [], supplierRegistry: [] });
    }
    if (req.method === 'POST' && pathname === '/api/data') {
      const body = await readBody(req);
      const json = JSON.parse(body);
      importDataset(json);
      return sendJson(res, 200, { ok: true, dataImportedAt: runtime.dataImportedAt });
    }

    if (req.method === 'GET' && pathname === '/api/techs') {
      return sendJson(res, 200, { techs: techRegistry() });
    }

    if (req.method === 'GET' && pathname === '/api/state') {
      return sendJson(res, 200, liveStateSnapshot());
    }

    // ---- SSE ----
    if (req.method === 'GET' && pathname === '/api/events') {
      res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
        'X-Accel-Buffering': 'no'
      });
      res.write(`data: ${JSON.stringify({ type: 'hello', state: liveStateSnapshot() })}\n\n`);
      sseClients.add(res);
      req.on('close', () => sseClients.delete(res));
      return;
    }

    // ---- Mutations ----
    if (req.method === 'POST' && pathname === '/api/assign') {
      const body = JSON.parse(await readBody(req));
      const created = assign(body);
      return sendJson(res, 200, { ok: true, assignments: created });
    }
    if (req.method === 'POST' && pathname === '/api/unassign') {
      const body = JSON.parse(await readBody(req));
      unassign(body.assignmentId);
      return sendJson(res, 200, { ok: true });
    }
    if (req.method === 'POST' && pathname === '/api/reorder') {
      const body = JSON.parse(await readBody(req));
      reorder(body.techInitials, body.assignmentIds);
      return sendJson(res, 200, { ok: true });
    }
    if (req.method === 'POST' && pathname === '/api/timer') {
      const body = JSON.parse(await readBody(req));
      const a = timerAction(body);
      return sendJson(res, 200, { ok: true, assignment: { ...a, live: liveElapsed(a) } });
    }

    // ---- Log ----
    if (req.method === 'GET' && pathname === '/api/log.json') {
      const since = parsed.query.since ? new Date(parsed.query.since).getTime() : 0;
      const out = since
        ? events.filter(e => new Date(e.ts).getTime() >= since)
        : events;
      return sendJson(res, 200, { events: out });
    }

    return sendText(res, 404, 'not found');
  } catch (err) {
    console.error(`[${req.method} ${pathname}]`, err.message);
    return sendJson(res, 400, { ok: false, error: err.message });
  }
});

server.listen(PORT, () => {
  console.log(`manufacturing-dashboard server listening on http://0.0.0.0:${PORT}`);
  console.log(`  dashboard:  http://localhost:${PORT}/bom-viewer.html`);
  console.log(`  tech view:  http://localhost:${PORT}/tech?tech=AC`);
  console.log(`  data dir:   ${DATA_DIR}`);
  console.log(`  dataset:    ${dashboardData ? `loaded (${dashboardData.nodes.length} nodes, ${dashboardData.docNodes.length} docs)` : 'none — POST to /api/data to seed'}`);
});
