/* integration.js — dashboard ↔ server bridge.
 *
 * Loaded by bom-viewer.html. Exposes window.MDS for any dashboard code that
 * needs to push assignments, read live runtime state, or re-render on
 * server-side change.
 *
 * Uses EventSource for push updates (same SSE stream tech.html reads) and
 * fetch() for mutations. Auto-reconnects on close with exponential backoff.
 *
 * Silent no-op if the dashboard is opened as a local file (no server). The
 * dashboard keeps working on localStorage; live features stay dormant.
 */
(function () {
  'use strict';

  // --- Config ---------------------------------------------------------------
  // When the dashboard is served from the same server (localhost:3737 today)
  // relative URLs work. When opened as file:// the dashboard is offline and
  // we skip everything.
  const SAME_ORIGIN = (typeof location !== 'undefined') &&
                      (location.protocol === 'http:' || location.protocol === 'https:');

  const API = {
    state:    '/api/state',
    events:   '/api/events',
    assign:   '/api/assign',
    unassign: '/api/unassign',
    reorder:  '/api/reorder',
    timer:    '/api/timer',
    data:     '/api/data',
    techs:    '/api/techs'
  };

  const RECONNECT_MIN_MS = 1000;
  const RECONNECT_MAX_MS = 15000;

  // --- State cache ----------------------------------------------------------
  // Mirrors whatever the server last broadcast. Shape comes from
  // server.js:liveStateSnapshot() — { dataImportedAt, queues, assignments,
  // completions }. Consumers should treat this as immutable/read-only.
  let cache = emptyCache();

  function emptyCache() {
    return {
      connected: false,
      dataImportedAt: null,
      queues: {},
      assignments: {},
      completions: { docs: {}, bomNodes: {} }
    };
  }

  // --- Subscriber registry --------------------------------------------------
  const stateSubs  = new Set();
  const eventSubs  = new Set();
  const statusSubs = new Set();

  function notifyState()  { for (const cb of stateSubs)  safeCall(cb, cache); }
  function notifyStatus() { for (const cb of statusSubs) safeCall(cb, connectionStatus); }
  function notifyEvent(e) { for (const cb of eventSubs)  safeCall(cb, e); }

  function safeCall(cb, arg) {
    try { cb(arg); }
    catch (err) { console.error('[MDS] subscriber error:', err); }
  }

  let connectionStatus = 'offline';  // 'offline' | 'connecting' | 'connected'

  function setStatus(s) {
    if (connectionStatus === s) return;
    connectionStatus = s;
    cache.connected = (s === 'connected');
    notifyStatus();
  }

  // --- SSE connection -------------------------------------------------------
  let es = null;
  let reconnectMs = RECONNECT_MIN_MS;
  let reconnectTimer = null;

  function connect() {
    if (!SAME_ORIGIN) return;   // file:// open — stay offline
    if (es) return;

    setStatus('connecting');
    try {
      es = new EventSource(API.events);
    } catch (err) {
      console.warn('[MDS] EventSource init failed:', err.message);
      scheduleReconnect();
      return;
    }

    es.onopen = () => {
      reconnectMs = RECONNECT_MIN_MS;
      setStatus('connected');
    };

    es.onmessage = (msg) => {
      let data;
      try { data = JSON.parse(msg.data); }
      catch { return; }
      handleSseFrame(data);
    };

    es.onerror = () => {
      // Browser automatically retries, but we want visible status and capped backoff.
      if (es) { es.close(); es = null; }
      setStatus('offline');
      scheduleReconnect();
    };
  }

  function scheduleReconnect() {
    if (reconnectTimer) return;
    const delay = reconnectMs;
    reconnectMs = Math.min(reconnectMs * 2, RECONNECT_MAX_MS);
    reconnectTimer = setTimeout(() => {
      reconnectTimer = null;
      connect();
    }, delay);
  }

  function handleSseFrame(frame) {
    if (!frame || !frame.type) return;

    if (frame.type === 'hello') {
      if (frame.state) applyRuntime(frame.state);
    }
    else if (frame.type === 'runtime') {
      if (frame.runtime) applyRuntime(frame.runtime);
    }
    else if (frame.type === 'data') {
      cache.dataImportedAt = frame.dataImportedAt || cache.dataImportedAt;
      notifyState();
    }
    else if (frame.type === 'event') {
      if (frame.event) notifyEvent(frame.event);
    }
  }

  function applyRuntime(rt) {
    // Server sends either the full runtime object or the derived liveState
    // snapshot. Both have queues/assignments/completions at the top level;
    // liveState additionally attaches `.live` onto each assignment.
    if (rt.dataImportedAt !== undefined) cache.dataImportedAt = rt.dataImportedAt;
    if (rt.queues)        cache.queues      = rt.queues;
    if (rt.assignments)   cache.assignments = rt.assignments;
    if (rt.completions)   cache.completions = rt.completions;
    notifyState();
  }

  // --- Initial state pull --------------------------------------------------
  // SSE "hello" already delivers a snapshot; this is a fallback for when the
  // first subscriber attaches before the stream settles.
  async function pullState() {
    if (!SAME_ORIGIN) return null;
    try {
      const r = await fetch(API.state);
      if (!r.ok) return null;
      const state = await r.json();
      applyRuntime(state);
      return state;
    } catch {
      return null;
    }
  }

  // --- Live elapsed helper -------------------------------------------------
  // Server attaches `.live = { productiveMs, downtimeMs }` on each assignment
  // in the /api/state snapshot, but SSE updates (type=runtime) ship the raw
  // runtime without .live. Dashboard callers that want a constantly-ticking
  // display should use this to derive the current values on demand.
  function liveElapsed(a, now) {
    now = now || Date.now();
    let prod = a.productiveMs || 0;
    let down = a.downtimeMs   || 0;
    if (a.state === 'active' && a.currentSegmentStartedAt) {
      prod += now - new Date(a.currentSegmentStartedAt).getTime();
    } else if (a.state === 'downtime' && a.currentDowntime) {
      down += now - new Date(a.currentDowntime.startedAt).getTime();
    }
    return { productiveMs: prod, downtimeMs: down };
  }

  // --- Grouping helpers ----------------------------------------------------
  // Assignments share a "group" when they share (techInitials, bomNodeId,
  // context). The sidebar renders one tile per group; the assigned doc chain
  // collapses inside. This matches how tech.html displays the queue.
  function groupKey(a) {
    const ctx = a.context || {};
    const ctxPart = ctx.kind === 'build'
      ? `build:${ctx.buildId || ''}:${ctx.sn || ''}`
      : ctx.kind === 'batch'
        ? `batch:${ctx.batchId || ''}:${ctx.qty || ''}`
        : 'none';
    return `${a.techInitials}|${a.bomNodeId}|${ctxPart}`;
  }

  // Returns array of groups for a tech, in queue order. Each group:
  //   { key, techInitials, bomNodeId, context, assignments[], headState }
  // `assignments` is in queue order. `headState` is the state of the first
  // not-complete assignment, or 'complete' if all are done.
  function queueGroupsFor(techInitials) {
    const order = cache.queues[techInitials] || [];
    const groups = [];
    const byKey = new Map();
    for (const aid of order) {
      const a = cache.assignments[aid];
      if (!a) continue;
      const k = groupKey(a);
      let g = byKey.get(k);
      if (!g) {
        g = {
          key: k,
          techInitials: a.techInitials,
          bomNodeId: a.bomNodeId,
          context: a.context || null,
          assignments: [],
          headState: 'queued'
        };
        byKey.set(k, g);
        groups.push(g);
      }
      g.assignments.push(a);
    }
    for (const g of groups) {
      const head = g.assignments.find(a => a.state !== 'complete');
      g.headState = head ? head.state : 'complete';
    }
    return groups;
  }

  // All assignments (across all techs and states) matching a BOM node + context
  // filter. Used by the dashboard to paint a BOM node with its current status.
  function assignmentsForNode(bomNodeId, contextMatch) {
    const out = [];
    for (const a of Object.values(cache.assignments)) {
      if (a.bomNodeId !== bomNodeId) continue;
      if (contextMatch && !contextMatch(a.context || {})) continue;
      out.push(a);
    }
    return out;
  }

  function techsAssignedToNode(bomNodeId, contextMatch) {
    const set = new Set();
    for (const a of assignmentsForNode(bomNodeId, contextMatch)) {
      if (a.state !== 'complete') set.add(a.techInitials);
    }
    return Array.from(set);
  }

  // --- Mutation wrappers ----------------------------------------------------
  async function postJson(url, body) {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {})
    });
    const json = await r.json().catch(() => ({}));
    if (!r.ok || json.ok === false) {
      const err = new Error(json.error || `HTTP ${r.status}`);
      err.status = r.status;
      err.detail = json;
      throw err;
    }
    return json;
  }

  const api = {
    assign(body)   { return postJson(API.assign,   body); },
    unassign(body) { return postJson(API.unassign, body); },
    reorder(body)  { return postJson(API.reorder,  body); },
    timer(body)    { return postJson(API.timer,    body); },
    async data(body) {
      return postJson(API.data, body);
    },
    async techs() {
      if (!SAME_ORIGIN) return [];
      const r = await fetch(API.techs);
      if (!r.ok) return [];
      const json = await r.json();
      return json.techs || [];
    },
    async log(sinceIso) {
      if (!SAME_ORIGIN) return [];
      const q = sinceIso ? `?since=${encodeURIComponent(sinceIso)}` : '';
      const r = await fetch(`/api/log.json${q}`);
      if (!r.ok) return [];
      const json = await r.json();
      return json.events || [];
    }
  };

  // --- Public surface -------------------------------------------------------
  const MDS = {
    isAvailable() { return SAME_ORIGIN; },
    state:  cache,                   // live-mutated cache, not a snapshot
    api,
    connectionStatus() { return connectionStatus; },
    onStateChange(cb)  { stateSubs.add(cb);  return () => stateSubs.delete(cb);  },
    onEvent(cb)        { eventSubs.add(cb);  return () => eventSubs.delete(cb);  },
    onStatusChange(cb) { statusSubs.add(cb); return () => statusSubs.delete(cb); },
    liveElapsed,
    groupKey,
    queueGroupsFor,
    assignmentsForNode,
    techsAssignedToNode,
    pullState,
    // Called by the dashboard after DOMContentLoaded.
    start() {
      if (!SAME_ORIGIN) return;
      pullState();  // fire-and-forget; SSE "hello" will usually win the race
      connect();
    }
  };

  // Kick off early if the document is already alive; otherwise wait.
  if (typeof window !== 'undefined') {
    window.MDS = MDS;
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => MDS.start());
    } else {
      // Defer one tick so code registering subscribers mid-script gets their
      // callback fired after they've finished setting up.
      setTimeout(() => MDS.start(), 0);
    }
  }
})();
