/* ── AI assistant panel — talks to the SHARED loopback relay (same one
   MPL_System uses): http://127.0.0.1:8765. This file carries NO credential;
   with the relay stopped the panel shows a friendly offline state and the
   rest of the dashboard is untouched (per skill qwen-assistant-static-html,
   shared-relay variant 2026-10-30). ──────────────────────────────────────── */
const RELAY = 'http://127.0.0.1:8765';
const aiState = { history: [], busy: false, online: null };

let aiUser = 'Anthony';
try { aiUser = localStorage.getItem('gpr-ai-user') || 'Anthony' } catch { /* default */ }

/* Allowlisted aggregate projection - totals only, never raw rows. */
function aiContext() {
  const t = T();
  const f = state.filters;
  const usd2 = (v) => Math.round(v).toLocaleString('en-US');
  const lines = [
    `CURRENT APP: Global Purchase Report 3D dashboard (self-contained HTML, NOT MPL_System). Answer about the purchase report only.`,
    `UI_LANGUAGE: ${state.lang === 'zh' ? 'zh-Hant (reply in Traditional Chinese)' : 'en (reply in English)'}`,
    `DATA PERIOD: ${OPTS.months[0]} to ${OPTS.months.at(-1)}; ${(OPTS.months.length)} months; rows total ${ALL_ROWS.length.toLocaleString('en-US')}`,
    `ACTIVE FILTERS: ${JSON.stringify({ entity: f.entity || null, month: f.month || null, stdComm: f.stdComm || null, vendor: f.vendor || null, status: f.status || null, search: state.search || null })}`,
    `SUMMARY (under active filters): totalUsd=${usd2(OV.summary.totalUsd)}, purchaseOrders=${OV.summary.purchaseOrders}, receiptLines=${OV.summary.records}, receivedQty=${Math.round(OV.summary.totalQuantity)}, vendors=${OV.summary.suppliers}, items=${OV.summary.items}, ignoreSharePct=${OV.summary.ignoreShare.toFixed(1)}, lastVsPrevMonthPct=${OV.summary.monthChange.toFixed(1)}`,
    `BY_MONTH(USD): ${OV.monthly.map((m) => `${m.month}:${usd2(m.value)}`).join(', ')}`,
    `BY_ENTITY(USD): ${OV.entities.map((e) => `${e.name}:${usd2(e.value)}`).join(', ')}`,
    `TOP_COMM(USD): ${OV.categories.slice(0, 12).map((c) => `${c.name}:${usd2(c.value)}`).join(', ')}`,
    `TOP_VENDOR(USD): ${OV.vendors.slice(0, 10).map((v) => `${v.name}:${usd2(v.value)}`).join(', ')}`,
    `ALLOWED FILTER VALUES: entity=[${OPTS.entities.join(', ')}] month=[${OPTS.months.join(', ')}] status=[Consider, Ignore] stdComm=[${OPTS.categories.slice(0, 30).join(', ')}] vendor=[${OV.vendors.slice(0, 10).map((v) => v.name).join(', ')}]`,
  ];
  return lines.join('\n').slice(0, 5800);
}

async function aiProbe() {
  try {
    const r = await fetch(RELAY + '/health', { signal: AbortSignal.timeout(2500) });
    aiState.online = r.ok;
  } catch { aiState.online = false }
  return aiState.online;
}

/* model proposes up to two "FILTER key=value" lines; we render them as
   human-confirmation chips - nothing mutates until the user clicks. */
function aiParseFilters(reply) {
  const out = [];
  const keep = [];
  for (const line of String(reply).split(/\r?\n/)) {
    const m = line.match(/^FILTER\s+(entity|month|stdComm|vendor|status)=(.+)$/);
    const value = m ? m[2].trim() : null;
    if (m && ['entity', 'month', 'stdComm', 'vendor', 'status'].includes(m[1]) &&
        (m[1] === 'status' ? ['Consider', 'Ignore'].includes(value)
          : m[1] === 'entity' ? OPTS.entities.includes(value)
          : m[1] === 'month' ? OPTS.months.includes(value)
          : m[1] === 'stdComm' ? OPTS.categories.includes(value)
          : OPTS.vendors.includes(value))) {
      out.push({ key: m[1], value });
    } else {
      keep.push(line);
    }
  }
  return { text: keep.join('\n').trim(), filters: out.slice(0, 2) };
}

function aiBubble(role, text, filters) {
  const t = T();
  const wrap = document.createElement('div');
  wrap.className = `ai-msg ${role}`;
  const body = document.createElement('div');
  body.className = 'ai-body';
  body.textContent = text;                     /* textContent only - model output is never HTML */
  wrap.appendChild(body);
  for (const f of (filters || [])) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'ai-chip';
    const keyLabel = { entity: t.fEntity, month: t.fMonth, stdComm: t.fComm, vendor: t.fVendor, status: t.fStatus }[f.key] || f.key;
    chip.textContent = `${t.aiApply} ${keyLabel}: ${f.value}`;
    chip.addEventListener('click', () => {
      updateFilter(f.key, f.value);
      chip.disabled = true;
      showToast(t.aiAppliedChip(f.key, f.value));
    });
    wrap.appendChild(chip);
  }
  return wrap;
}

async function aiSend(message) {
  const t = T();
  const logEl = $('#ai-log');
  const status = $('#ai-status');
  logEl.appendChild(aiBubble('user', message));
  const thinking = aiBubble('sys', t.aiThinking, null);
  logEl.appendChild(thinking);
  logEl.scrollTop = logEl.scrollHeight;
  aiState.busy = true;
  $('#ai-input').disabled = true;
  $('#ai-send').disabled = true;
  try {
    const r = await fetch(RELAY + '/chat', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      signal: AbortSignal.timeout(75000),
      body: JSON.stringify({ username: aiUser, message, history: aiState.history, context: aiContext() }),
    });
    const j = await r.json().catch(() => null);
    thinking.remove();
    if (!r.ok || !j || j.ok !== true) {
      logEl.appendChild(aiBubble('sys', '⚠ ' + ((j && j.error) || t.aiErrNet), null));
    } else {
      const parsed = aiParseFilters(j.reply);
      const tail = parsed.text + (j.truncated ? `\n${t.aiTrunc}` : '');
      logEl.appendChild(aiBubble('bot', tail || '(empty)', parsed.filters));
      aiState.history.push({ role: 'user', content: message });
      aiState.history.push({ role: 'assistant', content: parsed.text.slice(0, 2000) });
      if (aiState.history.length > 24) aiState.history.splice(0, aiState.history.length - 24);
      if (j.quota) status.querySelector('.ai-quota').textContent = t.aiQuota(j.quota.used, j.quota.cap);
      if (j.warn) logEl.appendChild(aiBubble('sys', '⚠ ' + j.warn, null));
    }
  } catch (e) {
    thinking.remove();
    logEl.appendChild(aiBubble('sys', '⚠ ' + t.aiErrNet, null));
    aiProbe().then((ok) => { if (!ok) aiSetOffline(); });
  }
  aiState.busy = false;
  $('#ai-input').disabled = false;
  $('#ai-send').disabled = false;
  logEl.scrollTop = logEl.scrollHeight;
  $('#ai-input').focus();
}

function aiSetOffline() {
  const t = T();
  const status = $('#ai-status');
  status.innerHTML = `<i class="dot off"></i><span>${t.aiOffline}</span>`;
  $('#ai-input').disabled = true;
  $('#ai-send').disabled = true;
}
function aiSetOnline() {
  const t = T();
  const status = $('#ai-status');
  status.innerHTML = `<i class="dot on"></i><span>${t.aiOnline}</span><span class="ai-quota"></span><button type="button" class="btn" id="ai-clear">${t.aiClear}</button>`;
  $('#ai-clear').addEventListener('click', () => {
    aiState.history = [];
    $('#ai-log').innerHTML = '';
    $('#ai-log').appendChild(aiBubble('sys', T().aiHello, null));
  });
  $('#ai-input').disabled = false;
  $('#ai-send').disabled = false;
}

function closeAi() {
  $('#ai-root').innerHTML = '';
  if (!selectedId && !createOpen) document.body.classList.remove('overlay-open');
  const b = $('#ai-btn'); if (b) b.focus();
}

function openAi() {
  const t = T();
  $('#ai-root').innerHTML = `
    <div class="overlay drawer-overlay">
      <aside class="drawer ai-drawer" role="dialog" aria-modal="true" aria-labelledby="ai-title">
        <header class="overlay-head">
          <div><span class="eyebrow">AI ASSISTANT · QWEN</span><h2 id="ai-title">${t.aiTitle}</h2></div>
          <button class="icon-btn" id="ai-close" type="button" aria-label="close">${icon('x', 19)}</button>
        </header>
        <div class="ai-status" id="ai-status"></div>
        <div class="ai-body-wrap">
          <div class="ai-log" id="ai-log"></div>
          <p class="ai-disclose">${t.aiDisclose}</p>
        </div>
        <form class="ai-compose" id="ai-form">
          <input id="ai-input" type="text" maxlength="4000" placeholder="${t.aiPh}" ${aiState.online === false ? 'disabled' : ''} />
          <button class="btn primary" id="ai-send" type="submit">${t.aiSend}</button>
        </form>
        <footer class="ai-userbar"><span>${t.aiName}</span><input id="ai-user" maxlength="40" value="${esc(aiUser)}" /></footer>
      </aside>
    </div>`;
  document.body.classList.add('overlay-open');
  $('#ai-close').addEventListener('click', closeAi);
  $('#ai-root .overlay').addEventListener('mousedown', (e) => { if (e.target === e.currentTarget) closeAi() });
  $('#ai-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const v = $('#ai-input').value.trim();
    if (!v || aiState.busy) return;
    $('#ai-input').value = '';
    aiSend(v);
  });
  $('#ai-user').addEventListener('change', () => {
    aiUser = $('#ai-user').value.trim() || 'Anthony';
    try { localStorage.setItem('gpr-ai-user', aiUser) } catch { /* fine */ }
  });
  $('#ai-log').appendChild(aiBubble('sys', t.aiHello, null));
  aiProbe().then((ok) => { ok ? aiSetOnline() : aiSetOffline() });
  $('#ai-input').focus();
}

/* global Escape while AI panel is open */
addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && $('#ai-root').firstElementChild) closeAi();
});

$('#ai-btn').addEventListener('click', openAi);
