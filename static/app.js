const qs = (s) => document.querySelector(s);
const qsa = (s) => [...document.querySelectorAll(s)];

const HIDE_FROM_SUMMARY = new Set([
  '6020000 - Finance Income',
  '6030000 - Finance Expense',
  '6010000 - Non Operating Gain Loss',
]);

const els = {
  loginScreen: qs('#loginScreen'),
  appShell: qs('#appShell'),
  loginUsername: qs('#loginUsername'),
  loginPassword: qs('#loginPassword'),
  loginBtn: qs('#loginBtn'),
  loginStatus: qs('#loginStatus'),
  logoutBtn: qs('#logoutBtn'),
  sapFile: qs('#sapFile'),
  osFile: qs('#osFile'),
  entity: qs('#entitySelect'),
  runBtn: qs('#runBtn'),
  hideZeros: qs('#hideZeros'),
  status: qs('#status'),
  overlay: qs('#overlay'),
  resultsPanel: qs('#resultsPanel'),
  summaryBody: qs('#summaryBody'),
  drillBody: qs('#drillBody'),
  debugBox: qs('#debugBox'),
  metaSapRows: qs('#metaSapRows'),
  metaOsRows: qs('#metaOsRows'),
  metaGlCodes: qs('#metaGlCodes'),
  metaUnmapped: qs('#metaUnmapped'),
  metaEntity: qs('#metaEntity'),
};

const DRILL_LAYOUT = [
  { summary: 'Revenue', items: ['4100000 - Revenue'] },
  {
    summary: 'Operating Expense (Ex-D And A)',
    items: [
      '5010000 - Cost Of Raw Materials And Supplies',
      '5020000 - Staff Costs',
      '5030000 - Licence Fees',
      '5050000 - Company Premise Utilities And Maintenance',
      '5060000 - Subcontracting Services',
      '5070000 - Travel And Transport',
      '5080000 - Other Costs',
    ],
  },
  { summary: 'EBITDA', items: ['5040000 - Depreciation And Amortisation'] },
  {
    summary: 'EBIT',
    items: [
      '6020000 - Finance Income',
      '6030000 - Finance Expense',
      '8010000 - Share Of Results Of AJV',
      '6010000 - Non Operating Gain Loss',
      '6990000 - Exceptional Items',
    ],
    derivedAfter: [{ name: 'Net Finance (Income / Expense)' }],
  },
  {
    summary: 'Profit Or Loss Before Tax',
    items: [
      '6070000 - Income Tax Expense',
      '8610000 - Profit Or Loss From Discontinued Operation (Net Of Tax)',
    ],
  },
  { summary: 'Net Profit Or Loss (PAT)', items: ['PL_MI - Minority Interest'] },
  { summary: 'Profit Or Loss Attributable To Owners Of The Company', items: [] },
];

let latestPayload = null;

function fmt(value) {
  const num = Number(value || 0);
  return num.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function hasValue(row) {
  return Number(row?.sap_bfc || row?.sap_amount || 0) !== 0 ||
    Number(row?.onestream || row?.os_amount || 0) !== 0 ||
    Number(row?.difference || 0) !== 0;
}

function setStatus(message, isError = false) {
  els.status.textContent = message || '';
  els.status.className = `status ${message ? (isError ? 'err' : 'ok') : ''}`.trim();
}

function setLoading(show) {
  els.overlay.classList.toggle('show', !!show);
  els.runBtn.disabled = !!show;
}

function renderMeta(meta) {
  els.metaSapRows.textContent = meta.sap_rows ?? '-';
  els.metaOsRows.textContent = meta.os_rows ?? '-';
  els.metaGlCodes.textContent = meta.gl_codes ?? '-';
  els.metaUnmapped.textContent = meta.unmapped_gl_codes ?? '-';
  els.metaEntity.textContent = meta.entity ?? '-';
}

function summaryRowClass(name) {
  if ([
    'Revenue',
    'Operating Expense (Ex-D And A)',
    'EBITDA',
    'EBIT',
    'Profit Or Loss Before Tax',
    'Net Profit Or Loss (PAT)',
    'Profit Or Loss Attributable To Owners Of The Company',
  ].includes(name)) return 'level-0 row-highlight summary-row';
  return 'level-1 summary-subrow';
}

function renderSummary(rows) {
  const hideZeros = els.hideZeros.checked;
  const html = rows
    .filter((r) => !HIDE_FROM_SUMMARY.has(r.line_item))
    .filter((r) => !hideZeros || hasValue(r))
    .map((r) => `
      <tr class="${summaryRowClass(r.line_item)}">
        <td>${r.line_item}${r.note ? `<div class="summary-note">${r.note}</div>` : ''}</td>
        <td class="num">${fmt(r.sap_bfc)}</td>
        <td class="num">${fmt(r.onestream)}</td>
        <td class="num">${fmt(r.difference)}</td>
      </tr>
    `)
    .join('');
  els.summaryBody.innerHTML = html || '<tr><td colspan="4">No rows to display.</td></tr>';
}

function renderDrilldown(summaryRows, groups) {
  const hideZeros = els.hideZeros.checked;
  const summaryMap = new Map((summaryRows || []).map((r) => [r.line_item, r]));
  const groupMap = new Map((groups || []).map((g) => [g.line_item, g]));
  const usedLineItems = new Set();
  let html = '';
  let seq = 0;

  DRILL_LAYOUT.forEach((section) => {
    const summary = summaryMap.get(section.summary) || { line_item: section.summary, sap_bfc: 0, onestream: 0, difference: 0 };
    const itemGroups = section.items.map((name) => groupMap.get(name)).filter(Boolean);
    const derivedRows = (section.derivedAfter || []).map((d) => summaryMap.get(d.name)).filter(Boolean);
    const sectionVisible = !hideZeros || hasValue(summary) || itemGroups.some((g) => hasValue(g) || (g.children || []).some(hasValue)) || derivedRows.some(hasValue);
    if (!sectionVisible) return;

    html += `
      <tr class="drill-section-row">
        <td class="drill-section-title">${section.summary}</td>
        <td></td>
        <td></td>
        <td class="num section-num">${fmt(summary.sap_bfc)}</td>
        <td class="num section-num">${fmt(summary.onestream)}</td>
        <td class="num section-num">${fmt(summary.difference)}</td>
      </tr>
    `;

    itemGroups.forEach((group) => {
      usedLineItems.add(group.line_item);
      const childRows = (group.children || []).filter((c) => !hideZeros || hasValue(c));
      const groupVisible = !hideZeros || hasValue(group) || childRows.length > 0;
      if (!groupVisible) return;
      const gid = `grp_${seq++}`;
      const expandable = childRows.length > 0;
      html += `
        <tr class="drill-subgroup-row" data-group="${gid}">
          <td>
            ${expandable ? `<button class="drill-subtoggle" type="button" data-toggle="${gid}"><span>${group.line_item}</span></button>` : `<span class="drill-subtitle static">${group.line_item}</span>`}
          </td>
          <td></td>
          <td></td>
          <td class="num">${fmt(group.sap_bfc)}</td>
          <td class="num">${fmt(group.onestream)}</td>
          <td class="num">${fmt(group.difference)}</td>
        </tr>
      `;
      childRows.forEach((child) => {
        html += `
          <tr class="drill-detail-row" data-child-of="${gid}" style="display:none">
            <td><span class="gl-code">${child.display_code || child.detail_key}</span></td>
            <td>${child.description || ''}</td>
            <td>${child.currency || ''}</td>
            <td class="num">${fmt(child.sap_bfc)}</td>
            <td class="num">${fmt(child.onestream)}</td>
            <td class="num">${fmt(child.difference)}</td>
          </tr>
        `;
      });
    });

    derivedRows.forEach((row) => {
      if (!hideZeros || hasValue(row)) {
        html += `
          <tr class="drill-derived-row">
            <td class="drill-subtitle static">${row.line_item}${row.note ? `<div class="summary-note">${row.note}</div>` : ''}</td>
            <td></td><td></td>
            <td class="num">${fmt(row.sap_bfc)}</td>
            <td class="num">${fmt(row.onestream)}</td>
            <td class="num">${fmt(row.difference)}</td>
          </tr>
        `;
      }
    });
  });

  const leftover = (groups || []).filter((g) => !usedLineItems.has(g.line_item));
  const leftoverVisible = leftover.filter((g) => !hideZeros || hasValue(g) || (g.children || []).some(hasValue));
  if (leftoverVisible.length) {
    html += `
      <tr class="drill-section-row">
        <td class="drill-section-title">Other Visible Lines</td>
        <td></td><td></td><td class="num"></td><td class="num"></td><td class="num"></td>
      </tr>
    `;
    leftoverVisible.forEach((group) => {
      const gid = `grp_${seq++}`;
      const childRows = (group.children || []).filter((c) => !hideZeros || hasValue(c));
      const expandable = childRows.length > 0;
      html += `
        <tr class="drill-subgroup-row" data-group="${gid}">
          <td>${expandable ? `<button class="drill-subtoggle" type="button" data-toggle="${gid}"><span>${group.line_item}</span></button>` : `<span class="drill-subtitle static">${group.line_item}</span>`}</td>
          <td></td><td></td>
          <td class="num">${fmt(group.sap_bfc)}</td>
          <td class="num">${fmt(group.onestream)}</td>
          <td class="num">${fmt(group.difference)}</td>
        </tr>
      `;
      childRows.forEach((child) => {
        html += `
          <tr class="drill-detail-row" data-child-of="${gid}" style="display:none">
            <td><span class="gl-code">${child.display_code || child.detail_key}</span></td>
            <td>${child.description || ''}</td>
            <td>${child.currency || ''}</td>
            <td class="num">${fmt(child.sap_bfc)}</td>
            <td class="num">${fmt(child.onestream)}</td>
            <td class="num">${fmt(child.difference)}</td>
          </tr>
        `;
      });
    });
  }

  els.drillBody.innerHTML = html || '<tr><td colspan="6">No drilldown rows to display.</td></tr>';
  qsa('[data-toggle]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const gid = btn.dataset.toggle;
      const rows = qsa(`[data-child-of="${gid}"]`);
      const expanded = rows.some((r) => r.style.display !== 'none');
      rows.forEach((r) => { r.style.display = expanded ? 'none' : ''; });
    });
  });
}

function renderDebug(debug) {
  els.debugBox.textContent = JSON.stringify(debug, null, 2);
}

function renderAll(payload) {
  latestPayload = payload;
  els.resultsPanel.style.display = '';
  renderMeta(payload.meta || {});
  renderSummary(payload.summary || []);
  renderDrilldown(payload.summary || [], payload.drilldown || []);
  renderDebug(payload.debug || {});
}

async function runRecon() {
  if (!els.sapFile.files[0] || !els.osFile.files[0]) {
    setStatus('Please upload both raw files first.', true);
    return;
  }
  const fd = new FormData();
  fd.append('sap_file', els.sapFile.files[0]);
  fd.append('os_file', els.osFile.files[0]);
  fd.append('entity', els.entity.value || '');

  setStatus('Running reconciliation...');
  setLoading(true);
  try {
    const resp = await fetch('/api/reconcile', { method: 'POST', body: fd });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || 'Reconciliation failed.');
    renderAll(data);
    const warningCount = (data.debug?.warnings || []).length;
    setStatus(warningCount ? `P&L reconciliation completed with ${warningCount} warning(s).` : 'P&L reconciliation completed.');
  } catch (err) {
    setStatus(err.message || String(err), true);
    els.resultsPanel.style.display = 'none';
  } finally {
    setLoading(false);
  }
}

qsa('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    qsa('.tab-btn').forEach((b) => b.classList.remove('active'));
    qsa('.tab-pane').forEach((p) => p.classList.remove('active'));
    btn.classList.add('active');
    qs(`#${btn.dataset.tab}`).classList.add('active');
  });
});

els.runBtn.addEventListener('click', runRecon);
els.hideZeros.addEventListener('change', () => { if (latestPayload) renderAll(latestPayload); });


function setLoginStatus(message, isError = false) {
  els.loginStatus.textContent = message || '';
  els.loginStatus.className = `status ${message ? (isError ? 'err' : 'ok') : ''}`.trim();
}

function setAuthenticatedView(isAuthenticated) {
  els.loginScreen.style.display = isAuthenticated ? 'none' : 'flex';
  els.appShell.style.display = isAuthenticated ? '' : 'none';
}

async function checkSession() {
  try {
    const resp = await fetch('/api/session');
    const data = await resp.json();
    setAuthenticatedView(!!data.authenticated);
    return !!data.authenticated;
  } catch {
    setAuthenticatedView(false);
    return false;
  }
}

async function login() {
  const username = (els.loginUsername.value || '').trim();
  const password = els.loginPassword.value || '';
  if (!username || !password) {
    setLoginStatus('Please enter both username and password.', true);
    return;
  }

  els.loginBtn.disabled = true;
  setLoginStatus('Signing in...');
  try {
    const resp = await fetch('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.error || 'Login failed.');
    setLoginStatus('Login successful.');
    setAuthenticatedView(true);
    els.loginPassword.value = '';
    setStatus('Logged in successfully.');
  } catch (err) {
    setLoginStatus(err.message || String(err), true);
    setAuthenticatedView(false);
  } finally {
    els.loginBtn.disabled = false;
  }
}

async function logout() {
  try {
    await fetch('/api/logout', { method: 'POST' });
  } catch {}
  latestPayload = null;
  els.resultsPanel.style.display = 'none';
  setStatus('');
  setAuthenticatedView(false);
  els.loginUsername.focus();
}

els.loginBtn?.addEventListener('click', login);
els.logoutBtn?.addEventListener('click', logout);
els.loginPassword?.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') login();
});
els.loginUsername?.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') login();
});

checkSession();
