/* =====================================================================
   Canteen Expenses Dashboard — app.js
   Centralised JS: navigation, data loading, table rendering, charts
   ===================================================================== */

'use strict';

// ------------------------------------------------------------------
// 1.  Currency / Number helpers
// ------------------------------------------------------------------
/**
 * Format a numeric value as "Rs 1,000" (no decimals for whole numbers).
 * This is the single canonical formatter used everywhere in the UI.
 * @param {number|string|null|undefined} value
 * @returns {string}
 */
function formatCurrencyRs(value) {
  if (value === null || value === undefined || value === '') return 'Rs 0';
  const num = Number(value);
  if (isNaN(num)) return 'Rs 0';
  return 'Rs ' + Math.round(num).toLocaleString('en-US');
}

/** Format a plain integer count with comma grouping. */
function formatCount(value) {
  const num = Number(value);
  if (isNaN(num)) return '0';
  return Math.round(num).toLocaleString('en-US');
}

/** Format a decimal (qty, price) — up to 2 decimal places. */
function formatDecimal(value) {
  const num = Number(value);
  if (isNaN(num)) return '0';
  // Remove trailing zeros
  return num % 1 === 0 ? num.toLocaleString('en-US') : num.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

// ------------------------------------------------------------------
// 2.  DOM helpers
// ------------------------------------------------------------------
const $ = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

function spinner() {
  return `<div class="spinner-wrap"><div class="spinner"></div><span>Loading…</span></div>`;
}

function emptyRow(colspan, msg = 'No data available') {
  return `<tr><td colspan="${colspan}" class="tbl-placeholder">
            <div class="ph-icon">📋</div>${msg}</td></tr>`;
}

// ------------------------------------------------------------------
// 3.  Chart.js colour palette
// ------------------------------------------------------------------
const PALETTE = {
  bangladeshi: { bg: 'rgba(88,101,242,.7)',  border: '#5865f2' },
  indian:      { bg: 'rgba(16,185,129,.7)',  border: '#10b981' },
  malagasy:    { bg: 'rgba(245,158,11,.7)',  border: '#f59e0b' },
  srilankan:   { bg: 'rgba(239,68,68,.7)',   border: '#ef4444' },
};

// ------------------------------------------------------------------
// 4.  Navigation
// ------------------------------------------------------------------
let _currentSection = null;
let _sectionLoaders = {};

function showSection(id) {
  if (_currentSection === id) return;
  _currentSection = id;

  $$('.section').forEach(s => s.classList.remove('active'));
  $$('.nav-item').forEach(l => l.classList.remove('active'));

  const section = document.getElementById(id);
  const navLink  = $(`[data-section="${id}"]`);
  if (section) section.classList.add('active');
  if (navLink) navLink.classList.add('active');

  // Update topbar title
  const titles = {
    overview:   'Overview',
    employees:  'Employees',
    comparison: 'Detailed Comparison',
    reports:    'Reports',
  };
  const el = $('#topbar-title');
  if (el) el.textContent = titles[id] || id;

  // Load section data
  if (_sectionLoaders[id]) _sectionLoaders[id]();
}

// ------------------------------------------------------------------
// 5.  Overview section
// ------------------------------------------------------------------
async function loadOverview() {
  setHTML('#stat-total-spend', spinner());

  try {
    const res  = await fetch('/api/overview');
    const json = await res.json();

    setHTML('#stat-total-spend',  formatCurrencyRs(json.total_spend));
    setHTML('#stat-employees',    formatCount(json.total_employees));
    setHTML('#stat-periods',      (json.periods || []).length);
    setHTML('#stat-nationalities', '4');

    if (json.data_source === 'Demo') {
      $$('.badge-demo').forEach(b => b.style.display = 'inline-block');
    }

    // Per-nationality stat cards
    const nt = json.nat_totals || {};
    setHTML('#stat-bd', formatCurrencyRs(nt.bangladeshi || 0));
    setHTML('#stat-in', formatCurrencyRs(nt.indian      || 0));
    setHTML('#stat-ml', formatCurrencyRs(nt.malagasy    || 0));
    setHTML('#stat-sl', formatCurrencyRs(nt.srilankan   || 0));

    // Overview donut chart
    renderOverviewChart(nt);

  } catch (e) {
    console.error('loadOverview', e);
  }
}

let _overviewChart = null;

function renderOverviewChart(natTotals) {
  const ctx = document.getElementById('overviewChart');
  if (!ctx) return;
  if (_overviewChart) { _overviewChart.destroy(); _overviewChart = null; }

  const labels = ['Bangladeshi', 'Indian', 'Malagasy', 'Srilankan'];
  const vals   = [
    natTotals.bangladeshi || 0,
    natTotals.indian      || 0,
    natTotals.malagasy    || 0,
    natTotals.srilankan   || 0,
  ];

  _overviewChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: vals,
        backgroundColor: [
          PALETTE.bangladeshi.bg,
          PALETTE.indian.bg,
          PALETTE.malagasy.bg,
          PALETTE.srilankan.bg,
        ],
        borderColor: [
          PALETTE.bangladeshi.border,
          PALETTE.indian.border,
          PALETTE.malagasy.border,
          PALETTE.srilankan.border,
        ],
        borderWidth: 2,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: '#94a3b8', font: { size: 12 }, padding: 14 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ${formatCurrencyRs(ctx.raw)}`,
          },
        },
      },
    },
  });
}

// ------------------------------------------------------------------
// 6.  Employees section
// ------------------------------------------------------------------
let _employeeChart = null;

async function loadEmployees() {
  setHTML('#employees-table-body', spinner());
  setHTML('#employees-chart-area', spinner());

  try {
    const res  = await fetch('/api/employees');
    const json = await res.json();
    renderEmployeesTable(json.data || []);
    renderEmployeeChart(json.data || []);
  } catch (e) {
    setHTML('#employees-table-body',
      `<tr><td colspan="6" class="tbl-placeholder">Failed to load data</td></tr>`);
    console.error('loadEmployees', e);
  }
}

function renderEmployeesTable(rows) {
  if (!rows.length) {
    setHTML('#employees-table-body', emptyRow(6));
    return;
  }

  let html = '';
  let totalBd = 0, totalIn = 0, totalMl = 0, totalSl = 0;

  rows.forEach(r => {
    totalBd += r.bangladeshi;
    totalIn += r.indian;
    totalMl += r.malagasy;
    totalSl += r.srilankan;
    html += `<tr>
      <td class="col-text">${r.period}</td>
      <td class="col-num">${formatCount(r.bangladeshi)}</td>
      <td class="col-num">${formatCount(r.indian)}</td>
      <td class="col-num">${formatCount(r.malagasy)}</td>
      <td class="col-num">${formatCount(r.srilankan)}</td>
      <td class="col-num"><strong>${formatCount(r.total)}</strong></td>
    </tr>`;
  });

  const grandTotal = totalBd + totalIn + totalMl + totalSl;
  html += `<tr class="row-grand">
    <td class="col-text">TOTAL</td>
    <td class="col-num">${formatCount(totalBd)}</td>
    <td class="col-num">${formatCount(totalIn)}</td>
    <td class="col-num">${formatCount(totalMl)}</td>
    <td class="col-num">${formatCount(totalSl)}</td>
    <td class="col-num">${formatCount(grandTotal)}</td>
  </tr>`;

  setHTML('#employees-table-body', html);
}

function renderEmployeeChart(rows) {
  const wrap = document.getElementById('employees-chart-area');
  if (!wrap) return;
  wrap.innerHTML = '<canvas id="employeeChart"></canvas>';
  const ctx = document.getElementById('employeeChart');
  if (!ctx) return;

  if (_employeeChart) { _employeeChart.destroy(); _employeeChart = null; }

  const labels = rows.map(r => r.period);
  const makeDS = (key, label, palette) => ({
    label,
    data: rows.map(r => r[key]),
    backgroundColor: palette.bg,
    borderColor:     palette.border,
    borderWidth: 1.5,
    borderRadius: 3,
  });

  _employeeChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        makeDS('bangladeshi', 'Bangladeshi', PALETTE.bangladeshi),
        makeDS('indian',      'Indian',      PALETTE.indian),
        makeDS('malagasy',    'Malagasy',    PALETTE.malagasy),
        makeDS('srilankan',   'Srilankan',   PALETTE.srilankan),
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: '#2d3148' } },
        y: { ticks: { color: '#94a3b8' }, grid: { color: '#2d3148' }, beginAtZero: true },
      },
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${formatCount(ctx.raw)}`,
          },
        },
      },
    },
  });
}

// ------------------------------------------------------------------
// 7.  Comparison section
// ------------------------------------------------------------------
let _comparisonChart = null;

async function loadComparison() {
  setHTML('#comparison-table-body', spinner());
  setHTML('#comparison-chart-area', spinner());

  try {
    const res  = await fetch('/api/comparison');
    const json = await res.json();
    renderComparisonTable(json.data || []);
    renderComparisonChart(json.data || []);
  } catch (e) {
    setHTML('#comparison-table-body',
      `<tr><td colspan="6" class="tbl-placeholder">Failed to load data</td></tr>`);
    console.error('loadComparison', e);
  }
}

function renderComparisonTable(rows) {
  if (!rows.length) {
    setHTML('#comparison-table-body', emptyRow(6));
    return;
  }

  let html = '';
  let totBd = 0, totIn = 0, totMl = 0, totSl = 0;

  rows.forEach(r => {
    totBd += r.bangladeshi;
    totIn += r.indian;
    totMl += r.malagasy;
    totSl += r.srilankan;
    html += `<tr>
      <td class="col-text">${r.period}</td>
      <td class="col-num">${formatCurrencyRs(r.bangladeshi)}</td>
      <td class="col-num">${formatCurrencyRs(r.indian)}</td>
      <td class="col-num">${formatCurrencyRs(r.malagasy)}</td>
      <td class="col-num">${formatCurrencyRs(r.srilankan)}</td>
      <td class="col-num"><strong>${formatCurrencyRs(r.grand_total)}</strong></td>
    </tr>`;
  });

  const grandTotal = totBd + totIn + totMl + totSl;
  html += `<tr class="row-grand">
    <td class="col-text">TOTAL</td>
    <td class="col-num">${formatCurrencyRs(totBd)}</td>
    <td class="col-num">${formatCurrencyRs(totIn)}</td>
    <td class="col-num">${formatCurrencyRs(totMl)}</td>
    <td class="col-num">${formatCurrencyRs(totSl)}</td>
    <td class="col-num">${formatCurrencyRs(grandTotal)}</td>
  </tr>`;

  setHTML('#comparison-table-body', html);
}

function renderComparisonChart(rows) {
  const wrap = document.getElementById('comparison-chart-area');
  if (!wrap) return;
  wrap.innerHTML = '<canvas id="comparisonChart"></canvas>';
  const ctx = document.getElementById('comparisonChart');
  if (!ctx) return;

  if (_comparisonChart) { _comparisonChart.destroy(); _comparisonChart = null; }

  const labels = rows.map(r => r.period);
  const makeDS = (key, label, palette) => ({
    label,
    data: rows.map(r => r[key]),
    backgroundColor: palette.bg,
    borderColor:     palette.border,
    borderWidth: 1.5,
    fill: false,
    tension: 0.3,
  });

  _comparisonChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [
        makeDS('bangladeshi', 'Bangladeshi', PALETTE.bangladeshi),
        makeDS('indian',      'Indian',      PALETTE.indian),
        makeDS('malagasy',    'Malagasy',    PALETTE.malagasy),
        makeDS('srilankan',   'Srilankan',   PALETTE.srilankan),
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: '#2d3148' } },
        y: {
          ticks: {
            color: '#94a3b8',
            callback: v => formatCurrencyRs(v),
          },
          grid: { color: '#2d3148' },
          beginAtZero: true,
        },
      },
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${formatCurrencyRs(ctx.raw)}`,
          },
        },
      },
    },
  });
}

// ------------------------------------------------------------------
// 8.  Reports section — inline report tables + export buttons
// ------------------------------------------------------------------
let _reportFilters = { period: '', nationality: '' };

async function loadReports() {
  setHTML('#reports-summary-area', spinner());
  setHTML('#reports-content-area', spinner());

  const params = new URLSearchParams();
  if (_reportFilters.period)      params.set('period',      _reportFilters.period);
  if (_reportFilters.nationality) params.set('nationality', _reportFilters.nationality);

  try {
    const res  = await fetch('/api/reports?' + params.toString());
    const json = await res.json();
    renderReportsSummary(json.summary || {});
    renderReportsContent(json);
  } catch (e) {
    setHTML('#reports-content-area', '<p style="color:var(--text-muted);padding:20px">Failed to load report data.</p>');
    console.error('loadReports', e);
  }
}

function renderReportsSummary(summary) {
  const html = `
    <div class="stat-grid" style="margin-bottom:0">
      <div class="stat-card accent">
        <div class="stat-label">Total Spend</div>
        <div class="stat-value">${formatCurrencyRs(summary.grand_total)}</div>
        <div class="stat-sub">${(summary.periods || []).length} period(s)</div>
      </div>
      <div class="stat-card success">
        <div class="stat-label">Employee Records</div>
        <div class="stat-value">${formatCount(summary.total_employees)}</div>
        <div class="stat-sub">cumulative headcount</div>
      </div>
      <div class="stat-card warning">
        <div class="stat-label">Nationalities</div>
        <div class="stat-value">${(summary.nationalities || []).length}</div>
        <div class="stat-sub">${(summary.nationalities || []).join(', ')}</div>
      </div>
    </div>`;
  setHTML('#reports-summary-area', html);
}

function renderReportsContent(data) {
  let html = '';

  // ---- Employee count table -----------------------------------------
  html += buildCollapsibleSection(
    'emp-section',
    '👥 Employee Count by Period & Nationality',
    buildEmployeeReportTable(data.employees || [])
  );

  // ---- Per-nationality expense tables ---------------------------------
  const natLabels = {
    Bangladeshi: ['nat-bd', '🇧🇩 Bangladeshi Expenses'],
    Indian:      ['nat-in', '🇮🇳 Indian Expenses'],
    Malagasy:    ['nat-ml', '🇲🇬 Malagasy Expenses'],
    Srilankan:   ['nat-sl', '🇱🇰 Srilankan Expenses'],
  };

  Object.entries(data.expenses || {}).forEach(([nat, rows]) => {
    const [badgeCls, title] = natLabels[nat] || ['', nat + ' Expenses'];
    html += buildCollapsibleSection(
      `exp-section-${nat.toLowerCase()}`,
      `<span class="nat-badge ${badgeCls}">${nat}</span> &nbsp;${title}`,
      buildExpenseReportTable(rows, nat)
    );
  });

  setHTML('#reports-content-area', html);

  // Wire up collapsible toggles
  $$('.report-section-title').forEach(el => {
    el.addEventListener('click', () => {
      el.classList.toggle('collapsed');
      const bodyId = el.dataset.target;
      const body   = document.getElementById(bodyId);
      if (body) body.style.display = el.classList.contains('collapsed') ? 'none' : '';
    });
  });
}

function buildCollapsibleSection(id, titleHtml, bodyHtml) {
  return `
    <div class="mb-card">
      <div class="report-section-title card" data-target="${id}-body">
        <h5>${titleHtml}</h5>
        <span class="toggle-icon">▼</span>
      </div>
      <div id="${id}-body" class="card" style="border-top:none;border-radius:0 0 var(--radius) var(--radius)">
        ${bodyHtml}
      </div>
    </div>`;
}

function buildEmployeeReportTable(rows) {
  if (!rows.length) return `<div class="tbl-placeholder"><div class="ph-icon">📋</div>No employee records match the filter.</div>`;

  let html = `
    <div class="table-scroll">
      <table class="data-table">
        <colgroup>
          <col style="width:110px">
          <col style="width:130px">
          <col style="width:110px">
          <col style="width:120px">
          <col style="width:120px">
          <col style="width:100px">
        </colgroup>
        <thead><tr>
          <th class="col-text">Period</th>
          <th class="col-num">Bangladeshi</th>
          <th class="col-num">Indian</th>
          <th class="col-num">Malagasy</th>
          <th class="col-num">Srilankan</th>
          <th class="col-num">Total</th>
        </tr></thead>
        <tbody>`;

  let totBd = 0, totIn = 0, totMl = 0, totSl = 0;
  rows.forEach(r => {
    totBd += r.bangladeshi; totIn += r.indian;
    totMl += r.malagasy;   totSl += r.srilankan;
    html += `<tr>
      <td class="col-text">${r.period}</td>
      <td class="col-num">${formatCount(r.bangladeshi)}</td>
      <td class="col-num">${formatCount(r.indian)}</td>
      <td class="col-num">${formatCount(r.malagasy)}</td>
      <td class="col-num">${formatCount(r.srilankan)}</td>
      <td class="col-num"><strong>${formatCount(r.total)}</strong></td>
    </tr>`;
  });

  html += `<tr class="row-grand">
    <td class="col-text">TOTAL</td>
    <td class="col-num">${formatCount(totBd)}</td>
    <td class="col-num">${formatCount(totIn)}</td>
    <td class="col-num">${formatCount(totMl)}</td>
    <td class="col-num">${formatCount(totSl)}</td>
    <td class="col-num">${formatCount(totBd+totIn+totMl+totSl)}</td>
  </tr>`;

  html += `</tbody></table></div>`;
  return html;
}

function buildExpenseReportTable(rows, nat) {
  if (!rows.length) return `<div class="tbl-placeholder"><div class="ph-icon">📋</div>No expense records match the filter.</div>`;

  // Group by period for subtotals
  const byPeriod = {};
  rows.forEach(r => {
    (byPeriod[r.period] = byPeriod[r.period] || []).push(r);
  });

  let html = `
    <div class="table-scroll">
      <table class="data-table">
        <colgroup>
          <col style="width:110px">
          <col style="width:160px">
          <col style="width:80px">
          <col style="width:80px">
          <col style="width:130px">
          <col style="width:140px">
        </colgroup>
        <thead><tr>
          <th class="col-text">Period</th>
          <th class="col-text">Item</th>
          <th class="col-num">Qty</th>
          <th class="col-ctr">Unit</th>
          <th class="col-num">Unit Price</th>
          <th class="col-num">Total</th>
        </tr></thead>
        <tbody>`;

  let grandTotal = 0;
  Object.entries(byPeriod).forEach(([period, periodRows]) => {
    let periodTotal = 0;
    periodRows.forEach(r => {
      periodTotal += r.total;
      html += `<tr>
        <td class="col-text">${r.period}</td>
        <td class="col-text">${r.items}</td>
        <td class="col-num">${formatDecimal(r.qty)}</td>
        <td class="col-ctr">${r.unit}</td>
        <td class="col-num">${formatCurrencyRs(r.unit_price)}</td>
        <td class="col-num">${formatCurrencyRs(r.total)}</td>
      </tr>`;
    });
    grandTotal += periodTotal;
    html += `<tr class="row-subtotal">
      <td class="col-text" colspan="5">Subtotal — ${period}</td>
      <td class="col-num">${formatCurrencyRs(periodTotal)}</td>
    </tr>`;
  });

  html += `<tr class="row-grand">
    <td class="col-text" colspan="5">GRAND TOTAL — ${nat}</td>
    <td class="col-num">${formatCurrencyRs(grandTotal)}</td>
  </tr>`;

  html += `</tbody></table></div>`;
  return html;
}

// ------------------------------------------------------------------
// 9.  Reports filter handlers
// ------------------------------------------------------------------
function applyReportFilter() {
  _reportFilters.period      = ($('#report-period-filter') || {}).value  || '';
  _reportFilters.nationality = ($('#report-nat-filter')    || {}).value  || '';
  loadReports();
}

function clearReportFilter() {
  const pf = $('#report-period-filter');
  const nf = $('#report-nat-filter');
  if (pf) pf.value = '';
  if (nf) nf.value = '';
  _reportFilters = { period: '', nationality: '' };
  loadReports();
}

// ------------------------------------------------------------------
// 10.  Download helpers
// ------------------------------------------------------------------
function downloadCSV() {
  const params = _downloadParams();
  window.location.href = '/download/csv' + (params ? '?' + params : '');
}

function downloadExcel() {
  const params = _downloadParams();
  window.location.href = '/download/excel' + (params ? '?' + params : '');
}

function _downloadParams() {
  const p = new URLSearchParams();
  if (_reportFilters.period)      p.set('period',      _reportFilters.period);
  if (_reportFilters.nationality) p.set('nationality', _reportFilters.nationality);
  return p.toString();
}

// ------------------------------------------------------------------
// 11.  Utility: setHTML without jank
// ------------------------------------------------------------------
function setHTML(selector, html) {
  const el = $(selector);
  if (el) el.innerHTML = html;
}

// ------------------------------------------------------------------
// 12.  Register section loaders & initialise
// ------------------------------------------------------------------
_sectionLoaders = {
  overview:   loadOverview,
  employees:  loadEmployees,
  comparison: loadComparison,
  reports:    loadReports,
};

document.addEventListener('DOMContentLoaded', () => {
  // Nav clicks
  $$('.nav-item[data-section]').forEach(el => {
    el.addEventListener('click', e => {
      e.preventDefault();
      showSection(el.dataset.section);
    });
  });

  // Report filter buttons
  const applyBtn = $('#apply-filter-btn');
  const clearBtn = $('#clear-filter-btn');
  if (applyBtn) applyBtn.addEventListener('click', applyReportFilter);
  if (clearBtn) clearBtn.addEventListener('click', clearReportFilter);

  // Download buttons
  const csvBtn   = $('#btn-csv');
  const xlsxBtn  = $('#btn-xlsx');
  if (csvBtn)  csvBtn.addEventListener('click',  downloadCSV);
  if (xlsxBtn) xlsxBtn.addEventListener('click', downloadExcel);

  // Start on overview
  showSection('overview');
});
