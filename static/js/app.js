/* ================================================================
   app.js – shared utilities for Canteen Expenses dashboard
   ================================================================ */

'use strict';

/* ----------------------------------------------------------------
   Chart colour palette (one colour per nationality, consistent)
   ---------------------------------------------------------------- */
const NAT_COLORS = {
  Bangladeshi: { border: '#3d8ef8', bg: 'rgba(61,142,248,0.25)' },
  Indian:      { border: '#f4c542', bg: 'rgba(244,197,66,0.25)'  },
  Malagasy:    { border: '#42d9a8', bg: 'rgba(66,217,168,0.25)'  },
  Srilankan:   { border: '#f87c6a', bg: 'rgba(248,124,106,0.25)' },
};

const NAT_DISPLAY = {
  Bangladeshi: 'Bangladeshi',
  Indian:      'Indian',
  Malagasy:    'Malagasy',
  Srilankan:   'Sri Lankan',
};

const NATIONALITIES = ['Bangladeshi', 'Indian', 'Malagasy', 'Srilankan'];

/* ----------------------------------------------------------------
   Period filter helpers
   ---------------------------------------------------------------- */
function populatePeriodSelects(selFromId, selToId, callback) {
  fetch('/api/periods')
    .then(r => r.json())
    .then(periods => {
      const $from = document.getElementById(selFromId);
      const $to   = document.getElementById(selToId);
      if (!$from || !$to) return;

      [$from, $to].forEach($s => {
        $s.innerHTML = '<option value="">All Periods</option>';
        periods.forEach(p => {
          const o = document.createElement('option');
          o.value = p; o.textContent = p;
          $s.appendChild(o);
        });
      });

      if (callback) callback(periods);
    });
}

function getFilterParams(fromId, toId) {
  const from = document.getElementById(fromId)?.value || '';
  const to   = document.getElementById(toId)?.value   || '';
  const p = new URLSearchParams();
  if (from) p.set('from', from);
  if (to)   p.set('to',   to);
  return p.toString() ? '?' + p.toString() : '';
}

/* ----------------------------------------------------------------
   DataTable factory – creates a styled DataTable with export buttons
   ---------------------------------------------------------------- */
function makeDataTable(tableId, columns, title, opts) {
  const defaults = {
    paging:   true,
    pageLength: 10,
    lengthMenu: [[10, 25, 50, -1], [10, 25, 50, 'All']],
    searching: true,
    ordering:  true,
    info:      true,
    autoWidth: false,
    scrollX:   false,
    dom: '<"row mb-2"<"col-sm-6"l><"col-sm-6 text-end"B>>' +
         '<"row"<"col-sm-12"tr>>' +
         '<"row mt-2"<"col-sm-5"i><"col-sm-7"p>>',
    buttons: [
      {
        extend: 'csvHtml5',
        text: '<i class="bi bi-filetype-csv"></i> CSV',
        className: 'dt-button',
        title: title,
      },
      {
        extend: 'excelHtml5',
        text: '<i class="bi bi-file-earmark-excel"></i> Excel',
        className: 'dt-button',
        title: title,
      },
      {
        extend: 'pdfHtml5',
        text: '<i class="bi bi-file-earmark-pdf"></i> PDF',
        className: 'dt-button',
        title: title,
        orientation: 'landscape',
        pageSize: 'A4',
        customize: function (doc) {
          doc.defaultStyle.fontSize = 9;
          doc.styles.tableHeader.fillColor = '#1a3560';
          doc.styles.tableHeader.color = '#ffffff';
        },
      },
      {
        extend: 'print',
        text: '<i class="bi bi-printer"></i> Print',
        className: 'dt-button',
        title: title,
        autoPrint: true,
      },
    ],
    columns: columns,
    language: {
      emptyTable: 'No data available',
      zeroRecords: 'No matching records found',
    },
  };

  const cfg = Object.assign({}, defaults, opts || {});

  if ($.fn.DataTable.isDataTable('#' + tableId)) {
    $('#' + tableId).DataTable().destroy();
  }

  return $('#' + tableId).DataTable(cfg);
}

/* ----------------------------------------------------------------
   Chart helpers
   ---------------------------------------------------------------- */
function makeLineChart(canvasId, labels, datasets) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  if (ctx._chart) ctx._chart.destroy();

  const chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: datasets.map(ds => ({
        label:           ds.label,
        data:            ds.data,
        borderColor:     ds.color || '#3d8ef8',
        backgroundColor: ds.bgColor || 'rgba(61,142,248,0.15)',
        tension:         0.3,
        fill:            false,
        pointRadius:     4,
        pointHoverRadius: 6,
        borderWidth: 2,
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#8aabd4', font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: Rs ${Number(ctx.parsed.y).toLocaleString()}`,
          },
        },
      },
      scales: {
        x: {
          ticks: { color: '#8aabd4', font: { size: 10 }, maxRotation: 30 },
          grid:  { color: 'rgba(30,58,110,0.4)' },
        },
        y: {
          ticks: {
            color: '#8aabd4',
            font:  { size: 10 },
            callback: v => 'Rs ' + Number(v).toLocaleString(),
          },
          grid: { color: 'rgba(30,58,110,0.4)' },
        },
      },
    },
  });

  ctx._chart = chart;
  return chart;
}

function makeBarChart(canvasId, labels, datasets, opts) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  if (ctx._chart) ctx._chart.destroy();

  const isCurrency = opts && opts.currency;

  const chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: datasets.map(ds => ({
        label:           ds.label,
        data:            ds.data,
        backgroundColor: ds.color || 'rgba(61,142,248,0.6)',
        borderColor:     ds.border || '#3d8ef8',
        borderWidth: 1,
        borderRadius: 4,
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#8aabd4', font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = Number(ctx.parsed.y);
              if (isCurrency) return ` ${ctx.dataset.label}: Rs ${v.toLocaleString()}`;
              return ` ${ctx.dataset.label}: ${v.toLocaleString()}`;
            },
          },
        },
      },
      scales: {
        x: {
          ticks: { color: '#8aabd4', font: { size: 10 }, maxRotation: 30 },
          grid:  { color: 'rgba(30,58,110,0.4)' },
        },
        y: {
          ticks: {
            color: '#8aabd4',
            font:  { size: 10 },
            callback: v => isCurrency ? ('Rs ' + Number(v).toLocaleString()) : Number(v).toLocaleString(),
          },
          grid: { color: 'rgba(30,58,110,0.4)' },
        },
      },
    },
  });

  ctx._chart = chart;
  return chart;
}

function makeDoughnutChart(canvasId, labels, data, colors) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  if (ctx._chart) ctx._chart.destroy();

  const chart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors,
        borderColor: '#0a1628',
        borderWidth: 3,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#8aabd4', font: { size: 12 } }, position: 'bottom' },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: Rs ${Number(ctx.parsed).toLocaleString()}`,
          },
        },
      },
    },
  });

  ctx._chart = chart;
  return chart;
}
