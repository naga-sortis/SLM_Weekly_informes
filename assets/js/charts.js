/* Chart helpers: on-screen charts and PNG rendering for the Word report. */
(function (root) {
  'use strict';

  const SERIES_COLORS = {
    'OTs': '#8a94a8',
    'Incidencias': '#2e6be6',
    'Inc. P1': '#d92d20',
    'Inc. P2': '#f79009',
    'Inc. P3': '#2e6be6',
    'Inc. P4': '#12a594',
    'Inc. otras': '#b0b6c2',
    'Ericsson': '#2e6be6',
    'Huawei': '#d92d20',
    'Nokia': '#12a594',
    'Inc. Current': '#f79009',
    'Inc. Resolved': '#12a594',
    'Inc. Restored': '#0ba5ec',
    'Inc. Closed': '#2e6be6',
    'Inc. Devuelto': '#d92d20',
    'OTs Current': '#fdc463',
    'OTs Cerradas': '#8a94a8',
    'OTs Devueltas': '#f97066',
    'Otros': '#b0b6c2',
  };
  const PALETTE = ['#2e6be6', '#f79009', '#12a594', '#d92d20', '#7a5af8', '#ee46bc',
    '#66a61e', '#a15c07', '#0ba5ec', '#6172f3', '#e04f16', '#5f6b7a'];

  function colorFor(name, i) {
    return SERIES_COLORS[name] || PALETTE[i % PALETTE.length];
  }

  function stackOf(name) {
    if (name.startsWith('OTs')) return 'OT';
    return 'INC';
  }

  /** Build a Chart.js stacked-bar config from report series. */
  function barConfig(labels, series, opts = {}) {
    const dark = !!opts.dark;
    const text = dark ? '#c9ceda' : '#3b4252';
    const grid = dark ? 'rgba(255,255,255,.08)' : 'rgba(0,0,0,.07)';
    return {
      type: 'bar',
      data: {
        labels,
        datasets: series.map((s, i) => ({
          label: s.name,
          data: s.data,
          backgroundColor: colorFor(s.name, i),
          borderColor: dark ? '#1b1e26' : '#ffffff',
          borderWidth: 1,
          borderRadius: 2,
          maxBarThickness: 46,
          stack: opts.splitStacks ? stackOf(s.name) : 'all',
        })),
      },
      options: {
        responsive: !opts.static,
        maintainAspectRatio: false,
        animation: opts.static ? false : { duration: 300 },
        devicePixelRatio: opts.static ? 2 : undefined,
        indexAxis: opts.horizontal ? 'y' : 'x',
        layout: { padding: { top: opts.horizontal ? 0 : 18, right: opts.horizontal ? 36 : 0 } },
        interaction: { mode: 'index', intersect: false, axis: opts.horizontal ? 'y' : 'x' },
        plugins: {
          legend: { position: 'bottom', labels: { color: text, boxWidth: 12, boxHeight: 12, padding: 12, font: { size: opts.static ? 13 : 12 } } },
          title: opts.title ? { display: true, text: opts.title, color: text, font: { size: 15, weight: '600' }, padding: { bottom: 12 } } : { display: false },
          tooltip: {
            callbacks: {
              footer: (items) => {
                const sum = items.reduce((a, it) => a + ((opts.horizontal ? it.parsed.x : it.parsed.y) || 0), 0);
                return items.length > 1 ? `Total: ${sum}` : '';
              },
            },
          },
        },
        scales: opts.horizontal ? {
          y: { stacked: true, ticks: { color: text, autoSkip: false, font: { size: opts.static ? 12 : 11 } }, grid: { display: false } },
          x: { stacked: true, beginAtZero: true, grace: '8%', ticks: { color: text, precision: 0, font: { size: opts.static ? 13 : 12 } }, grid: { color: grid } },
        } : {
          x: { stacked: true, ticks: { color: text, font: { size: opts.static ? 13 : 12 } }, grid: { display: false } },
          y: { stacked: true, beginAtZero: true, grace: '8%', ticks: { color: text, precision: 0, font: { size: opts.static ? 13 : 12 } }, grid: { color: grid } },
        },
      },
      plugins: opts.static ? [whiteBackground, stackTotals(text, opts)] : [stackTotals(text, opts)],
    };
  }

  /** Draws the total of each stacked bar (per stack) above the bar, or after it for horizontal charts. */
  function stackTotals(color, opts) {
    return {
      id: 'stackTotals',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const horizontal = chart.options.indexAxis === 'y';
        const n = chart.data.labels.length;
        ctx.save();
        ctx.fillStyle = color;
        ctx.font = `600 ${opts.static ? 13 : 11}px system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif`;
        ctx.textAlign = horizontal ? 'left' : 'center';
        ctx.textBaseline = horizontal ? 'middle' : 'bottom';
        for (let i = 0; i < n; i++) {
          const stacks = new Map();
          chart.data.datasets.forEach((ds, di) => {
            if (!chart.isDatasetVisible(di)) return;
            const v = Number(ds.data[i]) || 0;
            const key = ds.stack || 'all';
            const cur = stacks.get(key) || { sum: 0, el: null };
            cur.sum += v;
            if (v > 0) cur.el = chart.getDatasetMeta(di).data[i];
            stacks.set(key, cur);
          });
          for (const { sum, el } of stacks.values()) {
            if (!el || !sum) continue;
            if (horizontal) ctx.fillText(String(sum), el.x + 5, el.y);
            else ctx.fillText(String(sum), el.x, el.y - 3);
          }
        }
        ctx.restore();
      },
    };
  }

  function doughnutConfig(series, monthIndex, title) {
    const items = series.map((s, i) => ({ name: s.name, value: s.data[monthIndex] || 0, color: colorFor(s.name, i) }))
      .filter((x) => x.value > 0);
    const config = {
      type: 'doughnut',
      data: {
        labels: items.map((x) => `${x.name} (${x.value})`),
        datasets: [{ data: items.map((x) => x.value), backgroundColor: items.map((x) => x.color), borderColor: '#ffffff', borderWidth: 2 }],
      },
      options: {
        responsive: false,
        animation: false,
        devicePixelRatio: 2,
        cutout: '55%',
        plugins: {
          legend: { position: 'right', labels: { color: '#3b4252', boxWidth: 12, boxHeight: 12, padding: 10, font: { size: 13 } } },
          title: { display: true, text: title, color: '#3b4252', font: { size: 15, weight: '600' }, padding: { bottom: 10 } },
        },
      },
      plugins: [whiteBackground],
    };
    return { config, empty: items.length === 0 };
  }

  const whiteBackground = {
    id: 'whiteBackground',
    beforeDraw(chart) {
      const { ctx, width, height } = chart;
      ctx.save();
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, width, height);
      ctx.restore();
    },
  };

  /** Render a config off-screen and return PNG bytes + size (for docx ImageRun). */
  function renderPng(config, width, height) {
    const host = document.createElement('div');
    host.style.cssText = `position:fixed;left:-10000px;top:0;width:${width}px;height:${height}px;`;
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    canvas.style.width = width + 'px';
    canvas.style.height = height + 'px';
    host.appendChild(canvas);
    document.body.appendChild(host);
    let chart;
    try {
      chart = new root.Chart(canvas.getContext('2d'), config);
      chart.draw();
      const url = canvas.toDataURL('image/png');
      const b64 = url.slice(url.indexOf(',') + 1);
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return { bytes, width, height };
    } finally {
      if (chart) chart.destroy();
      host.remove();
    }
  }

  /**
   * Numbers behind a stacked chart: one row per series, subtotal rows for incidencias / OTs
   * when the chart mixes both, and a Total row. Columns = chart categories.
   */
  function chartTable(labels, series) {
    const rows = series.map((s) => ({ name: s.name, data: s.data.slice(), kind: 'series' }));
    const sumOf = (list) => labels.map((_, i) => list.reduce((a, s) => a + (Number(s.data[i]) || 0), 0));
    const inc = series.filter((s) => s.name.startsWith('Inc.') || s.name === 'Incidencias');
    const ot = series.filter((s) => s.name.startsWith('OTs'));
    if (inc.length > 1 || (inc.length && ot.length && series.length > 2)) {
      rows.push({ name: 'Total incidencias', data: sumOf(inc), kind: 'subtotal' });
      if (ot.length) rows.push({ name: 'Total OTs', data: sumOf(ot), kind: 'subtotal' });
    }
    rows.push({ name: 'Total', data: sumOf(series), kind: 'total' });
    rows.forEach((r) => { r.total = r.data.reduce((a, b) => a + b, 0); });
    return { labels, rows };
  }

  root.SLMCharts = { chartTable, barConfig, doughnutConfig, renderPng, colorFor };
})(typeof self !== 'undefined' ? self : this);
