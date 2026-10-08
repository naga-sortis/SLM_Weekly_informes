/*
 * Word (.docx) generation for the SLM-OSS informe, following the structure of
 * "INFORME-SLM-OSS-Sortis-<Mes>-<Año>.docx". Uses the docx library (global `docx`).
 * buildDocument() is DOM-free so it can be tested in Node.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SLMDocx = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const ORANGE = 'FF7900';
  const DARK = '2B3140';
  const GREY = 'F2F3F5';
  const BORDER = 'BFC4CC';
  const PAGE_W = 11906; // A4 portrait, twips
  const MARGIN = 1134; // 2 cm
  const CONTENT_W = PAGE_W - 2 * MARGIN;

  /**
   * @param {object} D       the docx namespace
   * @param {object} C       SLMCore
   * @param {object} report  result of SLMCore.computeReport
   * @param {object} meta    title, subtitle, revision, highlights…
   * @param {object} images  optional { logo:{bytes,width,height}, charts:{key:{bar,pie}} }
   */
  function buildDocument(D, C, report, meta, images = {}) {
    const S = C.xmlSafe;
    const rw = report.reportWeek;
    const monthName = report.periodLabel; // month(s) of the report week
    const weekTitle = `Semana ${rw.week} de ${rw.year}`;
    const weekDates = `del ${C.fmtDate(rw.start)} al ${C.fmtDate(rw.end - 86400000)}`;
    const trendMonth = C.monthLabel(report.trendEnd ?? report.reportMonth) + ((report.trendEndPartial ?? report.partialMonth) ? ` (hasta el ${C.fmtDate(report.asOf)})` : '');
    const weeks = report.weeks;
    const n = weeks.length;

    /* ---------- small builders ---------- */
    const run = (text, o = {}) => new D.TextRun({ text: S(text), ...o });
    const para = (children, o = {}) => new D.Paragraph({ children: Array.isArray(children) ? children : [children], ...o });
    const text = (t, o = {}, po = {}) => para(run(t, o), po);
    const h1 = (t) => new D.Paragraph({ heading: D.HeadingLevel.HEADING_1, children: [run(t)], spacing: { before: 360, after: 160 } });
    const h2 = (t) => new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: [run(t)], spacing: { before: 280, after: 120 } });
    const h3 = (t) => new D.Paragraph({ heading: D.HeadingLevel.HEADING_3, children: [run(t)], spacing: { before: 200, after: 100 } });
    const spacer = (after = 120) => new D.Paragraph({ children: [], spacing: { after } });
    const border = { style: D.BorderStyle.SINGLE, size: 4, color: BORDER };
    const borders = { top: border, bottom: border, left: border, right: border };
    const noBorders = {
      top: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    };

    function cell(value, o = {}) {
      const v = value === null || value === undefined ? '' : String(value);
      return new D.TableCell({
        width: { size: o.width, type: D.WidthType.DXA },
        columnSpan: o.span,
        borders: o.noBorder ? noBorders : borders,
        shading: o.fill ? { fill: o.fill, type: D.ShadingType.CLEAR, color: 'auto' } : undefined,
        verticalAlign: D.VerticalAlign.CENTER,
        margins: { top: 40, bottom: 40, left: 70, right: 70 },
        children: [para(run(v, { bold: !!o.bold, color: o.color, size: o.size || 17 }), {
          alignment: o.left ? D.AlignmentType.LEFT : D.AlignmentType.CENTER,
        })],
      });
    }

    function table(rows, widths) {
      return new D.Table({
        width: { size: widths.reduce((a, b) => a + b, 0), type: D.WidthType.DXA },
        columnWidths: widths,
        layout: D.TableLayoutType.FIXED,
        rows,
      });
    }

    /** Numbers of a chart (rows = series, columns = categories) shown under its image. */
    function chartDataTable(t) {
      if (!t || !t.labels || !t.labels.length) return null;
      const multi = t.labels.length > 1;
      const nCols = t.labels.length + (multi ? 1 : 0);
      const LW = Math.min(2400, Math.max(1700, CONTENT_W - nCols * 520));
      const CW = Math.floor((CONTENT_W - LW) / nCols);
      const widths = [LW, ...Array(nCols).fill(CW)];
      const sz = nCols > 9 ? 13 : 15;
      const head = new D.TableRow({ tableHeader: true, children: [
        cell('Casos', { width: LW, fill: DARK, color: 'FFFFFF', bold: true, left: true, size: sz }),
        ...t.labels.map((l) => cell(l, { width: CW, fill: DARK, color: 'FFFFFF', bold: true, size: sz })),
        ...(multi ? [cell('Total', { width: CW, fill: ORANGE, color: 'FFFFFF', bold: true, size: sz })] : []),
      ] });
      const rows = t.rows.map((r) => {
        const fill = r.kind === 'total' ? 'E4E7EC' : (r.kind === 'subtotal' ? GREY : undefined);
        const bold = r.kind !== 'series';
        return new D.TableRow({ children: [
          cell(r.name, { width: LW, left: true, fill, bold, size: sz }),
          ...r.data.map((v) => cell(v, { width: CW, fill, bold, size: sz, color: v === 0 ? '9AA0A6' : undefined })),
          ...(multi ? [cell(r.total, { width: CW, fill, bold: true, size: sz })] : []),
        ] });
      });
      return table([head, ...rows], widths);
    }

    /* ---------- weekly tables: label | sep | INC weeks | sep | OT weeks ---------- */
    const SEP = 140;
    const LABEL = Math.min(3600, CONTENT_W - 2 * SEP - 2 * n * 430);
    const WEEK_W = Math.floor((CONTENT_W - LABEL - 2 * SEP) / (2 * n));
    const dualWidths = [LABEL, SEP, ...weeks.map(() => WEEK_W), SEP, ...weeks.map(() => WEEK_W)];
    const wkLabel = (w) => `sem. ${w.week}`;

    function dualTable(title, rowsSpec, opts = {}) {
      const groupRow = new D.TableRow({
        tableHeader: true,
        children: [
          cell('', { width: LABEL, noBorder: true }), cell('', { width: SEP, noBorder: true }),
          cell('Nº Incidencias', { span: n, width: WEEK_W * n, fill: ORANGE, color: 'FFFFFF', bold: true }),
          cell('', { width: SEP, noBorder: true }),
          cell('Nº OTs', { span: n, width: WEEK_W * n, fill: ORANGE, color: 'FFFFFF', bold: true }),
        ],
      });
      const head = new D.TableRow({
        tableHeader: true,
        children: [
          cell(title, { width: LABEL, fill: DARK, color: 'FFFFFF', bold: true, left: true }),
          cell('', { width: SEP, noBorder: true }),
          ...weeks.map((w) => cell(wkLabel(w), { width: WEEK_W, fill: DARK, color: 'FFFFFF', bold: true })),
          cell('', { width: SEP, noBorder: true }),
          ...weeks.map((w) => cell(wkLabel(w), { width: WEEK_W, fill: DARK, color: 'FFFFFF', bold: true })),
        ],
      });
      const body = rowsSpec.map((r, idx) => {
        const fill = r.total ? 'E4E7EC' : (idx % 2 ? GREY : undefined);
        return new D.TableRow({
          children: [
            cell(r.label, { width: LABEL, left: true, fill, bold: !!r.total }),
            cell('', { width: SEP, noBorder: true }),
            ...r.inc.map((v) => cell(v, { width: WEEK_W, fill, bold: !!r.total })),
            cell('', { width: SEP, noBorder: true }),
            ...r.ot.map((v) => cell(v, { width: WEEK_W, fill, bold: !!r.total })),
          ],
        });
      });
      return table([groupRow, head, ...body], dualWidths);
    }

    const pick = (cat, fn) => report[cat].map(fn);
    const both = (fn) => ({ inc: pick('inc', fn), ot: pick('ot', fn) });

    const newRows = [
      { label: 'Nuevos durante la semana', ...both((x) => x.nuevos) },
      { label: 'Resueltas de las abiertas durante la semana', ...both((x) => x.resueltas) },
      { label: 'Sin resolver (no escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverNoEsc) },
      { label: 'Sin resolver (escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverEsc) },
      { label: 'Devueltas de las abiertas durante la semana', ...both((x) => x.devueltas) },
      { label: 'Abiertas esa semana por el SLM', ...both((x) => x.abiertasSLM) },
    ];
    const backlogRows = (key) => [
      { label: 'Backlog al comenzar la semana', ...both((x) => x[key].inicio) },
      { label: 'Resueltos', ...both((x) => x[key].resueltos) },
      { label: 'Sin resolver', ...both((x) => x[key].sinResolver) },
    ];
    const total = (x) => x.sinResolverNoEsc + x.sinResolverEsc + x.backlogNoEsc.sinResolver + x.backlogEsc.sinResolver;
    const unresolvedRows = [
      { label: 'Sin resolver (no escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverNoEsc) },
      { label: 'Sin resolver (escaladas) de las abiertas esa semana', ...both((x) => x.sinResolverEsc) },
      { label: 'Sin resolver (no escaladas) de semanas anteriores (backlog)', ...both((x) => x.backlogNoEsc.sinResolver) },
      { label: 'Sin resolver (escaladas) de semanas anteriores (backlog)', ...both((x) => x.backlogEsc.sinResolver) },
      { label: 'TOTAL', total: true, ...both(total) },
    ];

    /* ---------- gestor table ---------- */
    function gestorTable() {
      const GL = Math.max(2600, Math.min(4200, CONTENT_W - n * 900 - 800));
      const GW = Math.floor((CONTENT_W - GL - 800) / n);
      const widths = [GL, ...weeks.map(() => GW), 800];
      const head = new D.TableRow({
        tableHeader: true,
        children: [
          cell('CASOS POR GESTOR', { width: GL, fill: DARK, color: 'FFFFFF', bold: true, left: true }),
          ...weeks.map((w) => cell(wkLabel(w), { width: GW, fill: DARK, color: 'FFFFFF', bold: true })),
          cell('Total', { width: 800, fill: ORANGE, color: 'FFFFFF', bold: true }),
        ],
      });
      const rows = report.gestores.map((g, idx) => new D.TableRow({
        children: [
          cell(g.gestor, { width: GL, left: true, fill: idx % 2 ? GREY : undefined, color: g.unidentified ? '606060' : undefined }),
          ...g.counts.map((v) => cell(v, { width: GW, fill: idx % 2 ? GREY : undefined })),
          cell(g.total, { width: 800, bold: true, fill: idx % 2 ? GREY : undefined }),
        ],
      }));
      const totals = weeks.map((_, i) => report.gestores.reduce((a, g) => a + g.counts[i], 0));
      rows.push(new D.TableRow({
        children: [
          cell('TOTAL', { width: GL, left: true, bold: true, fill: 'E4E7EC' }),
          ...totals.map((v) => cell(v, { width: GW, bold: true, fill: 'E4E7EC' })),
          cell(totals.reduce((a, b) => a + b, 0), { width: 800, bold: true, fill: 'E4E7EC' }),
        ],
      }));
      return table([head, ...rows], widths);
    }

    /* ---------- cover ---------- */
    const children = [];
    if (images.logo) {
      const w = 420;
      const h = Math.round(images.logo.height * (w / images.logo.width));
      children.push(para(new D.ImageRun({ type: 'png', data: images.logo.bytes, transformation: { width: w, height: h } }),
        { alignment: D.AlignmentType.CENTER, spacing: { before: 600, after: 600 } }));
    } else {
      children.push(spacer(1800));
    }
    children.push(
      text(meta.title || 'Informe SLM-OSS', { bold: true, size: 52, color: ORANGE }, { alignment: D.AlignmentType.CENTER, spacing: { before: 800, after: 120 } }),
      text(meta.subtitle || '', { size: 32, color: '404040' }, { alignment: D.AlignmentType.CENTER, spacing: { after: 400 } }),
      text(`${weekTitle} · ${monthName}`, { bold: true, size: 36 }, { alignment: D.AlignmentType.CENTER, spacing: { after: 120 } }),
      text(`Del ${C.fmtDate(rw.start)} al ${C.fmtDate(rw.end - 86400000)}`,
        { size: 22, color: '606060' }, { alignment: D.AlignmentType.CENTER, spacing: { after: 1600 } }),
      text('Revisiones:', { bold: true, size: 20 }, { spacing: { after: 80 } }),
    );
    const RW = [1300, 1500, 2300, 2300, CONTENT_W - 7400];
    const rev = meta.revision || {};
    children.push(table([
      new D.TableRow({ tableHeader: true, children: ['Versión', 'Fecha', 'Realizado por', 'Supervisado', 'Comentarios'].map((h, i) => cell(h, { width: RW[i], fill: DARK, color: 'FFFFFF', bold: true })) }),
      new D.TableRow({ children: [rev.version, rev.date, rev.author, rev.supervisor, rev.comments].map((v, i) => cell(v || '', { width: RW[i] })) }),
    ], RW));

    /* ---------- table of contents (built at the end from the headings actually included) ---------- */
    children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    children.push(text('Tabla de contenidos', { bold: true, size: 28, color: ORANGE }, { spacing: { after: 200 } }));
    const tocAt = children.length;
    const tocEntries = [];
    const num = [0, 0, 0];
    const secNum = {};
    // Auto-numbered heading (levels 1–3); levels 1–2 are listed in the table of contents.
    const H = (level, title, key) => {
      num[level - 1]++;
      for (let i = level; i < 3; i++) num[i] = 0;
      const nb = num.slice(0, level).join('.');
      if (key) secNum[key] = nb;
      if (level <= 2) tocEntries.push([nb, title, level - 1]);
      return (level === 1 ? h1 : level === 2 ? h2 : h3)(`${nb} ${title}`);
    };
    // Which parts of the report to include (everything by default).
    const inc = { intro: true, s21: true, s22: true, s23: true, s24: true, s25: true, s26: true, monthly: true, ...(meta.include || {}) };
    let needBreak = true; // the first section after the table of contents starts on a new page
    const startTop = (forceBreak) => {
      if (forceBreak || needBreak) children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
      needBreak = false;
    };

    /* ---------- Introducción ---------- */
    if (inc.intro) {
      startTop(true);
      children.push(H(1, 'INTRODUCCIÓN'));
      children.push(text(
        `El presente documento recoge la actividad del servicio SLM-OSS (${meta.subtitle || 'Gestores Propietarios'}) ` +
        `correspondiente a la ${weekTitle.toLowerCase()} (${weekDates}, ${monthName})` +
        (n > 2 ? `, junto con la evolución de las semanas ${weeks[0].week} a ${weeks[n - 2].week} (desde el ${C.fmtDate(weeks[0].start)}).`
          : n === 2 ? `, junto con la evolución de la semana ${weeks[0].week} (desde el ${C.fmtDate(weeks[0].start)}).` : '.') +
        ' Cada semana refleja la situación de los casos al cierre de esa semana.', { size: 21 }, { spacing: { after: 120 } }));
      children.push(text(
        'Las tablas y gráficos presentados en este informe tienen como base el fichero de BRISE que nos envía semanalmente el CC – OSS. ' +
        `Datos disponibles hasta el ${C.fmtDateTime(meta.dataUntil)}.` +
        (report.partialWeek ? ` La semana ${rw.week} se presenta con los datos disponibles a esa fecha.` : ''), { size: 21 }, { spacing: { after: 120 } }));
    }

    /* ---------- Detalle semanal ---------- */
    if (inc.s21 || inc.s22 || inc.s23 || inc.s24 || inc.s25 || inc.s26) {
      startTop(false);
      children.push(H(1, 'DETALLE SEMANAL'));
    }
    if (inc.s21) {
      children.push(H(2, 'Casos tratados', 's21'));
      children.push(dualTable('CASOS NUEVOS SEMANA', newRows), spacer());
      children.push(dualTable('CASOS DEL BACKLOG NO ESCALADOS', backlogRows('backlogNoEsc')), spacer());
      children.push(dualTable('CASOS DEL BACKLOG ESCALADOS', backlogRows('backlogEsc')), spacer());
      if ((images.charts || {}).weekly) {
        const img = images.charts.weekly.bar;
        const scale = Math.min(1, 620 / img.width);
        children.push(para(new D.ImageRun({ type: 'png', data: img.bytes, transformation: { width: Math.round(img.width * scale), height: Math.round(img.height * scale) } }),
          { alignment: D.AlignmentType.CENTER, spacing: { after: 80 } }));
        const dt = chartDataTable(images.charts.weekly.table);
        if (dt) children.push(dt, spacer(160));
      }
      children.push(text('Las tablas y gráficos presentados en este informe tienen como base el fichero de BRISE que nos envía semanalmente el CC – OSS',
        { italics: true, size: 17, color: '606060' }));
    }

    if (inc.s22) {
      children.push(H(2, 'Nº Casos por gestor'));
      if (report.gestores.length) {
        children.push(gestorTable());
        const ref = secNum.s21 ? `de la tabla ${secNum.s21}` : 'de la tabla de casos tratados';
        const notes = [`Incidencias nuevas de cada semana agrupadas por el gestor indicado en la acción actual (GESTOR - PROBLEMA - TÉCNICO); el total de cada semana coincide con “Nuevos durante la semana” ${ref}.`];
        if (report.gestorUnidentified) notes.push(`“Sin gestor identificado”: incidencias cuya acción actual no indica gestor (sin acción, devueltas o cerradas). Los casos con un técnico trabajando en la acción actual (p. ej. “JR - Trabajando”) se cuentan como OTs.`);
        if (report.gestorInferred) notes.push(`${report.gestorInferred} incidencia(s) sin gestor en la acción actual se han asignado según el gestor citado en su descripción.`);
        if (meta.manualReviewed) notes.push(`${meta.manualReviewed} caso(s) de estas semanas se han revisado manualmente y se cuentan como incidencia u OT según esa revisión.`);
        children.push(text(notes.join(' '), { italics: true, size: 17, color: '606060' }, { spacing: { before: 80 } }));
      } else {
        children.push(text('No hay incidencias nuevas en las semanas del informe.', { italics: true }));
      }
      children.push(spacer());
    }

    if (inc.s23) {
      children.push(H(2, 'Casos sin resolver'));
      children.push(dualTable('CASOS SIN RESOLVER', unresolvedRows), spacer());
    }

    if (inc.s24) {
      children.push(H(2, 'Detalle de los casos sin resolver'));
      children.push(text(`Casos pendientes al cierre de la ${weekTitle.toLowerCase()} (${C.fmtDateTime(report.asOf)}): ${report.pending.length}.`, { size: 20 }, { spacing: { after: 120 } }));
      if (!report.pending.length) {
        children.push(text('No hay casos sin resolver al cierre del periodo.', { italics: true }));
      }
      for (const t of report.pending) {
        const queue = t.restorationGroupId || (t.category === 'INC' ? meta.queueInc : meta.queueOt) || '';
        const lines = [
          ['TicketID', t.id],
          ['Fecha de creación', C.fmtDateTime(t.created)],
          ['Estado', t.open ? t.status : `Current (estado actual: ${t.status})`],
          ['Tipo', t.type + (t.classification
            ? ` (clasificado manualmente como ${t.classification.category === 'INC' ? 'incidencia' : 'OT'} por ${t.classification.user}${t.classification.at ? ' el ' + C.fmtDate(Date.parse(t.classification.at)) : ''})`
            : (t.categoryByAction ? ' (contado como OT: técnico trabajando en la acción actual)' : ''))],
          ['Prioridad', t.priority],
          ['Persona que lo inicia', t.userName],
          ['Grupo que lo inicia', t.groupId + (t.groupName ? ` (${t.groupName})` : '')],
          ['Acción actual', t.action],
          ['Cola asociada', queue],
          ['Referencia Third Party', t.thirdParty + (t.vendor && t.vendor !== 'Otro' ? ` (${t.vendor})` : '')],
        ];
        if (meta.includeDescription) {
          let d = t.description || '';
          if (d.length > 700) d = d.slice(0, 700).trimEnd() + '…';
          lines.push(['Descripción', d.replace(/\s*\n+\s*/g, ' ')]);
        }
        lines.forEach(([k, v], i) => {
          children.push(para([run(`${k}: `, { bold: true, size: 19 }), run(v || '', { size: 19, bold: i === 0 })],
            { spacing: { after: 20 }, keepNext: i < lines.length - 1, keepLines: true }));
        });
        children.push(spacer(160));
      }
    }

    if (inc.s25) {
      children.push(H(2, 'Incidencias destacadas'));
      bulletList(meta.highlights);
    }
    if (inc.s26) {
      children.push(H(2, 'Otros trabajos a destacar'));
      bulletList(meta.otherWork);
    }

    function bulletList(items) {
      const list = (items || []).map((s) => s.trim()).filter(Boolean);
      if (!list.length) { children.push(text('Sin elementos a destacar.', { italics: true, color: '606060' })); return; }
      for (const it of list) children.push(new D.Paragraph({ children: [run(it, { size: 21 })], bullet: { level: 0 }, spacing: { after: 60 } }));
    }

    /* ---------- Detalle mensual ---------- */
    if (inc.monthly) {
      startTop(true);
      children.push(H(1, 'DETALLE MENSUAL'));
      children.push(text(report.trendYear
        ? `Evolución mensual del año ${report.trendYear} (de enero a ${C.MONTHS_ES[(report.trendEnd % 100) - 1].toLowerCase()}${report.trendEndPartial ? `, hasta el ${C.fmtDate(report.asOf)}` : ''}), con datos hasta el cierre de la ${weekTitle.toLowerCase()}.`
        : `Evolución de los últimos ${report.months.length} meses hasta ${trendMonth}, con datos hasta el cierre de la ${weekTitle.toLowerCase()}.`, { size: 20 }, { spacing: { after: 120 } }));
      children.push(text(C.weekSplitText(report), { size: 20 }, { spacing: { after: 120 } }));
      const ch = images.charts || {};
      const intro = (what, by) => text(
        `Gráfico de acuerdo a los casos ${what} al grupo SLM en cola Oceane ${meta.queueInc || 'XSP00025'} y OTs en cola ${meta.queueOt || 'XSP00027'}, ` +
        `teniendo en cuenta ${by} y la evolución a lo largo de los meses.`, { size: 20 }, { spacing: { after: 120 } });
      const addCharts = (key) => {
        const c = ch[key];
        if (!c) return;
        const maxW = 620;
        for (const img of [c.bar, c.pie]) {
          if (!img) continue;
          const scale = Math.min(1, maxW / img.width);
          children.push(para(new D.ImageRun({ type: 'png', data: img.bytes, transformation: { width: Math.round(img.width * scale), height: Math.round(img.height * scale) } }),
            { alignment: D.AlignmentType.CENTER, spacing: { after: img === c.bar && c.table ? 80 : 160 } }));
          if (img === c.bar) { const dt = chartDataTable(c.table); if (dt) children.push(dt, spacer(160)); }
        }
      };
      children.push(H(2, 'Casos abiertos'));
      children.push(intro('abiertos', 'la prioridad'));
      addCharts('opened');
      children.push(H(2, 'Casos resueltos'));
      children.push(intro('resueltos', 'la prioridad'));
      addCharts('resolved');
      children.push(H(2, 'Casos escalados'));
      children.push(H(3, 'Casos escalados por prioridad'));
      children.push(intro('escalados', 'la prioridad'));
      addCharts('escalated');
      children.push(H(3, 'Casos escalados por gestor'));
      children.push(intro('escalados', 'el gestor'));
      addCharts('escByGestor');
      children.push(H(3, 'Casos escalados por fabricante'));
      children.push(intro('escalados', 'el fabricante (STA- o CSR Ericsson, H- Huawei, 1- Nokia)'));
      addCharts('escByVendor');
      children.push(H(2, 'Estado de los casos'));
      children.push(intro('abiertos', 'el estado'));
      addCharts('status');
    }

    const imgAt = (x) => {
      if (!x) return;
      const scale = Math.min(1, 620 / x.width);
      children.push(para(new D.ImageRun({ type: 'png', data: x.bytes, transformation: { width: Math.round(x.width * scale), height: Math.round(x.height * scale) } }),
        { alignment: D.AlignmentType.CENTER, spacing: { after: 160 } }));
    };

    /* ---------- Resumen del periodo (Month / Year filter) ---------- */
    if (meta.period && meta.period.data) {
      const P = meta.period.data;
      const I = P.summary.INC, O = P.summary.OT;
      const pc = images.periodCharts || {};
      startTop(true);
      children.push(H(1, `RESUMEN DEL PERIODO: ${P.label.toUpperCase()}`));
      children.push(text(meta.period.note, { size: 20 }, { spacing: { after: 160 } }));
      const SW = [CONTENT_W - 3 * 1500, 1500, 1500, 1500];
      const srows = [
        ['Casos abiertos en el periodo', I.abiertos, O.abiertos, true],
        ['   Resueltos', I.resueltos, O.resueltos],
        ['   Pendientes al cierre', I.pendientes, O.pendientes],
        ['   Devueltos', I.devueltos, O.devueltos],
        ['Escalados (Ericsson / Huawei / Nokia)', I.escalados, O.escalados],
        ['Abiertos por el SLM', I.abiertosSLM, O.abiertosSLM],
        ['Resueltos durante el periodo (incl. abiertos antes)', I.resueltosEnPeriodo, O.resueltosEnPeriodo],
        ['Pendientes totales al cierre (incl. backlog)', I.pendientesTotales, O.pendientesTotales],
      ];
      children.push(table([
        new D.TableRow({ tableHeader: true, children: [P.label.toUpperCase(), 'Nº Incidencias', 'Nº OTs', 'Total'].map((h, i) =>
          cell(h, { width: SW[i], fill: i ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i === 0 })) }),
        ...srows.map(([l, a, b, bold], idx) => new D.TableRow({ children: [
          cell(l, { width: SW[0], left: true, bold, fill: bold ? 'E4E7EC' : (idx % 2 ? GREY : undefined) }),
          cell(a, { width: SW[1], bold, fill: bold ? 'E4E7EC' : (idx % 2 ? GREY : undefined) }),
          cell(b, { width: SW[2], bold, fill: bold ? 'E4E7EC' : (idx % 2 ? GREY : undefined) }),
          cell(a + b, { width: SW[3], bold: true, fill: bold ? 'E4E7EC' : (idx % 2 ? GREY : undefined) }),
        ] })),
      ], SW), spacer(160));
      const img = imgAt;
      children.push(H(2, 'Casos abiertos y resueltos'));
      const tbl = (k) => { const dt = chartDataTable((images.periodTables || {})[k]); if (dt) children.push(dt, spacer(160)); };
      img(pc.opened); tbl('opened'); img(pc.resolved); tbl('resolved');
      children.push(H(2, 'Estado de los casos'));
      img(pc.status); tbl('status');
      children.push(H(2, 'Casos escalados por fabricante'));
      img(pc.vendor); tbl('vendor');
      children.push(H(2, 'Casos por gestor'));
      img(pc.gestor);
      if (meta.period.gestorCaption) children.push(text(meta.period.gestorCaption, { italics: true, size: 17, color: '606060' }, { spacing: { after: 120 } }));
      const GW4 = [CONTENT_W - 3 * 1400, 1400, 1400, 1400];
      const tot = P.gestores.reduce((a, g) => [a[0] + g.inc, a[1] + g.ot, a[2] + g.total], [0, 0, 0]);
      children.push(table([
        new D.TableRow({ tableHeader: true, children: ['GESTOR', 'Nº Incidencias', 'Nº OTs', 'Total'].map((h, i) =>
          cell(h, { width: GW4[i], fill: i ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i === 0 })) }),
        ...P.gestores.map((g, idx) => new D.TableRow({ children: [
          cell(g.gestor, { width: GW4[0], left: true, fill: idx % 2 ? GREY : undefined, color: g.unidentified || g.techOT ? '606060' : undefined }),
          cell(g.inc, { width: GW4[1], fill: idx % 2 ? GREY : undefined }),
          cell(g.ot, { width: GW4[2], fill: idx % 2 ? GREY : undefined }),
          cell(g.total, { width: GW4[3], bold: true, fill: idx % 2 ? GREY : undefined }),
        ] })),
        new D.TableRow({ children: ['TOTAL', ...tot].map((v, i) => cell(v, { width: GW4[i], bold: true, left: i === 0, fill: 'E4E7EC' })) }),
      ], GW4));
      children.push(text('El gestor se obtiene de la acción actual (GESTOR - PROBLEMA - TÉCNICO); los casos con un técnico trabajando en la acción actual (p. ej. “JR - Trabajando”) se cuentan como OTs en la fila “OTs (técnico trabajando)”, y los casos sin gestor en la acción aparecen como “Sin gestor identificado”. El total coincide con los casos abiertos en el periodo.',
        { italics: true, size: 17, color: '606060' }, { spacing: { before: 80 } }));
    }

    /* ---------- Equipo Sortis (restoration groups XSP00025 / XSP00027) ---------- */
    if (meta.team && meta.team.data) {
      const T = meta.team.data;
      const tc = images.teamCharts || {};
      const fmtH = (h) => (h === null || h === undefined ? '—' : (h < 10 ? h.toFixed(1) : String(Math.round(h))));
      startTop(true);
      children.push(H(1, `EQUIPO SORTIS (${T.groups.join(' / ')}): ${T.label.toUpperCase()}`));
      children.push(text(meta.team.note, { size: 20 }, { spacing: { after: 160 } }));
      children.push(H(2, 'Casos restaurados por cola'));
      const QW = [1500, CONTENT_W - 1500 - 4 * 1250, 1250, 1250, 1250, 1250];
      children.push(table([
        new D.TableRow({ tableHeader: true, children: ['Cola', 'Nombre', 'Nº Incidencias', 'Nº OTs', 'Total', 'Ingenieros'].map((h, i) =>
          cell(h, { width: QW[i], fill: i > 1 ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i < 2 })) }),
        ...T.byGroup.map((g, idx) => new D.TableRow({ children: [g.group, g.name, g.inc, g.ot, g.total, g.engineers].map((v, i) =>
          cell(v, { width: QW[i], left: i < 2, bold: i === 4, fill: idx % 2 ? GREY : undefined })) })),
        new D.TableRow({ children: ['TOTAL', '', T.inc, T.ot, T.total, ''].map((v, i) => cell(v, { width: QW[i], left: i < 2, bold: true, fill: 'E4E7EC' })) }),
      ], QW), spacer(160));

      children.push(H(2, 'Casos restaurados por ingeniero'));
      imgAt(tc.engineers);
      const gcols = T.groups.flatMap((g) => [`${g} Inc.`, `${g} OTs`]);
      const nc = gcols.length + 3;
      const EW0 = Math.max(2400, CONTENT_W - nc * 900);
      const EWc = Math.floor((CONTENT_W - EW0) / nc);
      const EW = [EW0, ...Array(nc).fill(EWc)];
      const sz = 15;
      children.push(table([
        new D.TableRow({ tableHeader: true, children: ['Ingeniero (Restoration user name)', ...gcols, 'Total', 'Mediana horas hasta restauración', 'Cerrados por el mismo usuario'].map((h, i) =>
          cell(h, { width: EW[i], fill: i ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i === 0, size: sz })) }),
        ...T.engineers.map((e, idx) => new D.TableRow({ children: [
          cell(e.user, { width: EW[0], left: true, size: sz, fill: idx % 2 ? GREY : undefined, color: e.noUser ? '606060' : undefined }),
          ...T.groups.flatMap((g) => [e.byGroup[g].inc, e.byGroup[g].ot]).map((v) => cell(v, { width: EWc, size: sz, fill: idx % 2 ? GREY : undefined, color: v === 0 ? '9AA0A6' : undefined })),
          cell(e.total, { width: EWc, bold: true, size: sz, fill: idx % 2 ? GREY : undefined }),
          cell(fmtH(e.medianHours), { width: EWc, size: sz, fill: idx % 2 ? GREY : undefined }),
          cell(e.closedByThem, { width: EWc, size: sz, fill: idx % 2 ? GREY : undefined }),
        ] })),
        new D.TableRow({ children: [
          cell('TOTAL', { width: EW[0], left: true, bold: true, size: sz, fill: 'E4E7EC' }),
          ...T.groups.flatMap((g) => [T.engineers.reduce((a, e) => a + e.byGroup[g].inc, 0), T.engineers.reduce((a, e) => a + e.byGroup[g].ot, 0)])
            .map((v) => cell(v, { width: EWc, bold: true, size: sz, fill: 'E4E7EC' })),
          cell(T.total, { width: EWc, bold: true, size: sz, fill: 'E4E7EC' }),
          cell(fmtH(T.medianHours), { width: EWc, bold: true, size: sz, fill: 'E4E7EC' }),
          cell(T.engineers.reduce((a, e) => a + e.closedByThem, 0), { width: EWc, bold: true, size: sz, fill: 'E4E7EC' }),
        ] }),
      ], EW));
      children.push(text('Ingeniero = “Restoration user name” del Excel. Horas desde la fecha de creación hasta la fecha de restauración (mediana).',
        { italics: true, size: 17, color: '606060' }, { spacing: { before: 80, after: 160 } }));

      if (tc.trend) {
        children.push(H(2, 'Evolución por ingeniero'));
        imgAt(tc.trend);
        const dt = chartDataTable((images.teamTables || {}).trend);
        if (dt) children.push(dt, spacer(160));
      }

      children.push(H(2, 'Grupos de restauración en el periodo (todos los casos)'));
      const RG = T.restorationGroups.slice(0, 25);
      const GWr = [1700, CONTENT_W - 1700 - 4 * 1150, 1150, 1150, 1150, 1150];
      children.push(table([
        new D.TableRow({ tableHeader: true, children: ['Grupo', 'Nombre', 'Nº Incidencias', 'Nº OTs', 'Total', 'Usuarios'].map((h, i) =>
          cell(h, { width: GWr[i], fill: i > 1 ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i < 2 })) }),
        ...RG.map((g, idx) => new D.TableRow({ children: [g.group, g.name, g.inc, g.ot, g.total, g.users].map((v, i) =>
          cell(v, { width: GWr[i], left: i < 2, bold: i === 4 || g.sortis, fill: g.sortis ? 'FFF1E5' : (idx % 2 ? GREY : undefined) })) })),
      ], GWr));
      if (T.restorationGroups.length > RG.length) {
        children.push(text(`Se muestran los ${RG.length} grupos con más casos de ${T.restorationGroups.length}.`, { italics: true, size: 17, color: '606060' }, { spacing: { before: 80 } }));
      }
    }

    /* ---------- fill in the table of contents ---------- */
    children.splice(tocAt, 0, ...tocEntries.map(([nb, title, lvl]) =>
      para([run(`${nb}  `, { bold: lvl === 0 }), run(title, { bold: lvl === 0 })], { indent: { left: lvl * 400 }, spacing: { after: 60 } })));

    const footer = new D.Footer({
      children: [para([
        run(`${meta.title || 'Informe SLM-OSS'} · ${weekTitle} · ${monthName}`, { size: 16, color: '808080' }),
        new D.TextRun({ children: ['\tPágina ', D.PageNumber.CURRENT, ' de ', D.PageNumber.TOTAL_PAGES], size: 16, color: '808080' }),
      ], { tabStops: [{ type: D.TabStopType.RIGHT, position: CONTENT_W }] })],
    });

    return new D.Document({
      creator: 'SLM Weekly Informes',
      title: `${meta.title || 'Informe SLM-OSS'} ${weekTitle} ${monthName}`,
      description: 'Informe SLM-OSS generado automáticamente',
      styles: {
        default: { document: { run: { font: 'Calibri', size: 20 } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 30, bold: true, color: ORANGE, font: 'Calibri' } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 25, bold: true, color: '1F2937', font: 'Calibri' } },
          { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, color: '374151', font: 'Calibri' } },
        ],
      },
      sections: [{
        properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
        footers: { default: footer },
        children,
      }],
    });
  }

  return { buildDocument };
});
