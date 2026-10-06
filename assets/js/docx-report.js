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
    const trendMonth = C.monthLabel(report.reportMonth) + (report.partialMonth ? ` (hasta el ${C.fmtDate(report.asOf)})` : '');
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

    /* ---------- table of contents (static) ---------- */
    children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    children.push(text('Tabla de contenidos', { bold: true, size: 28, color: ORANGE }, { spacing: { after: 200 } }));
    const toc = [
      ['1', 'INTRODUCCIÓN', 0], ['2', 'DETALLE SEMANAL', 0], ['2.1', 'Casos tratados', 1], ['2.2', 'Nº Casos por gestor', 1],
      ['2.3', 'Casos sin resolver', 1], ['2.4', 'Detalle de los casos sin resolver', 1], ['2.5', 'Incidencias destacadas', 1],
      ['2.6', 'Otros trabajos a destacar', 1], ['3', 'DETALLE MENSUAL', 0], ['3.1', 'Casos abiertos', 1], ['3.2', 'Casos resueltos', 1],
      ['3.3', 'Casos escalados', 1], ['3.4', 'Estado de los casos', 1],
    ];
    if (meta.period && meta.period.data) {
      toc.push([`4`, `RESUMEN DEL PERIODO: ${meta.period.data.label.toUpperCase()}`, 0], ['4.1', 'Casos abiertos y resueltos', 1],
        ['4.2', 'Estado de los casos', 1], ['4.3', 'Casos escalados por fabricante', 1], ['4.4', 'Casos por gestor', 1]);
    }
    for (const [num, title, lvl] of toc) {
      children.push(para([run(`${num}  `, { bold: lvl === 0 }), run(title, { bold: lvl === 0 })], { indent: { left: lvl * 400 }, spacing: { after: 60 } }));
    }

    /* ---------- 1 Introducción ---------- */
    children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    children.push(h1('1 INTRODUCCIÓN'));
    children.push(text(
      `El presente documento recoge la actividad del servicio SLM-OSS (${meta.subtitle || 'Gestores Propietarios'}) ` +
      `correspondiente a la ${weekTitle.toLowerCase()} (${weekDates}, ${monthName})` +
      (n > 1 ? `, junto con la evolución de las semanas ${weeks[0].week} a ${weeks[n - 2].week} (desde el ${C.fmtDate(weeks[0].start)}).` : '.') +
      ' Cada semana refleja la situación de los casos al cierre de esa semana.', { size: 21 }, { spacing: { after: 120 } }));
    children.push(text(
      'Las tablas y gráficos presentados en este informe tienen como base el fichero de BRISE que nos envía semanalmente el CC – OSS. ' +
      `Datos disponibles hasta el ${C.fmtDateTime(meta.dataUntil)}.` +
      (report.partialWeek ? ` La semana ${rw.week} se presenta con los datos disponibles a esa fecha.` : ''), { size: 21 }, { spacing: { after: 120 } }));

    /* ---------- 2 Detalle semanal ---------- */
    children.push(h1('2 DETALLE SEMANAL'));
    children.push(h2('2.1 Casos tratados'));
    children.push(dualTable('CASOS NUEVOS SEMANA', newRows), spacer());
    children.push(dualTable('CASOS DEL BACKLOG NO ESCALADOS', backlogRows('backlogNoEsc')), spacer());
    children.push(dualTable('CASOS DEL BACKLOG ESCALADOS', backlogRows('backlogEsc')), spacer());
    if ((images.charts || {}).weekly) {
      const img = images.charts.weekly.bar;
      const scale = Math.min(1, 620 / img.width);
      children.push(para(new D.ImageRun({ type: 'png', data: img.bytes, transformation: { width: Math.round(img.width * scale), height: Math.round(img.height * scale) } }),
        { alignment: D.AlignmentType.CENTER, spacing: { after: 120 } }));
    }
    children.push(text('Las tablas y gráficos presentados en este informe tienen como base el fichero de BRISE que nos envía semanalmente el CC – OSS',
      { italics: true, size: 17, color: '606060' }));

    children.push(h2('2.2 Nº Casos por gestor'));
    if (report.gestores.length) {
      children.push(gestorTable());
      const notes = ['Incidencias nuevas de cada semana agrupadas por el gestor indicado en la acción actual (GESTOR - PROBLEMA - TÉCNICO); el total de cada semana coincide con “Nuevos durante la semana” de la tabla 2.1.'];
      if (report.gestorUnidentified) notes.push(`“Sin gestor identificado”: incidencias cuya acción actual no indica gestor (técnico trabajando, sin acción o devueltas).`);
      if (report.gestorInferred) notes.push(`${report.gestorInferred} incidencia(s) sin gestor en la acción actual se han asignado según el gestor citado en su descripción.`);
      children.push(text(notes.join(' '), { italics: true, size: 17, color: '606060' }, { spacing: { before: 80 } }));
    } else {
      children.push(text('No hay incidencias nuevas en las semanas del informe.', { italics: true }));
    }
    children.push(spacer());

    children.push(h2('2.3 Casos sin resolver'));
    children.push(dualTable('CASOS SIN RESOLVER', unresolvedRows), spacer());

    children.push(h2('2.4 Detalle de los casos sin resolver'));
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
        ['Tipo', t.type],
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

    children.push(h2('2.5 Incidencias destacadas'));
    bulletList(meta.highlights);
    children.push(h2('2.6 Otros trabajos a destacar'));
    bulletList(meta.otherWork);

    function bulletList(items) {
      const list = (items || []).map((s) => s.trim()).filter(Boolean);
      if (!list.length) { children.push(text('Sin elementos a destacar.', { italics: true, color: '606060' })); return; }
      for (const it of list) children.push(new D.Paragraph({ children: [run(it, { size: 21 })], bullet: { level: 0 }, spacing: { after: 60 } }));
    }

    /* ---------- 3 Detalle mensual ---------- */
    children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    children.push(h1('3 DETALLE MENSUAL'));
    children.push(text(`Evolución de los últimos ${report.months.length} meses hasta ${trendMonth}, con datos hasta el cierre de la ${weekTitle.toLowerCase()}.`, { size: 20 }, { spacing: { after: 120 } }));
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
          { alignment: D.AlignmentType.CENTER, spacing: { after: 160 } }));
      }
    };
    children.push(h2('3.1 Casos abiertos'));
    children.push(intro('abiertos', 'la prioridad'));
    addCharts('opened');
    children.push(h2('3.2 Casos resueltos'));
    children.push(intro('resueltos', 'la prioridad'));
    addCharts('resolved');
    children.push(h2('3.3 Casos escalados'));
    children.push(h3('3.3.1 Casos escalados por prioridad'));
    children.push(intro('escalados', 'la prioridad'));
    addCharts('escalated');
    children.push(h3('3.3.2 Casos escalados por gestor'));
    children.push(intro('escalados', 'el gestor'));
    addCharts('escByGestor');
    children.push(h3('3.3.3 Casos escalados por fabricante'));
    children.push(intro('escalados', 'el fabricante (STA- Ericsson, H- Huawei, 1- Nokia)'));
    addCharts('escByVendor');
    children.push(h2('3.4 Estado de los casos'));
    children.push(intro('abiertos', 'el estado'));
    addCharts('status');

    /* ---------- 4 Resumen del periodo (Month / Year filter) ---------- */
    if (meta.period && meta.period.data) {
      const P = meta.period.data;
      const I = P.summary.INC, O = P.summary.OT;
      const pc = images.periodCharts || {};
      children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
      children.push(h1(`4 RESUMEN DEL PERIODO: ${P.label.toUpperCase()}`));
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
      const img = (x) => {
        if (!x) return;
        const scale = Math.min(1, 620 / x.width);
        children.push(para(new D.ImageRun({ type: 'png', data: x.bytes, transformation: { width: Math.round(x.width * scale), height: Math.round(x.height * scale) } }),
          { alignment: D.AlignmentType.CENTER, spacing: { after: 160 } }));
      };
      children.push(h2('4.1 Casos abiertos y resueltos'));
      img(pc.opened); img(pc.resolved);
      children.push(h2('4.2 Estado de los casos'));
      img(pc.status);
      children.push(h2('4.3 Casos escalados por fabricante'));
      img(pc.vendor);
      children.push(h2('4.4 Casos por gestor'));
      img(pc.gestor);
      if (meta.period.gestorCaption) children.push(text(meta.period.gestorCaption, { italics: true, size: 17, color: '606060' }, { spacing: { after: 120 } }));
      const GW4 = [CONTENT_W - 3 * 1400, 1400, 1400, 1400];
      const tot = P.gestores.reduce((a, g) => [a[0] + g.inc, a[1] + g.ot, a[2] + g.total], [0, 0, 0]);
      children.push(table([
        new D.TableRow({ tableHeader: true, children: ['GESTOR', 'Nº Incidencias', 'Nº OTs', 'Total'].map((h, i) =>
          cell(h, { width: GW4[i], fill: i ? ORANGE : DARK, color: 'FFFFFF', bold: true, left: i === 0 })) }),
        ...P.gestores.map((g, idx) => new D.TableRow({ children: [
          cell(g.gestor, { width: GW4[0], left: true, fill: idx % 2 ? GREY : undefined, color: g.unidentified ? '606060' : undefined }),
          cell(g.inc, { width: GW4[1], fill: idx % 2 ? GREY : undefined }),
          cell(g.ot, { width: GW4[2], fill: idx % 2 ? GREY : undefined }),
          cell(g.total, { width: GW4[3], bold: true, fill: idx % 2 ? GREY : undefined }),
        ] })),
        new D.TableRow({ children: ['TOTAL', ...tot].map((v, i) => cell(v, { width: GW4[i], bold: true, left: i === 0, fill: 'E4E7EC' })) }),
      ], GW4));
      children.push(text('El gestor se obtiene de la acción actual (GESTOR - PROBLEMA - TÉCNICO); los casos sin gestor en la acción aparecen como “Sin gestor identificado”. El total coincide con los casos abiertos en el periodo.',
        { italics: true, size: 17, color: '606060' }, { spacing: { before: 80 } }));
    }

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
