/* Monitor de conciliación YPF – interfaz y datos.
   Usa YPFEngine (engine.js), SheetJS (XLSX) y, si hay configuración, Firebase (login con Google + Firestore). */
(function () {
  'use strict';
  const E = window.YPFEngine;
  const $ = (s, el = document) => el.querySelector(s);
  const app = $('#app');

  /* ---------- utilidades ---------- */
  const nf = new Intl.NumberFormat('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const money = (x) => '$ ' + nf.format(x || 0);
  const mill = (x) => '$ ' + ((x || 0) / 1e6).toLocaleString('es-AR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' M';
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const dmy = (iso) => (iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '');
  const hoyISO = () => new Date().toLocaleDateString('sv-SE', { timeZone: 'America/Argentina/Buenos_Aires' });
  const fechaHora = (iso) => { try { return new Date(iso).toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch (e) { return iso; } };
  const MEDIO = { certribr: 'Retención (cert.)', chdifele: 'Echeq diferido', chdifter: 'Cheque dif. terceros', chrecter: 'Cheque común terceros', chrecele: 'Echeq', chdifrec: 'Cheque dif. recibido', efectivo: 'Efectivo / depósito', redo: 'Redondeo' };
  const medio = (t) => MEDIO[t] || t;
  const nf0 = new Intl.NumberFormat('es-AR', { maximumFractionDigits: 0 });
  const DIVS = { GR: 'GR · Retenciones', DE: 'DE · Derivados (pagos por convenio)', YF: 'YF · YPF Ruta' };
  const ETQ = { PD: 'Pago directo de cliente', NOCONT: 'Pago no contabilizado en Nodum', RETDUP: 'Retención duplicada', SALDO: 'Saldo de compensación', OTRO: 'Otro' };
  const ETQF = { DEBITADO: 'Débito automático informado (pendiente de reflejo en YPF)' };
  const CHRE = { rech: 'Rechazado', recup: 'Recuperado', cerr: 'Cerrado (form Cheques Devuelto Proveedor)' };
  const parseImp = (x) => { let t = String(x || '').trim().replace(/[$\s]/g, ''); if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.'); else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, ''); const n = parseFloat(t); return isNaN(n) ? null : Math.round(n * 100) / 100; };
  let toastT;
  function toast(t) { let el = $('.toast'); if (!el) { el = document.createElement('div'); el.className = 'toast'; el.setAttribute('role', 'status'); document.body.appendChild(el); } el.textContent = t; el.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { el.hidden = true; }, 3200); }

  /* ---------- estado ---------- */
  const S = {
    user: null, permitido: false, modo: 'local', listo: { p: false, r: false, c: false },
    pagosDocs: {}, recibosDocs: {}, corridas: [], lineas: {}, recibos: {},
    files: { extranet: null, qCO: null, qPQ: null, nCO: null, nPQ: null }, fileMsgs: [],
    vista: (() => { try { return localStorage.getItem('ypf-vista') || 'dir'; } catch (e) { return 'dir'; } })(), reporte: 'real', tab: 'pend', filtroExt: 'revisar', editando: null, cargaAbierta: false, ocupado: false, error: null,
  };
  function rebuild() {
    S.lineas = {}; S.recibos = {};
    for (const d of Object.values(S.pagosDocs)) Object.assign(S.lineas, d.lineas || {});
    for (const d of Object.values(S.recibosDocs)) Object.assign(S.recibos, d.items || {});
  }

  /* ---------- almacenamiento ---------- */
  const cfg = window.FIREBASE_CONFIG;
  const usarFirebase = !!(cfg && cfg.apiKey && window.firebase);
  const firebaseCaido = !!(cfg && cfg.apiKey && !window.firebase);
  let fb = null;

  const Store = {
    async init() {
      if (usarFirebase) {
        S.modo = 'firebase';
        firebase.initializeApp(cfg);
        fb = { auth: firebase.auth(), db: firebase.firestore() };
        // Redes con proxy o antivirus cortan la conexión en vivo de Firestore: long polling es más robusto
        try { fb.db.settings({ experimentalForceLongPolling: true, merge: true }); } catch (e) { /* ya configurado */ }
        fb.auth.onAuthStateChanged(async (u) => {
          S.user = u; S.permitido = false; S.error = null;
          if (u) {
            try {
              const p = await fb.db.doc('permitidos/' + (u.email || '').toLowerCase()).get();
              S.permitido = p.exists;
            } catch (e) { S.permitido = false; }
            if (S.permitido) Store.subscribe();
          }
          render();
        });
      } else if (firebaseCaido) {
        S.modo = 'firebase'; S.user = null;
        app.innerHTML = '<div class="msg err">No se pudo cargar Firebase. Revisá la conexión y recargá la página.</div>';
        return;
      } else {
        S.modo = 'local'; S.user = { email: 'modo local (sin Firebase)' }; S.permitido = true;
        try { const j = JSON.parse(localStorage.getItem('ypf-monitor-local') || 'null'); if (j) { S.pagosDocs = j.pagos || {}; S.recibosDocs = j.recibos || {}; S.corridas = (j.corridas || []).map(decodeCorrida); } } catch (e) { /* sin almacenamiento */ }
        rebuild(); S.listo = { p: true, r: true, c: true };
        render();
      }
    },
    subscribe() {
      if (Store._subs) return;
      const db = fb.db; const onErr = (e) => { S.error = 'No pude leer la base: ' + e.message; render(); };
      Store._subs = [
        db.collection('pagos').onSnapshot((s) => { S.pagosDocs = {}; s.forEach((d) => { S.pagosDocs[d.id] = d.data(); }); S.listo.p = true; rebuild(); render(); }, onErr),
        db.collection('recibos').onSnapshot((s) => { S.recibosDocs = {}; s.forEach((d) => { S.recibosDocs[d.id] = d.data(); }); S.listo.r = true; rebuild(); render(); }, onErr),
        db.collection('corridas').orderBy('fecha', 'desc').limit(40).onSnapshot((s) => { S.corridas = s.docs.map((d) => decodeCorrida({ id: d.id, ...d.data() })); S.listo.c = true; render(); }, onErr),
      ];
    },
    async recargar() {
      if (S.modo !== 'firebase') return;
      const db = fb.db;
      const [p, r, c] = await Promise.all([db.collection('pagos').get(), db.collection('recibos').get(), db.collection('corridas').orderBy('fecha', 'desc').limit(40).get()]);
      S.pagosDocs = {}; p.forEach((d) => { S.pagosDocs[d.id] = d.data(); });
      S.recibosDocs = {}; r.forEach((d) => { S.recibosDocs[d.id] = d.data(); });
      S.corridas = c.docs.map((d) => decodeCorrida({ id: d.id, ...d.data() }));
      S.listo = { p: true, r: true, c: true }; rebuild(); render();
    },
    saveLocal() { try { localStorage.setItem('ypf-monitor-local', JSON.stringify({ pagos: S.pagosDocs, recibos: S.recibosDocs, corridas: S.corridas.map(encodeCorrida) })); } catch (e) { /* lleno o bloqueado */ } },
    async commit(pagosDocs, recibosDocs, corrida) {
      if (S.modo === 'firebase') {
        const b = fb.db.batch();
        for (const [id, d] of Object.entries(pagosDocs)) b.set(fb.db.doc('pagos/' + id), d);
        for (const [id, d] of Object.entries(recibosDocs)) b.set(fb.db.doc('recibos/' + id), d);
        b.set(fb.db.doc('corridas/' + corrida.id), encodeCorrida(corrida));
        await b.commit();
        await Store.recargar();
      } else {
        Object.assign(S.pagosDocs, pagosDocs); Object.assign(S.recibosDocs, recibosDocs);
        S.corridas = [corrida, ...S.corridas]; rebuild(); Store.saveLocal(); render();
      }
    },
    async patchPago(docId, key, patch) {
      if (S.modo === 'firebase') {
        const p = {}; for (const [k, v] of Object.entries(patch)) p[k] = v === null ? firebase.firestore.FieldValue.delete() : v;
        await fb.db.doc('pagos/' + docId).set({ lineas: { [key]: p } }, { merge: true });
        await Store.recargar();
      } else {
        const d = S.pagosDocs[docId]; const l = d && d.lineas[key]; if (!l) return;
        for (const [k, v] of Object.entries(patch)) { if (v === null) delete l[k]; else l[k] = v; }
        rebuild(); Store.saveLocal(); render();
      }
    },
    async addPago(docId, key, line) {
      if (S.modo === 'firebase') {
        await fb.db.doc('pagos/' + docId).set({ unidad: line.u, mes: (line.f || '').slice(0, 7), lineas: { [key]: line } }, { merge: true });
        await Store.recargar();
      } else {
        const d = S.pagosDocs[docId] = S.pagosDocs[docId] || { unidad: line.u, mes: (line.f || '').slice(0, 7), lineas: {} };
        d.lineas[key] = line; rebuild(); Store.saveLocal(); render();
      }
    },
    async patchRecibo(docId, nro, patch) {
      if (S.modo === 'firebase') {
        const p = {}; for (const [k, v] of Object.entries(patch)) p[k] = v === null ? firebase.firestore.FieldValue.delete() : v;
        await fb.db.doc('recibos/' + docId).set({ items: { [nro]: p } }, { merge: true });
        await Store.recargar();
      } else {
        S.recibosDocs[docId] = S.recibosDocs[docId] || { mes: docId, items: {} };
        const r = S.recibosDocs[docId].items[nro] = S.recibosDocs[docId].items[nro] || { nro };
        for (const [k, v] of Object.entries(patch)) { if (v === null) delete r[k]; else r[k] = v; }
        rebuild(); Store.saveLocal(); render();
      }
    },
    async importar(seed) {
      const corridas = Object.entries(seed.corridas || {}).map(([id, c]) => ({ id, ...c }));
      if (S.modo === 'firebase') {
        const b = fb.db.batch();
        for (const [id, d] of Object.entries(seed.pagos || {})) b.set(fb.db.doc('pagos/' + id), d);
        for (const [id, d] of Object.entries(seed.recibos || {})) b.set(fb.db.doc('recibos/' + id), d);
        for (const c of corridas) b.set(fb.db.doc('corridas/' + c.id), encodeCorrida(c));
        await b.commit();
        await Store.recargar();
      } else {
        S.pagosDocs = seed.pagos || {}; S.recibosDocs = seed.recibos || {}; S.corridas = corridas.map(decodeCorrida).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
        rebuild(); Store.saveLocal(); render();
      }
    },
  };
  // Firestore no admite listas dentro de listas: la tabla de la Extranet se guarda como texto JSON
  function encodeCorrida(c) { const o = { ...c }; if (Array.isArray(o.extranet)) o.extranet = JSON.stringify(o.extranet); return o; }
  function decodeCorrida(c) { const o = { ...c }; if (typeof o.extranet === 'string') { try { o.extranet = JSON.parse(o.extranet); } catch (e) { o.extranet = []; } } return o; }

  /* ---------- carga de archivos ---------- */
  async function leerArchivos(list) {
    S.fileMsgs = [];
    for (const f of list) {
      try {
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const p = E.detect(XLSX, wb, f.name);
        if (p.kind === 'extranet') {
          S.files.extranet = p;
          if (p.sinLeer && p.sinLeer.length) S.fileMsgs.push({ t: 'warn', m: `${f.name}: ${p.sinLeer.length} fila(s) con importe que no pude leer porque no tienen número de documento (${p.sinLeer.map((o) => (o.tipo || 'sin tipo') + ' ' + money(o.orig)).join(' · ')}). Revisá el archivo antes de conciliar: el total no va a coincidir con la Extranet.` });
        }
        else if (p.kind === 'query' && p.u) S.files['q' + p.u] = p;
        else if (p.kind === 'nav' && p.u) S.files['n' + p.u] = p;
        else if (p.kind === 'error') S.fileMsgs.push({ t: 'err', m: `${f.name}: ${p.error}` });
        else S.fileMsgs.push({ t: 'err', m: `${f.name}: no lo reconozco. Espero la Deuda de la Extranet, la query "Medio pago por titular y rango" o el navegador contable de Nodum.` });
      } catch (e) { S.fileMsgs.push({ t: 'err', m: `${f.name}: no se pudo leer (${e.message}).` }); }
    }
    render();
  }

  async function conciliar() {
    const F = S.files;
    if (!F.extranet) return;
    S.ocupado = true; render();
    try {
      // Siempre conciliar contra lo último guardado en la base, nunca contra una pantalla desactualizada
      if (S.modo === 'firebase') await Store.recargar();
      const today = hoyISO();
      const nowISO = new Date().toISOString();
      const inputs = { extranet: F.extranet, queries: [F.qCO, F.qPQ].filter(Boolean), navs: [F.nCO, F.nPQ].filter(Boolean) };
      const r = E.ingest({ lineas: S.lineas, recibos: S.recibos }, inputs, today, nowISO);
      const id = 'c' + nowISO.replace(/[-:T]/g, '').slice(0, 14);
      const corrida = { id, ...r.corrida, usuario: (S.user && S.user.email) || '' };
      const v = E.analizar(r.lineas, r.recibos, corrida, S.corridas[0] || null, today);
      corrida.resumen = { deudaYPF: v.deudaYPF, paCO: v.paCO, paPQ: v.paPQ, deudaReal: v.deudaReal, retdup: v.retdupTotal, sinIdentificar: v.sin.length };
      const pagosDocs = {};
      for (const docId of r.chP) {
        const [u, ...m] = docId.split('-');
        pagosDocs[docId] = { unidad: u, mes: m.join('-'), lineas: {} };
      }
      for (const [k, l] of Object.entries(r.lineas)) { const d = E.pagoDocId(l); if (pagosDocs[d]) pagosDocs[d].lineas[k] = l; }
      const recibosDocs = {};
      for (const docId of r.chR) recibosDocs[docId] = { mes: docId, items: {} };
      for (const [n, rec] of Object.entries(r.recibos)) { const d = E.reciboDocId(rec); if (recibosDocs[d]) recibosDocs[d].items[n] = rec; }
      await Store.commit(pagosDocs, recibosDocs, corrida);
      S.files = { extranet: null, qCO: null, qPQ: null, nCO: null, nPQ: null }; S.cargaAbierta = false; S.tab = 'pend';
      toast('Conciliación guardada');
    } catch (e) {
      S.fileMsgs = [{ t: 'err', m: 'No se pudo guardar la conciliación: ' + e.message }];
    }
    S.ocupado = false; render();
  }

  /* ---------- acciones manuales ---------- */
  const quien = () => (S.user && S.user.email) || '';
  async function accionPago(key, accion) {
    const l = S.lineas[key]; if (!l) return;
    const docId = E.pagoDocId(l);
    const hoy = hoyISO();
    try {
      if (accion === 'ok') await Store.patchPago(docId, key, { est: 'OK', okVia: 'manual', okF: hoy, okBy: quien() });
      if (accion === 'chr') await Store.patchPago(docId, key, { est: 'CHR', okF: hoy, okBy: quien(), chrE: 'rech', chrF: hoy, chrBy: quien() });
      if (accion === 'pa') await Store.patchPago(docId, key, { est: 'PA', okVia: null, okF: null, okBy: null, chrE: null, chrF: null, chrBy: null });
      toast(accion === 'pa' ? 'Volvió a pendiente' : accion === 'ok' ? 'Marcado como acreditado' : 'Marcado como rechazado');
    } catch (e) { toast('No se pudo guardar: ' + e.message); }
  }
  async function guardarRecibo(nro, fdoc) {
    const rec = S.recibos[nro] || {};
    const docId = (rec.f || fdoc || '0000-00').slice(0, 7);
    const val = (id) => { const el = $('#' + id + '-' + CSS.escape(nro)); return el ? el.value : ''; };
    const etq = val('etq'), nota = val('nota').trim(), vinc = val('vinc'), uni = val('uni');
    const base = { nota: nota || null, by: quien(), u: uni || null, uBy: uni ? quien() : null };
    if (!rec.f && fdoc) base.f = fdoc;
    if (!rec.nro) base.nro = nro;
    try {
      if (vinc) {
        const l = S.lineas[vinc];
        await Store.patchRecibo(docId, nro, { ...base, match: vinc, mt: 'manual', mf: hoyISO(), etq: null, sug: null });
        if (l && l.est === 'PA') await Store.patchPago(E.pagoDocId(l), vinc, { est: 'OK', okVia: nro, okF: hoyISO(), okBy: quien() });
      } else {
        await Store.patchRecibo(docId, nro, { ...base, etq: etq || null });
      }
      S.editando = null; render(); toast('Partida actualizada');
    } catch (e) { toast('No se pudo guardar: ' + e.message); }
  }
  async function guardarChr(key) {
    const l = S.lineas[key]; if (!l) return;
    const q = (attr) => document.querySelector(`[${attr}="${CSS.escape(key)}"]`);
    const est = q('data-chre').value, nota = q('data-chrn').value.trim();
    try {
      await Store.patchPago(E.pagoDocId(l), key, { chrE: est, chrNota: nota || null, chrF: hoyISO(), chrBy: quien() });
      toast(est === 'cerr' ? 'Cheque cerrado' : 'Cheque actualizado');
    } catch (e) { toast('No se pudo guardar: ' + e.message); }
  }
  async function agregarChr() {
    const g = (id) => ($('#' + id) ? $('#' + id).value.trim() : '');
    const u = g('nc-u'), f = g('nc-f'), tipo = g('nc-tipo') || 'chdifter', nro = parseInt(g('nc-nro'), 10), imp = parseImp(g('nc-imp'));
    const trans = parseInt(g('nc-trans'), 10) || 0;
    if (!u || !/^\d{4}-\d{2}-\d{2}$/.test(f) || !nro || !imp) { toast('Completá unidad, fecha, número e importe'); return; }
    const hoy = hoyISO();
    const line = { u, f, trans, tipo, serie: '0', nro, imp, est: 'CHR', ref: g('nc-ref'), ant: parseInt(g('nc-ant'), 10) || null, pv: hoy, orig: 'manual',
      okF: hoy, okBy: quien(), chrE: g('nc-est') || 'rech', chrNota: g('nc-nota') || null, chrF: hoy, chrBy: quien() };
    const key = E.lineKey(u, trans, tipo, nro, imp);
    if (S.lineas[key]) { toast('Ese cheque ya está cargado'); return; }
    try { await Store.addPago(E.pagoDocId(line), key, line); S.chrForm = false; render(); toast('Cheque rechazado agregado'); }
    catch (e) { toast('No se pudo guardar: ' + e.message); }
  }
  async function desvincular(nro) {
    const rec = S.recibos[nro]; if (!rec) return;
    try {
      await Store.patchRecibo((rec.f || '0000-00').slice(0, 7), nro, { match: null, mt: null, mf: null, etq: null });
      const l = rec.match && S.lineas[rec.match];
      if (l && l.okVia === nro) await Store.patchPago(E.pagoDocId(l), rec.match, { est: 'PA', okVia: null, okF: null, okBy: null });
      S.editando = null; render(); toast('Recibo desvinculado');
    } catch (e) { toast('No se pudo guardar: ' + e.message); }
  }

  /* ---------- Excel ---------- */
  /* ---------- Excel con formato (ExcelJS) ---------- */
  const XL = {
    rojo: 'FFD9141C', negro: 'FF111111', gris: 'FF595959', linea: 'FFD9D9D9', suave: 'FFF2F2F2',
    CO: 'FFC4D79B', PQ: 'FFFFC000', venc: 'FFC00000', hoy: 'FFC65911', ok: 'FF2E7D32',
  };
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const lin = (argb) => ({ style: 'thin', color: { argb } });
  const BORDE = { top: lin('FFBFBFBF'), left: lin('FFBFBFBF'), bottom: lin('FFBFBFBF'), right: lin('FFBFBFBF') };
  const BORDE_H = { top: lin('FF111111'), left: lin('FF595959'), bottom: lin('FF111111'), right: lin('FF595959') };
  const FMT = { num: '#,##0.00;[Red]-#,##0.00', int: '0', date: 'dd/mm/yyyy', usd: '#,##0.00' };
  const xDate = (iso) => (iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10), 12)) : null);
  const uNombre = (u) => ({ CO: 'Combustibles', PQ: 'Petroquímicos' }[u] || 'Sin asignar');

  function xHoja(wb, nombre, titulo, sub) {
    const ws = wb.addWorksheet(nombre, { views: [{ showGridLines: false }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } } });
    ws.headerFooter.oddFooter = '&LCopsa · Conciliación YPF&RPágina &P de &N';
    const t = ws.addRow([titulo]); t.font = { bold: true, size: 15, color: { argb: XL.rojo } }; t.height = 22;
    if (sub) { const r = ws.addRow([sub]); r.font = { italic: true, size: 10, color: { argb: XL.gris } }; }
    ws.addRow([]);
    ws._anchos = [];
    return ws;
  }
  // cols: [{h, t: 'txt'|'num'|'int'|'date'|'usd', w}]
  function xTabla(ws, cols, rows, opt = {}) {
    if (opt.titulo) { const r = ws.addRow([opt.titulo]); r.font = { bold: true, size: 12, color: { argb: XL.negro } }; }
    const hr = ws.addRow(cols.map((c) => c.h));
    hr.height = 20;
    hr.eachCell((cell, i) => {
      const c = cols[i - 1];
      cell.font = { bold: true, color: { argb: c.headFill ? XL.negro : 'FFFFFFFF' } };
      cell.fill = fill(c.headFill || XL.negro);
      cell.alignment = { vertical: 'middle', horizontal: c.t === 'txt' || c.t === 'date' ? 'left' : 'right', wrapText: true };
      cell.border = BORDE_H;
    });
    const first = hr.number + 1;
    for (const row of rows) {
      const data = Array.isArray(row) ? row : row.values;
      const r = ws.addRow(data.map((v, i) => (cols[i] && cols[i].t === 'date' ? xDate(v) : v)));
      r.eachCell({ includeEmpty: true }, (cell, i) => {
        const c = cols[i - 1]; if (!c) return;
        if (FMT[c.t]) cell.numFmt = FMT[c.t];
        cell.border = BORDE;
        cell.alignment = { vertical: 'top', horizontal: c.t === 'txt' || c.t === 'date' ? 'left' : 'right', indent: (row.indent && i === 1) ? 2 : 0, wrapText: !!c.wrap };
      });
      if (row.font) r.eachCell((cell) => { cell.font = { ...(cell.font || {}), ...row.font }; });
      if (row.total) r.eachCell({ includeEmpty: true }, (cell, i) => { if (!cols[i - 1]) return; cell.font = { ...(cell.font || {}), bold: true }; cell.border = { left: BORDE.left, right: BORDE.right, top: { style: 'thin', color: { argb: XL.negro } }, bottom: { style: 'double', color: { argb: XL.negro } } }; cell.fill = fill(XL.suave); });
      if (row.cellFont) for (const [ci, f] of Object.entries(row.cellFont)) r.getCell(+ci).font = { ...(r.getCell(+ci).font || {}), ...f };
    }
    const last = ws.lastRow.number;
    if (opt.filtro && rows.length) ws.autoFilter = { from: { row: first - 1, column: 1 }, to: { row: last, column: cols.length } };
    if (opt.congelar) ws.views = [{ state: 'frozen', ySplit: first - 1, showGridLines: false }];
    cols.forEach((c, i) => { ws._anchos[i] = Math.max(ws._anchos[i] || 0, c.w || 12); ws.getColumn(i + 1).width = ws._anchos[i]; });
    ws.addRow([]);
    return { first, last };
  }

  async function descargarExcel(v, corr) {
    if (!window.ExcelJS) { toast('No se pudo cargar el generador de Excel. Revisá la conexión y recargá.'); return; }
    const mm = S._m || metricas(v);
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Monitor de conciliación YPF – Copsa'; wb.created = new Date();
    const fechaTxt = fechaHora(corr.fecha);
    const sub = `Conciliación del ${fechaTxt}${corr.usuario ? ' · ' + corr.usuario : ''}`;
    const sinA = Math.abs(mm.porU['—'] || 0) > 0.005;
    const uCols = [{ h: 'Combustibles', t: 'num', w: 20 }, { h: 'Petroquímicos', t: 'num', w: 20 }, ...(sinA ? [{ h: 'Sin asignar', t: 'num', w: 16 }] : []), { h: 'Total', t: 'num', w: 20 }];
    const uv = (co, pq, sa, tot) => (sinA ? [co, pq, sa, tot] : [co, pq, tot]);
    const rdU = (u) => sumBy(v.retdup.filter((f) => mm.uOf(f) === u), (f) => f.rec.imp);

    // Resumen
    let ws = xHoja(wb, 'Resumen', 'Cuenta corriente con YPF · Resumen', sub);
    xTabla(ws, [{ h: 'Concepto', t: 'txt', w: 50 }, ...uCols], [
      ['Deuda según YPF (Extranet, original en pesos)', ...uv(mm.porU.CO, mm.porU.PQ, mm.porU['—'] || 0, v.deudaYPF)],
      ['Pendiente de acreditar', ...uv(-v.paCO, -v.paPQ, 0, -(v.paCO + v.paPQ))],
      { values: ['Deuda real', ...uv(mm.porU.CO - v.paCO, mm.porU.PQ - v.paPQ, mm.porU['—'] || 0, v.deudaReal)], total: true },
    ], { titulo: 'Deuda por unidad de negocio' });
    xTabla(ws, [{ h: 'Partidas a gestionar', t: 'txt', w: 50 }, ...uCols], [
      ['Retenciones duplicadas acreditadas por YPF (a anular)', ...uv(rdU('CO'), rdU('PQ'), 0, v.retdupTotal)],
      ['Recibos sin identificar', ...uv(null, null, null, sumBy(v.sin, (f) => f.rec.imp))],
      ['Comprobantes sin registrar en Nodum', ...uv(sumBy(v.noCargadas.filter((e) => e.u === 'CO'), 'orig'), sumBy(v.noCargadas.filter((e) => e.u === 'PQ'), 'orig'), 0, sumBy(v.noCargadas, 'orig'))],
      ['Pendiente de acreditar con más de 7 días', ...uv(sumBy(v.viejos.filter((p) => p.u === 'CO'), 'imp'), sumBy(v.viejos.filter((p) => p.u === 'PQ'), 'imp'), 0, sumBy(v.viejos, 'imp'))],
    ]);
    xTabla(ws, [{ h: 'División', t: 'txt', w: 50 }, { h: 'Unidad', t: 'txt', w: 20, unidad: true }, { h: 'Débitos', t: 'num', w: 20 }, { h: 'Créditos', t: 'num', w: 20 }, { h: 'Neto', t: 'num', w: 20 }, { h: 'Partidas', t: 'int', w: 10 }],
      [...v.divs.map((d) => [DIVS[d.div] || d.div, { CO: 'Combustibles', PQ: 'Petroquímicos' }[E.unidadDeDiv(d.div)] || 'Según el pago', d.debitos, d.creditos, d.debitos + d.creditos, d.n]),
        { values: ['Total', '', sumBy(v.divs, 'debitos'), sumBy(v.divs, 'creditos'), v.deudaYPF, sumBy(v.divs, 'n')], total: true }],
      { titulo: 'Deuda según YPF por división' });
    const nota = ws.addRow(['Unidad: las facturas se asignan por división (MN, MA, GR → Combustibles; PQ, ME → Petroquímicos); los pagos de la división DE (Derivados), por el pago de Nodum con el que se cruzaron, salvo las partidas reasignadas a mano.']);
    nota.font = { italic: true, size: 9, color: { argb: XL.gris } };
    const arch = ws.addRow([`Archivos: ${(corr.archivos || []).join(' · ')}`]); arch.font = { size: 9, color: { argb: XL.gris } };

    // Vencimientos
    ws = xHoja(wb, 'Vencimientos', 'Vencimientos de facturas y notas de débito', sub);
    const filasV = [];
    for (const g of mm.vencG) {
      const font = g.g === 'venc' && g.tot ? { color: { argb: XL.venc }, bold: true } : g.g === 'hoy' && g.tot ? { color: { argb: XL.hoy }, bold: true } : g.g === 'aven' ? { bold: true } : null;
      filasV.push({ values: [g.l, g.rows.length, g.co, g.pq, g.tot], font });
      if (g.g === 'aven') for (const b of mm.venc.filter((x) => x.g === 'aven')) filasV.push({ values: [b.l.replace('A vencer · ', ''), b.rows.length, b.co, b.pq, b.tot], indent: true, font: { color: { argb: XL.gris } } });
    }
    filasV.push({ values: ['Total', mm.deuda.length, sumBy(mm.vencG, 'co'), sumBy(mm.vencG, 'pq'), sumBy(mm.vencG, 'tot')], total: true });
    if (mm.deb.rows.length) filasV.push({ values: ['Débito automático informado (no se cuenta como vencido)', mm.deb.rows.length, mm.deb.co, mm.deb.pq, mm.deb.tot], font: { color: { argb: XL.gris } } });
    const colsV = [{ h: 'Estado', t: 'txt', w: 28 }, { h: 'Comprobantes', t: 'int', w: 14 }, { h: 'Combustibles', t: 'num', w: 20 }, { h: 'Petroquímicos', t: 'num', w: 20 }, { h: 'Total', t: 'num', w: 20 }];
    xTabla(ws, colsV, filasV, { titulo: 'Resumen' });
    xTabla(ws, colsV, [
      ['Exigible hoy (vencido + vence hoy)', mm.vencG[0].rows.length + mm.vencG[1].rows.length, mm.exig.co, mm.exig.pq, mm.exig.tot],
      ['Acreditado por YPF sin compensar', mm.acred.rows.length, mm.acred.co, mm.acred.pq, mm.acred.tot],
      { values: ['Exigible neto', null, mm.neto.co, mm.neto.pq, mm.neto.tot], total: true },
    ], { titulo: 'Exigible neto de lo acreditado' });
    const nn = ws.addRow(['Lo acreditado sin compensar son los recibos y saldos a favor que figuran en la Extranet (sin las retenciones duplicadas). El detalle está al final de esta hoja.']);
    nn.font = { italic: true, size: 9, color: { argb: XL.gris } }; ws.addRow([]);
    const det = [];
    for (const g of mm.vencG) for (const f of g.rows.slice().sort((x, y) => (x.fvenc < y.fvenc ? -1 : 1))) {
      const d = E.daysBetween(mm.hoy, f.fvenc);
      det.push({ values: [g.l, f.fvenc, Math.abs(d), uNombre(mm.uOf(f)), f.tipo, f.nro, f.div, f.mon, f.usd || null, f.orig], cellFont: g.g === 'venc' ? { 1: { color: { argb: XL.venc }, bold: true } } : g.g === 'hoy' ? { 1: { color: { argb: XL.hoy }, bold: true } } : null });
    }
    for (const f of mm.deb.rows) det.push({ values: ['Débito automático informado', f.fvenc, Math.abs(E.daysBetween(mm.hoy, f.fvenc)), uNombre(mm.uOf(f)), f.tipo, f.nro, f.div, f.mon, f.usd || null, f.orig], font: { color: { argb: XL.gris } } });
    xTabla(ws, [{ h: 'Estado', t: 'txt', w: 28 }, { h: 'Vence', t: 'date', w: 14 }, { h: 'Días (atraso / faltan)', t: 'int', w: 12 }, { h: 'Unidad', t: 'txt', w: 16, unidad: true }, { h: 'Tipo', t: 'txt', w: 15 }, { h: 'Nro. legal', t: 'txt', w: 16 }, { h: 'División', t: 'txt', w: 9 }, { h: 'Moneda', t: 'txt', w: 9 }, { h: 'USD', t: 'usd', w: 14 }, { h: 'Importe ARS', t: 'num', w: 20 }], det, { titulo: 'Detalle de vencimientos', filtro: true });
    const acr = [];
    for (const u of ['CO', 'PQ', '—']) {
      const rs = mm.acred.rows.filter((f) => mm.uOf(f) === u).sort((a, b) => (a.fdoc < b.fdoc ? -1 : 1)); if (!rs.length) continue;
      for (const f of rs) acr.push([uNombre(u) + (f.uManual ? ' (a mano)' : ''), f.nro, f.tpos, f.div, f.fdoc, f.pago ? f.pago.trans : null, f.pago ? f.pago.ant || null : null, f.pago ? f.pago.ref || '' : (f.rec && f.rec.etq ? ETQ[f.rec.etq] || f.rec.etq : ''), f.orig]);
      acr.push({ values: [`Total ${uNombre(u)}`, '', '', '', null, null, null, '', sumBy(rs, 'orig')], total: true });
    }
    if (acr.length) xTabla(ws, [{ h: 'Unidad', t: 'txt', w: 28 }, { h: 'Recibo', t: 'txt', w: 16 }, { h: 'T. Pos.', t: 'txt', w: 20 }, { h: 'Div.', t: 'txt', w: 9 }, { h: 'F. Doc.', t: 'date', w: 12 }, { h: 'Trans Nodum', t: 'int', w: 12 }, { h: 'Anticipo', t: 'int', w: 10 }, { h: 'Referencia', t: 'txt', w: 20 }, { h: 'Importe', t: 'num', w: 20 }], acr, { titulo: 'Acreditado por YPF sin compensar, por unidad' });

    // Pendientes de acreditar
    ws = xHoja(wb, 'Pendientes de acreditar', 'Pagos pendientes de acreditar por YPF', sub);
    const pend = [];
    for (const u of ['CO', 'PQ']) {
      const ps = v.pendientes.filter((p) => p.u === u); if (!ps.length) continue;
      for (const p of ps) pend.push({ values: [uNombre(u), p.f, p.dias, p.trans, p.ant || null, p.ref || '', medio(p.tipo), p.serie && p.serie !== '0' ? p.serie : '', p.nro || null, p.imp], cellFont: p.dias > 7 ? { 3: { color: { argb: XL.venc }, bold: true } } : null });
      pend.push({ values: [`Total ${uNombre(u)}`, null, null, null, null, '', '', '', null, sumBy(ps, 'imp')], total: true });
    }
    pend.push({ values: ['Total general', null, null, null, null, '', '', '', null, v.paCO + v.paPQ], total: true });
    xTabla(ws, [{ h: 'Unidad', t: 'txt', w: 22, unidad: true }, { h: 'Fecha', t: 'date', w: 12 }, { h: 'Días', t: 'int', w: 7 }, { h: 'Transacción', t: 'int', w: 12 }, { h: 'Anticipo', t: 'int', w: 10 }, { h: 'Referencia', t: 'txt', w: 22 }, { h: 'Medio de pago', t: 'txt', w: 22 }, { h: 'Serie', t: 'txt', w: 7 }, { h: 'Número', t: 'int', w: 12 }, { h: 'Importe', t: 'num', w: 18 }], pend, { congelar: true });

    // Extranet
    ws = xHoja(wb, 'Extranet', 'Partidas de la Extranet de YPF', sub);
    const colEst = { cruzado: XL.ok, parcial: XL.gris, retdup: XL.hoy, posibleDup: XL.hoy, sin: XL.venc, etiquetado: XL.gris };
    xTabla(ws, [{ h: 'Estado', t: 'txt', w: 22 }, { h: 'Tipo', t: 'txt', w: 14 }, { h: 'Nro. legal', t: 'txt', w: 16 }, { h: 'Clase', t: 'txt', w: 7 }, { h: 'Div.', t: 'txt', w: 6 }, { h: 'Unidad', t: 'txt', w: 15, unidad: true }, { h: 'T. Pos.', t: 'txt', w: 20 }, { h: 'F. Doc.', t: 'date', w: 12 }, { h: 'F. Venc.', t: 'date', w: 12 }, { h: 'Moneda', t: 'txt', w: 8 }, { h: 'Pendiente ARS (TC)', t: 'num', w: 18 }, { h: 'Pendiente USD', t: 'usd', w: 14 }, { h: 'Original ARS', t: 'num', w: 18 }, { h: 'Cruza con (Nodum)', t: 'txt', w: 46, wrap: true }, { h: 'Etiqueta', t: 'txt', w: 26 }, { h: 'Nota', t: 'txt', w: 28, wrap: true }],
      [...v.filas.map((f) => ({ values: [estadoTxt(f), f.tipo, f.nro, f.clase, f.div, uNombre(mm.uOf(f)) + (f.uManual ? ' (a mano)' : ''), f.tpos, f.fdoc, f.fvenc, f.mon, f.ars, f.usd || null, f.orig,
        f.pago ? `${uNombre(f.pago.u)} · trans ${f.pago.trans}${f.pago.ant ? ' · anticipo ' + f.pago.ant : ''} · ${medio(f.pago.tipo)} ${f.pago.nro || ''} · ${nf.format(f.pago.imp)}` : '', f.rec && f.rec.etq ? (f.recibo ? ETQ[f.rec.etq] || f.rec.etq : ETQF[f.rec.etq] || '') : '', (f.rec && f.rec.nota) || ''],
        cellFont: colEst[f.estado] ? { 1: { color: { argb: colEst[f.estado] }, bold: f.estado !== 'cruzado' } } : null })),
        { values: ['Total', '', '', '', '', '', '', null, null, '', v.deudaTC, sumBy(v.filas, (f) => f.usd || 0), v.deudaYPF, '', '', ''], total: true }],
      { filtro: true, congelar: true });

    // Excepciones
    ws = xHoja(wb, 'Excepciones', 'Excepciones a revisar', sub);
    const exc = [];
    const add = (tipo, det, imp, color) => exc.push({ values: [tipo, det, imp], cellFont: color ? { 1: { color: { argb: color }, bold: true } } : null });
    v.sin.forEach((f) => add('Recibo sin identificar', `${f.nro} · ${f.tpos} · ${dmy(f.fdoc)}`, f.rec.imp, XL.venc));
    v.posibleDup.forEach((f) => add('Posible retención duplicada', `${f.nro} · ${f.tpos} · ${dmy(f.fdoc)}`, f.rec.imp, XL.hoy));
    v.retdup.forEach((f) => add('Retención duplicada (solicitar anulación)', `${f.nro} · ${f.tpos} · ${dmy(f.fdoc)} · ${E.daysBetween(f.fdoc, mm.hoy)} días`, f.rec.imp, XL.hoy));
    v.parciales.forEach((f) => add('Saldo parcial', `${f.nro} · ${f.tpos}: YPF muestra ${nf.format(f.rec.imp)} de un pago de ${nf.format(f.pago.imp)}`, f.rec.imp));
    v.noCargadas.forEach((e) => add('Sin registrar en Nodum', `${e.tipo} ${e.nro} · ${uNombre(e.u)} · ${dmy(e.fdoc)}`, e.orig, XL.hoy));
    v.alertas.forEach((a) => add(a.tipo === 'difContable' ? 'Diferencia query / contable' : 'Pago que ya no figura en Nodum', a.texto, a.imp, XL.venc));
    v.viejos.forEach((p) => add('Pendiente de acreditar con más de 7 días', `${uNombre(p.u)} · trans ${p.trans} · ${medio(p.tipo)} ${p.nro || ''} · ${dmy(p.f)} (${p.dias} días)`, p.imp, XL.venc));
    if (!exc.length) exc.push(['Sin excepciones abiertas', '', null]);
    xTabla(ws, [{ h: 'Tipo', t: 'txt', w: 40 }, { h: 'Detalle', t: 'txt', w: 90, wrap: true }, { h: 'Importe', t: 'num', w: 18 }], exc, { filtro: true });

    // Pagos Nodum
    ws = xHoja(wb, 'Pagos Nodum', 'Pagos a YPF registrados en Nodum', sub);
    const EST = { PA: 'Pendiente de acreditar', OK: 'Acreditado', CHR: 'Cheque rechazado', R: 'Redondeo' };
    const estCol = { PA: XL.hoy, OK: XL.ok, CHR: XL.venc, R: XL.gris };
    const ls = Object.values(S.lineas).sort((a, b) => (a.u + a.f + String(a.trans).padStart(9, '0') > b.u + b.f + String(b.trans).padStart(9, '0') ? 1 : -1));
    xTabla(ws, [{ h: 'Unidad', t: 'txt', w: 16, unidad: true }, { h: 'Fecha', t: 'date', w: 12 }, { h: 'Transacción', t: 'int', w: 12 }, { h: 'Anticipo', t: 'int', w: 10 }, { h: 'Referencia', t: 'txt', w: 22 }, { h: 'Medio de pago', t: 'txt', w: 22 }, { h: 'Serie', t: 'txt', w: 7 }, { h: 'Número', t: 'int', w: 12 }, { h: 'Importe', t: 'num', w: 18 }, { h: 'Estado', t: 'txt', w: 22 }, { h: 'Acreditado con', t: 'txt', w: 16 }, { h: 'Fecha estado', t: 'date', w: 12 }, { h: 'Por', t: 'txt', w: 28 }],
      ls.map((l) => ({ values: [uNombre(l.u), l.f, l.trans, l.ant || null, l.ref || '', medio(l.tipo), l.serie && l.serie !== '0' ? l.serie : '', l.nro || null, l.imp, l.est === 'CHR' ? `Cheque rechazado · ${CHRE[l.chrE] || (l.chrE === 'nd' ? CHRE.recup : CHRE.rech)}` : EST[l.est] || l.est, l.okVia || '', l.okF || null, l.okBy || ''], cellFont: estCol[l.est] ? { 10: { color: { argb: estCol[l.est] }, bold: l.est !== 'OK' } } : null })),
      { filtro: true, congelar: true });

    // Cheques rechazados
    ws = xHoja(wb, 'Cheques rechazados', 'Cheques rechazados: seguimiento', sub);
    const chrRows = mm.chrs.map((c) => ({ values: [uNombre(c.u), c.f, c.trans || null, c.ant || null, c.ref || '', medio(c.tipo), c.nro || null, c.imp, CHRE[c.chrE], c.chrNota || '', c.chrF || null, c.chrBy || '', c.dias], cellFont: c.chrE !== 'cerr' ? { 9: { color: { argb: XL.venc }, bold: true } } : { 9: { color: { argb: XL.ok } } } }));
    chrRows.push({ values: ['Total abiertos', null, null, null, '', '', null, sumBy(mm.chrAb, 'imp'), `${mm.chrAb.length} cheques`, '', null, '', null], total: true });
    xTabla(ws, [{ h: 'Unidad', t: 'txt', w: 16 }, { h: 'Fecha', t: 'date', w: 12 }, { h: 'Transacción', t: 'int', w: 12 }, { h: 'Anticipo', t: 'int', w: 10 }, { h: 'Referencia', t: 'txt', w: 16 }, { h: 'Medio', t: 'txt', w: 22 }, { h: 'Número', t: 'int', w: 12 }, { h: 'Importe', t: 'num', w: 18 }, { h: 'Estado', t: 'txt', w: 24 }, { h: 'Justificación', t: 'txt', w: 36, wrap: true }, { h: 'Último cambio', t: 'date', w: 12 }, { h: 'Por', t: 'txt', w: 28 }, { h: 'Días', t: 'int', w: 8 }], chrRows, { filtro: true, congelar: true });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Conciliacion YPF ${(corr.fecha || new Date().toISOString()).slice(0, 10)}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }

  /* ---------- render ---------- */
  function estadoTxt(f) {
    return { deuda: f.clase === 'DD' ? 'Nota de débito' : /^Saldo/i.test(f.tipo) ? 'Saldo de factura' : 'Factura', cruzado: 'Cruzado con Nodum', parcial: 'Saldo parcial', retdup: 'Retención duplicada', etiquetado: 'Etiquetado', posibleDup: 'Posible ret. duplicada', sin: 'Sin identificar' }[f.estado] || f.estado;
  }
  function estadoPill(f) {
    const cls = { deuda: 'p-n', cruzado: 'p-ok', parcial: 'p-info', retdup: 'p-warn', etiquetado: 'p-info', posibleDup: 'p-warn', sin: 'p-bad' }[f.estado] || 'p-n';
    return `<span class="pill ${cls}">${esc(estadoTxt(f))}</span>`;
  }

const BRAND = `<div class="brandbar"><img class="logo" src="logo-copsa.jpg" alt="Compañía Petrolera COPSA S.A." width="153" height="44"><span class="sep" aria-hidden="true"></span><div class="who"><span class="lbl">Cuenta corriente con YPF</span><h1>Conciliación YPF</h1></div></div>`;
  function render() {
    try { renderInner(); } catch (e) {
      console.error(e);
      app.innerHTML = `${BRAND}<div class="msg err">Algo falló al mostrar el monitor: ${esc(e.message)}. Probá recargar la página; si sigue, avisá.</div>`;
    }
  }
  function renderInner() {
    if (S.modo === 'firebase' && !S.user) { app.innerHTML = vistaLogin(); return; }
    if (S.modo === 'firebase' && !S.permitido) { app.innerHTML = vistaNoHabilitado(); return; }
    const cargando = !(S.listo.p && S.listo.r && S.listo.c);
    const corr = S.corridas[0] || null;
    const prev = S.corridas[1] || null;
    const v = corr ? E.analizar(S.lineas, S.recibos, corr, prev, hoyISO()) : null;
    S._v = v; S._corr = corr;
    app.innerHTML = `
      ${BRAND}
      <header class="top">
        <div class="who">
          <span class="muted small">${corr ? `Última conciliación: <b>${esc(fechaHora(corr.fecha))}</b>${corr.usuario ? ' · ' + esc(corr.usuario) : ''}` : cargando ? 'Leyendo la base…' : 'Todavía no hay conciliaciones guardadas.'}</span>
        </div>
        <div class="actions">
          <button id="b-carga" class="${S.cargaAbierta || !corr ? '' : 'primary'}" data-a="toggleCarga">${S.cargaAbierta ? 'Cerrar carga' : 'Cargar archivos del día'}</button>
          <button id="b-xls" data-a="excel" ${v ? '' : 'disabled'}>Descargar Excel</button>
          <a class="btn" href="manual.html" target="_blank" rel="noopener">Manual</a>
          ${S.modo === 'firebase' ? `<button id="b-salir" data-a="salir" title="${esc(S.user.email)}">Salir</button>` : ''}
        </div>
      </header>
      ${S.modo === 'local' ? `<div class="msg warn">Modo local: no hay Firebase configurado. Lo que cargues queda solo en este navegador. <button class="mini" data-a="borrarLocal">Borrar datos de este navegador</button></div>` : ''}
      ${S.error ? `<div class="msg err">${esc(S.error)}</div>` : ''}
      <div class="viewsw" role="tablist" aria-label="Vista">
        <button role="tab" aria-selected="${S.vista === 'dir'}" data-a="vista" data-v="dir">Dirección</button>
        <button role="tab" aria-selected="${S.vista === 'conc'}" data-a="vista" data-v="conc">Conciliación</button>
      </div>
      ${(S.cargaAbierta || (!corr && !cargando)) ? vistaCarga(!corr) : ''}
      ${v ? vistaResultado(v, corr) : (!cargando && !corr ? `<div class="panel empty">Todavía no hay conciliaciones. ${S.vista === 'dir' ? 'Los informes aparecen después de la primera conciliación.' : 'Cargá los archivos del día para empezar.'}</div>` : `<div class="panel empty">Leyendo la base…</div>`)}
    `;
  }

  function vistaLogin() {
    return `${BRAND}
      <div class="panel carga"><h2>Ingresá con tu cuenta de Copsa</h2><p class="muted">Solo pueden entrar las cuentas habilitadas.</p><div><button class="primary" id="b-login" data-a="login">Ingresar con Google</button></div></div>`;
  }
  function vistaNoHabilitado() {
    return `${BRAND}<div class="actions"><button data-a="salir">Salir</button></div>
      <div class="panel carga"><h2>Tu cuenta no está habilitada</h2><p class="muted">Ingresaste como <b>${esc(S.user.email)}</b>. Pedile a Sandra que agregue este mail a la lista de habilitados.</p></div>`;
  }

  function vistaCarga(vacia) {
    const F = S.files;
    const slot = (k, titulo, ayuda) => {
      const p = F[k];
      const det = !p ? `<span class="muted small">${ayuda}</span>` :
        p.kind === 'extranet' ? `<span class="small">${p.rows.length} partidas · ${esc(money(p.total))}</span>` :
        p.kind === 'query' ? `<span class="small">${p.lines.length} medios de pago · ${dmy(p.desde)} al ${dmy(p.hasta)} · ${esc(money(p.total))}</span>` :
        `<span class="small">${p.comprobantes.length} comprobantes · ${Object.keys(p.pagos).length} pagos · ${dmy(p.desde)} al ${dmy(p.hasta)}</span>`;
      return `<div class="slot ${p ? 'ok' : ''}"><span class="lbl">${titulo}</span>${p ? `<span class="fn"><b>${esc(p.filename)}</b></span>` : ''}${det}</div>`;
    };
    const faltan = [];
    if (!F.qCO) faltan.push('query CO'); if (!F.qPQ) faltan.push('query PQ');
    return `<section class="panel carga" aria-label="Carga de archivos">
      <div class="carga-head"><h2>Cargar archivos del día</h2><span class="muted small">Se reconocen solos: arrastrá todos juntos.</span></div>
      ${vacia ? vistaImportar() : ''}
      <label class="drop" id="drop" for="f-in"><b>Soltá los archivos acá</b> o tocá para elegirlos<br><span class="muted small">Deuda de la Extranet (.xlsx) · Query "Medio pago por titular y rango" CO y PQ (.xls) · Navegador contable CO y PQ (.xlsx, opcional)</span></label>
      <input id="f-in" type="file" multiple accept=".xls,.xlsx" hidden>
      <div class="slots">
        ${slot('extranet', 'Extranet YPF (Deuda)', 'Obligatorio')}
        ${slot('qCO', 'Query Nodum · Combustibles', 'Pagos del mes, titular 3000266')}
        ${slot('qPQ', 'Query Nodum · Petroquímicos', 'Pagos del mes, titular 0001003')}
        ${slot('nCO', 'Navegador contable · CO', 'Opcional')}
        ${slot('nPQ', 'Navegador contable · PQ', 'Opcional')}
      </div>
      ${S.fileMsgs.map((m) => `<div class="msg ${m.t}">${esc(m.m)}</div>`).join('')}
      ${F.extranet && faltan.length ? `<div class="msg warn">Falta ${faltan.join(' y ')}: los pagos nuevos de esa unidad no van a entrar en esta conciliación.</div>` : ''}
      <div class="actions"><button class="primary" id="b-conc" data-a="conciliar" ${F.extranet && !S.ocupado ? '' : 'disabled'}>${S.ocupado ? 'Conciliando…' : 'Conciliar y guardar'}</button>
      ${Object.values(F).some(Boolean) ? '<button data-a="limpiar">Quitar archivos</button>' : ''}</div>
    </section>`;
  }
  function vistaImportar() {
    return `<div class="msg warn"><b>La base está vacía.</b> Si tenés el archivo de estado inicial (pendientes de la planilla manual), importalo una sola vez antes de la primera conciliación.
      <div class="actions" style="margin-top:8px"><label class="btn" for="f-seed">Importar estado inicial (.json)</label><input id="f-seed" type="file" accept=".json" hidden></div></div>`;
  }

  /* ---------- métricas para Dirección ---------- */
  const sumBy = (arr, f) => arr.reduce((a, x) => a + (typeof f === 'function' ? f(x) : x[f]) || 0, 0);
  const CAJAS = { CCO: 'Caja Combustibles', CSL: 'Caja San Lorenzo', CPQ: 'Caja Petroquímicos', LP: 'LP', TRB: 'Transferencias' };
  function caja(l) {
    const r = String(l.ref || '').toUpperCase();
    if (l.tipo === 'certribr' || /IIBB/.test(r)) return 'Retenciones IIBB';
    const m = r.match(/^(CCO|CSL|CPQ|LP|TRB)\b/);
    if (m) return CAJAS[m[1]];
    if (/DEPOSITO|DEP\.? PICO|PICO/.test(r)) return 'Depósitos de estaciones';
    return l.u === 'CO' ? 'Otros pagos CO' : 'Otros pagos PQ';
  }
  const fechaCorr = (c) => (c.hoy || (c.fecha ? new Date(c.fecha).toLocaleDateString('sv-SE', { timeZone: 'America/Argentina/Buenos_Aires' }) : ''));

  function metricas(v) {
    const hoy = hoyISO();
    const deudaTodas = v.filas.filter((f) => !f.recibo && f.orig > 0);
    const deuda = deudaTodas.filter((f) => !f.debitado);
    const uOf = (f) => f.u || E.unidadDeDiv(f.div) || '—';
    const sumU = (rows, k) => ({ rows, tot: sumBy(rows, k), co: sumBy(rows.filter((f) => uOf(f) === 'CO'), k), pq: sumBy(rows.filter((f) => uOf(f) === 'PQ'), k), sa: sumBy(rows.filter((f) => uOf(f) === '—'), k) });
    // Facturas con débito automático informado: no cuentan como vencidas hasta que YPF las baje
    const deb = { g: 'deb', l: 'Débito automático informado', ...sumU(deudaTodas.filter((f) => f.debitado), 'orig') };
    // Acreditado por YPF y todavía no compensado contra facturas (sin retenciones duplicadas)
    const acred = sumU(v.filas.filter((f) => f.recibo && f.orig < 0 && !['retdup', 'posibleDup'].includes(f.estado)), 'orig');
    const B = [['venc', 'Vencido', (d) => d < 0, 'venc'], ['hoy', 'Vence hoy', (d) => d === 0, 'hoy'], ['a7', 'A vencer · próximos 7 días', (d) => d >= 1 && d <= 7, 'aven'], ['a15', 'A vencer · 8 a 15 días', (d) => d >= 8 && d <= 15, 'aven'], ['amas', 'A vencer · más de 15 días', (d) => d > 15, 'aven']];
    const venc = B.map(([k, l, fn, g]) => { const rows = deuda.filter((f) => fn(E.daysBetween(hoy, f.fvenc))); return { k, l, g, rows, tot: sumBy(rows, 'orig'), co: sumBy(rows.filter((f) => uOf(f) === 'CO'), 'orig'), pq: sumBy(rows.filter((f) => uOf(f) === 'PQ'), 'orig') }; });
    const grupo = (g, l) => { const bs = venc.filter((b) => b.g === g); const rows = bs.flatMap((b) => b.rows); return { g, l, rows, tot: sumBy(bs, 'tot'), co: sumBy(bs, 'co'), pq: sumBy(bs, 'pq') }; };
    const vencG = [grupo('venc', 'Vencido'), grupo('hoy', 'Vence hoy'), grupo('aven', 'A vencer')];
    const exig = { co: vencG[0].co + vencG[1].co, pq: vencG[0].pq + vencG[1].pq, tot: vencG[0].tot + vencG[1].tot };
    const neto = { co: exig.co + acred.co, pq: exig.pq + acred.pq, tot: exig.tot + acred.tot };
    // Cheques rechazados: circuito hasta cerrarlos
    const chrs = Object.entries(S.lineas).filter(([, l]) => l.est === 'CHR').map(([k, l]) => ({ ...l, key: k, chrE: CHRE[l.chrE] ? l.chrE : l.chrE === 'nd' ? 'recup' : 'rech', dias: E.daysBetween(l.okF || l.f, hoy) }))
      .sort((a, b) => (a.f < b.f ? -1 : a.f > b.f ? 1 : 0));
    const chrAb = chrs.filter((c) => c.chrE !== 'cerr');
    const creditos = sumBy(v.filas.filter((f) => f.recibo), 'orig');
    const usdRows = v.filas.filter((f) => f.mon === 'USD');
    const usd = { rows: usdRows, usd: sumBy(usdRows, 'usd'), ars: sumBy(usdRows, 'ars'), orig: sumBy(usdRows, 'orig') };
    usd.tc = usd.usd ? usd.ars / usd.usd : 0;
    const PB = [['0 a 3 días', 0, 3], ['4 a 7 días', 4, 7], ['8 a 30 días', 8, 30], ['Más de 30 días', 31, 1e9]];
    const paBuckets = PB.map(([l, a, b]) => { const ps = v.pendientes.filter((p) => p.dias >= a && p.dias <= b); return { l, co: sumBy(ps.filter((p) => p.u === 'CO'), 'imp'), pq: sumBy(ps.filter((p) => p.u === 'PQ'), 'imp'), n: ps.length }; });
    const masViejo = v.pendientes.reduce((a, p) => Math.max(a, p.dias || 0), 0);
    // deuda según YPF por unidad: facturas por división; recibos por el pago con el que cruzaron
    const porU = { CO: 0, PQ: 0, '—': 0 };
    for (const f of v.filas) { const u = uOf(f); porU[u] = (porU[u] || 0) + f.orig; }
    const mes = hoy.slice(0, 7);
    const cajasMap = {};
    for (const l of Object.values(S.lineas)) {
      if (!l.f || l.f.slice(0, 7) !== mes || l.est === 'R' || l.est === 'CHR') continue;
      const c = caja(l); const o = cajasMap[c] = cajasMap[c] || { caja: c, n: 0, tot: 0, pend: 0 };
      o.n++; o.tot += l.imp; if (l.est === 'PA') o.pend += l.imp;
    }
    const cajas = Object.values(cajasMap).sort((a, b) => b.tot - a.tot);
    const evol = S.corridas.slice(0, 30).filter((c) => c.resumen).slice().reverse().map((c) => ({ f: fechaCorr(c), fecha: c.fecha, ypf: c.resumen.deudaYPF, real: c.resumen.deudaReal }));
    const reclamos = { retdup: v.retdup, posibleDup: v.posibleDup, sin: v.sin, parciales: v.parciales };
    reclamos.tot = sumBy(v.retdup, (f) => f.rec.imp) + sumBy(v.sin, (f) => f.rec.imp);
    return { hoy, deuda, venc, vencG, uOf, creditos, usd, paBuckets, masViejo, porU, cajas, evol, reclamos, mes, deb, acred, exig, neto, chrs, chrAb };
  }

  /* ---------- gráficos (sin librerías) ---------- */
  function hbars(items, fmtv) {
    const max = Math.max(1, ...items.map((i) => Math.abs(i.value)));
    return `<div class="hbars" role="list">${items.map((i) => `<div class="hb" role="listitem" title="${esc(i.label)}: ${esc(fmtv(i.value))}${i.extra ? ' · ' + esc(i.extra) : ''}">
      <span class="hb-l">${esc(i.label)}</span><span class="hb-t"><span class="hb-b ${i.cls || ''}" style="width:${(Math.abs(i.value) / max * 100).toFixed(2)}%;${i.value ? '' : 'min-width:0'}"></span></span><span class="hb-v num">${esc(fmtv(i.value))}</span></div>`).join('')}</div>`;
  }
  function lineChart(pts, series) {
    if (pts.length < 2) return `<p class="muted small">El gráfico aparece a partir de la segunda conciliación.</p>`;
    const W = 1000, H = 300, L = 100, R = 130, T = 16, Bm = 34;
    const vals = pts.flatMap((p) => series.map((s) => p[s.k]));
    let lo = Math.min(...vals), hi = Math.max(...vals); const pad = (hi - lo || hi * 0.1 || 1) * 0.12; lo -= pad; hi += pad;
    const x = (i) => L + (i * (W - L - R)) / (pts.length - 1);
    const y = (val) => T + (1 - (val - lo) / (hi - lo)) * (H - T - Bm);
    const ticks = [0, 1, 2, 3].map((i) => lo + ((hi - lo) * i) / 3);
    const step = Math.max(1, Math.ceil(pts.length / 7));
    let svg = `<svg class="lc" viewBox="0 0 ${W} ${H}" role="img" aria-label="Evolución de la deuda según YPF y de la deuda real">`;
    for (const t of ticks) svg += `<line x1="${L}" x2="${W - R}" y1="${y(t)}" y2="${y(t)}" class="grid"/><text x="${L - 8}" y="${y(t) + 4}" text-anchor="end" class="axis">${esc(mill(t).replace('$ ', ''))}</text>`;
    pts.forEach((p, i) => { if (i % step === 0 || i === pts.length - 1) svg += `<text x="${x(i)}" y="${H - 8}" text-anchor="middle" class="axis">${esc(dmy(p.f).slice(0, 5))}</text>`; });
    const ends = [];
    for (const s of series) {
      svg += `<polyline fill="none" class="ln ${s.cls}" points="${pts.map((p, i) => `${x(i).toFixed(1)},${y(p[s.k]).toFixed(1)}`).join(' ')}"/>`;
      pts.forEach((p, i) => { svg += `<circle cx="${x(i)}" cy="${y(p[s.k])}" r="4.5" class="dot ${s.cls}"><title>${esc(s.name)} · ${esc(fechaHora(p.fecha))}: ${esc(money(p[s.k]))}</title></circle>`; });
      ends.push({ s, yy: y(pts[pts.length - 1][s.k]) });
    }
    ends.sort((a, b) => a.yy - b.yy); for (let i = 1; i < ends.length; i++) if (ends[i].yy - ends[i - 1].yy < 16) ends[i].yy = ends[i - 1].yy + 16;
    for (const e of ends) svg += `<text x="${W - R + 10}" y="${e.yy + 4}" class="endlbl">${esc(e.s.name)}</text>`;
    return svg + `</svg>`;
  }

  /* ---------- vista general (KPIs + selector) ---------- */
  function vistaResultado(v, corr) {
    const m = metricas(v); S._m = m;
    return `
      <section class="panel eq" aria-label="Resultado">
        <div class="cell"><span class="lbl">Deuda según YPF</span><span class="big">${esc(mill(v.deudaYPF))}</span><span class="muted small num">${esc(money(v.deudaYPF))}</span>
          <span class="split"><span>Combustibles <b class="num">${esc(mill(m.porU.CO))}</b></span><span>Petroquímicos <b class="num">${esc(mill(m.porU.PQ))}</b></span></span></div>
        <div class="op">−</div>
        <div class="cell"><span class="lbl"><abbr title="Pagos registrados en Nodum que YPF todavía no acreditó en su cuenta corriente">Pendiente de acreditar</abbr></span><span class="big">${esc(mill(v.paCO + v.paPQ))}</span>
          <span class="split"><span>Combustibles <b class="num">${esc(mill(v.paCO))}</b></span><span>Petroquímicos <b class="num">${esc(mill(v.paPQ))}</b></span></span></div>
        <div class="op">=</div>
        <div class="cell real"><span class="lbl">Deuda real</span><span class="big">${esc(mill(v.deudaReal))}</span><span class="small num">${esc(money(v.deudaReal))}</span>
          <span class="split"><span>Combustibles <b class="num">${esc(mill(m.porU.CO - v.paCO))}</b></span><span>Petroquímicos <b class="num">${esc(mill(m.porU.PQ - v.paPQ))}</b></span></span>
          ${v.retdupTotal ? `<span class="small">+ ${esc(money(v.retdupTotal))} al anularse las retenciones duplicadas</span>` : ''}</div>
      </section>
      ${S.vista === 'dir' ? vistaDir(v, m, corr) : vistaConc(v, m, corr)}
      ${hojaResumen(v, m, corr)}`;
  }

  /* ---------- Dirección ---------- */
  function vistaDir(v, m, corr) {
    const ult = m.evol.length > 1 ? m.evol[m.evol.length - 1].real - m.evol[m.evol.length - 2].real : null;
    const [gV, gH, gA] = m.vencG;
    const nNod = v.noCargadas.length;
    const cards = [
      ['real', 'Deuda real', mill(v.deudaReal), `Combustibles ${mill(m.porU.CO - v.paCO)} · Petroquímicos ${mill(m.porU.PQ - v.paPQ)}`, ''],
      ['venc', 'Vencimientos', mill(gV.tot + gH.tot), `exigible hoy · neto de lo acreditado ${mill(Math.max(0, m.neto.tot))} · a vencer ${mill(gA.tot)}`, gV.tot ? 'bad' : gH.tot ? 'warn' : ''],
      ['usd', 'Exposición en dólares', 'USD ' + nf0.format(m.usd.usd), m.usd.usd ? `TC implícito ${nf.format(m.usd.tc)}` : 'Sin facturas en USD', ''],
      ['pa', 'Pendiente de acreditar', mill(v.paCO + v.paPQ), `${v.pendientes.length} valores · el más antiguo, ${m.masViejo} días`, m.masViejo > 7 ? 'warn' : ''],
      ['recl', 'Reclamos a YPF', String(v.retdup.length + v.posibleDup.length + v.sin.length), `partidas · ${mill(m.reclamos.tot)}`, v.retdup.length + v.sin.length ? 'warn' : ''],
      ['nodum', 'Pendiente de registro en Nodum', String(nNod), nNod ? `comprobantes · ${mill(sumBy(v.noCargadas, 'orig'))}` : 'comprobantes', nNod ? 'warn' : ''],
      ['evol', 'Evolución de la deuda', ult == null ? '—' : (ult >= 0 ? '+ ' : '− ') + mill(Math.abs(ult)).replace('$ ', '$ '), 'deuda real vs. conciliación anterior', ''],
      ['cajas', 'Pagos del mes por caja', mill(sumBy(m.cajas, 'tot')), `${m.cajas.length} cajas y conceptos`, ''],
      ['chr', 'Cheques rechazados', String(m.chrAb.length), m.chrAb.length ? `abiertos · ${mill(sumBy(m.chrAb, 'imp'))}` : 'sin cheques abiertos', m.chrAb.length ? 'warn' : ''],
    ];
    return `
      <div class="sec-head"><h2>Informes</h2><div class="actions"><button data-a="copiarRes">Copiar resumen</button><button class="primary" data-a="resumen">Resumen ejecutivo (PDF)</button></div></div>
      <div class="cards" role="tablist" aria-label="Informes">${cards.map((c) => `<button class="rcard ${c[4]}" role="tab" aria-selected="${S.reporte === c[0]}" data-a="rep" data-r="${c[0]}">
        <span class="lbl">${esc(c[1])}</span><span class="rnum num">${esc(c[2])}</span><span class="small muted">${esc(c[3])}</span></button>`).join('')}</div>
      <section class="panel card" id="detalle">${detalle(v, m)}</section>`;
  }

  function tablaAcred(m) {
    if (!m.acred.rows.length) return `<p class="muted small">No hay recibos ni saldos a favor sin compensar.</p>`;
    const rows = m.acred.rows.slice().sort((a, b) => (m.uOf(a) + a.fdoc > m.uOf(b) + b.fdoc ? 1 : -1));
    return `<div class="tw"><table><thead><tr><th>Unidad</th><th>Recibo</th><th>T. Pos.</th><th>Div.</th><th>F. Doc.</th><th>Pago en Nodum</th><th class="n">Importe</th></tr></thead><tbody>
      ${rows.map((f) => `<tr><td>${esc(uNombre(m.uOf(f)))}${f.uManual ? ' <span class="muted small">(a mano)</span>' : ''}</td><td class="num">${esc(f.nro)}</td><td class="small">${esc(f.tpos)}</td><td>${esc(f.div)}</td><td>${dmy(f.fdoc)}</td><td class="small">${f.pago ? `trans ${f.pago.trans}${f.pago.ant ? ' · ANT ' + f.pago.ant : ''}${f.pago.ref ? ' · ' + esc(f.pago.ref) : ''}` : f.rec && f.rec.etq ? esc(ETQ[f.rec.etq] || f.rec.etq) : '—'}</td><td class="n">${esc(nf.format(f.orig))}</td></tr>`).join('')}
      <tr class="tot"><td colspan="6">Total · Combustibles ${esc(money(m.acred.co))} · Petroquímicos ${esc(money(m.acred.pq))}</td><td class="n">${esc(nf.format(m.acred.tot))}</td></tr></tbody></table></div>`;
  }
  function detalle(v, m) {
    const r = S.reporte;
    if (r === 'real') {
      const fila = (u, n) => `<tr><td>${n}</td><td class="n">${esc(money(m.porU[u] || 0))}</td><td class="n">${esc(money(u === 'CO' ? -v.paCO : u === 'PQ' ? -v.paPQ : 0))}</td><td class="n"><b>${esc(money((m.porU[u] || 0) - (u === 'CO' ? v.paCO : u === 'PQ' ? v.paPQ : 0)))}</b></td></tr>`;
      return `<h3>Deuda real por unidad de negocio</h3>
        <div class="tw"><table><thead><tr><th>Unidad</th><th class="n">Deuda según YPF</th><th class="n">Pendiente de acreditar</th><th class="n">Deuda real</th></tr></thead><tbody>
        ${fila('CO', 'Combustibles (3000266)')}${fila('PQ', 'Petroquímicos (0001003)')}${Math.abs(m.porU['—'] || 0) > 0.5 ? fila('—', 'Sin asignar') : ''}
        <tr class="tot"><td>Total</td><td class="n">${esc(money(v.deudaYPF))}</td><td class="n">${esc(money(-(v.paCO + v.paPQ)))}</td><td class="n">${esc(money(v.deudaReal))}</td></tr></tbody></table></div>
        <h3>Deuda según YPF por división</h3>
        <div class="tw"><table><thead><tr><th>División</th><th class="n">Débitos</th><th class="n">Créditos</th><th class="n">Neto</th><th class="n">Partidas</th></tr></thead><tbody>
        ${v.divs.map((d) => `<tr><td>${esc(DIVS[d.div] || d.div)}</td><td class="n">${esc(money(d.debitos))}</td><td class="n">${esc(money(d.creditos))}</td><td class="n">${esc(money(d.debitos + d.creditos))}</td><td class="n">${d.n}</td></tr>`).join('')}
        </tbody></table></div>
        <p class="muted small">Unidad: las facturas se asignan por división (MN, MA, GR → Combustibles; PQ, ME → Petroquímicos); los recibos, por el pago de Nodum con el que se cruzaron. Cualquier partida puede reasignarse a mano desde la pestaña Extranet.</p>`;
    }
    if (r === 'venc') {
      const uN = { CO: 'Combustibles', PQ: 'Petroquímicos', '—': '—' };
      const [gV, gH, gA] = m.vencG;
      const fila = (b, cls) => `<tr class="${cls || ''}"><td>${esc(b.l)}</td><td class="n">${b.rows.length}</td><td class="n">${esc(money(b.co))}</td><td class="n">${esc(money(b.pq))}</td><td class="n">${esc(money(b.tot))}</td></tr>`;
      const tabla = (g, modo) => {
        if (!g.rows.length) return `<p class="muted small">Sin comprobantes.</p>`;
        const rows = g.rows.slice().sort((a, b) => (a.fvenc < b.fvenc ? -1 : a.fvenc > b.fvenc ? 1 : 0));
        return `<div class="tw"><table><thead><tr><th>Vence</th><th class="n">${modo === 'venc' ? 'Días de atraso' : modo === 'aven' ? 'Faltan (días)' : modo === 'deb' ? 'Días desde el vto.' : 'Días'}</th><th>Comprobante</th><th>Unidad</th><th>División</th><th>Moneda</th><th class="n">USD</th><th class="n">Importe ARS</th></tr></thead><tbody>
          ${rows.map((f) => { const d = E.daysBetween(m.hoy, f.fvenc); return `<tr><td>${dmy(f.fvenc)}</td><td class="n"><span class="age ${d < 0 ? 'old' : ''}">${Math.abs(d)}</span></td><td>${esc(f.tipo)} <span class="num">${esc(f.nro)}</span></td><td>${esc(uN[m.uOf(f)])}</td><td>${esc(DIVS[f.div] || f.div)}</td><td>${esc(f.mon)}</td><td class="n">${f.usd ? esc(nf.format(f.usd)) : ''}</td><td class="n">${esc(nf.format(f.orig))}</td></tr>`; }).join('')}
          <tr class="tot"><td colspan="7">Total ${esc(g.l.toLowerCase())} · Combustibles ${esc(money(g.co))} · Petroquímicos ${esc(money(g.pq))}</td><td class="n">${esc(nf.format(g.tot))}</td></tr></tbody></table></div>`;
      };
      return `<h3>Vencimientos de facturas y notas de débito</h3>
        <div class="tw"><table><thead><tr><th>Estado</th><th class="n">Comprobantes</th><th class="n">Combustibles</th><th class="n">Petroquímicos</th><th class="n">Total</th></tr></thead><tbody>
          ${fila(gV, gV.tot ? 'row-bad' : '')}${fila(gH, gH.tot ? 'row-warn' : '')}${fila(gA, 'grp')}
          ${m.venc.filter((b) => b.g === 'aven').map((b) => fila({ ...b, l: '   ' + b.l.replace('A vencer · ', '') }, 'sub')).join('')}
          <tr class="tot"><td>Total</td><td class="n">${m.deuda.length}</td><td class="n">${esc(money(gV.co + gH.co + gA.co))}</td><td class="n">${esc(money(gV.pq + gH.pq + gA.pq))}</td><td class="n">${esc(money(gV.tot + gH.tot + gA.tot))}</td></tr>
        </tbody></table></div>
        <h3>Exigible neto de lo acreditado</h3>
        <div class="tw"><table><thead><tr><th></th><th class="n">Combustibles</th><th class="n">Petroquímicos</th><th class="n">Total</th></tr></thead><tbody>
          <tr><td>Exigible hoy (vencido + vence hoy)</td><td class="n">${esc(money(m.exig.co))}</td><td class="n">${esc(money(m.exig.pq))}</td><td class="n">${esc(money(m.exig.tot))}</td></tr>
          <tr><td>Acreditado por YPF sin compensar (recibos y saldos a favor)</td><td class="n">${esc(money(m.acred.co))}</td><td class="n">${esc(money(m.acred.pq))}</td><td class="n">${esc(money(m.acred.tot))}</td></tr>
          <tr class="tot"><td>Exigible neto</td><td class="n">${esc(money(m.neto.co))}</td><td class="n">${esc(money(m.neto.pq))}</td><td class="n">${esc(money(m.neto.tot))}</td></tr>
        </tbody></table></div>
        <p class="muted small">Lo acreditado y todavía no compensado se descuenta de lo exigible. No incluye las retenciones duplicadas.${m.neto.tot < 0 ? ' Un neto negativo significa que lo acreditado alcanza para cubrir todo lo exigible hoy y sobra.' : ''}${m.acred.sa ? ` Hay ${esc(money(m.acred.sa))} sin unidad asignada.` : ''}</p>
        <details><summary class="small">Ver lo acreditado sin compensar, por unidad (${m.acred.rows.length})</summary>${tablaAcred(m)}</details>
        ${m.deb.rows.length ? `<p class="small">Además hay <b>${esc(money(m.deb.tot))}</b> en ${m.deb.rows.length} facturas con débito automático informado, que no se cuentan como vencidas hasta que YPF las baje.</p>` : ''}
        <h3 class="vh bad">Vencido <span class="num">${esc(money(gV.tot))}</span></h3>${tabla(gV, 'venc')}
        <h3 class="vh warn">Vence hoy <span class="num">${esc(money(gH.tot))}</span></h3>${tabla(gH, 'hoy')}
        <h3 class="vh">A vencer <span class="num">${esc(money(gA.tot))}</span></h3>${tabla(gA, 'aven')}
        ${m.deb.rows.length ? `<h3 class="vh">Débito automático informado <span class="num">${esc(money(m.deb.tot))}</span></h3>${tabla(m.deb, 'deb')}` : ''}`;
    }
    if (r === 'chr') {
      const porE = Object.keys(CHRE).map((k) => { const cs = m.chrs.filter((c) => c.chrE === k); return { k, l: CHRE[k], n: cs.length, tot: sumBy(cs, 'imp') }; });
      return `<h3>Cheques rechazados</h3>
        <div class="tw"><table><thead><tr><th>Estado</th><th class="n">Cheques</th><th class="n">Importe</th></tr></thead><tbody>
        ${porE.map((e) => `<tr><td>${esc(e.l)}</td><td class="n">${e.n}</td><td class="n">${esc(money(e.tot))}</td></tr>`).join('')}
        <tr class="tot"><td>Abiertos (sin cerrar)</td><td class="n">${m.chrAb.length}</td><td class="n">${esc(money(sumBy(m.chrAb, 'imp')))}</td></tr></tbody></table></div>
        ${m.chrAb.length ? `<div class="tw"><table><thead><tr><th>Unidad</th><th>Fecha</th><th>Referencia</th><th>Cheque</th><th class="n">Importe</th><th>Estado</th><th class="n">Días</th><th>Justificación</th></tr></thead><tbody>
        ${m.chrAb.map((c) => `<tr><td>${esc(uNombre(c.u))}</td><td>${dmy(c.f)}</td><td>${esc(c.ref || '')}</td><td>${esc(medio(c.tipo))} <span class="num">${c.nro || ''}</span></td><td class="n">${esc(nf.format(c.imp))}</td><td>${esc(CHRE[c.chrE])}</td><td class="n">${c.dias == null ? '' : c.dias}</td><td class="small">${esc(c.chrNota || '')}</td></tr>`).join('')}</tbody></table></div>` : ''}
        <div class="actions"><button data-a="irConc" data-t="chr">Gestionar los cheques</button></div>`;
    }
    if (r === 'usd') {
      const u = m.usd;
      if (!u.rows.length) return `<h3>Exposición en dólares</h3><p class="muted">No hay comprobantes en dólares pendientes.</p>`;
      const sens = [-10, -5, 5, 10].map((p) => ({ p, ars: u.usd * u.tc * (1 + p / 100) }));
      return `<h3>Exposición en dólares</h3>
        <div class="kv kv-wide">
          <span>Pendiente en dólares</span><span class="n">USD ${esc(nf.format(u.usd))}</span>
          <span>Tipo de cambio implícito (Extranet)</span><span class="n">${esc(nf.format(u.tc))}</span>
          <span>Equivalente en pesos al tipo de cambio actual</span><span class="n">${esc(money(u.ars))}</span>
          <span>Importe original en pesos</span><span class="n">${esc(money(u.orig))}</span>
          <span>Diferencia de cambio</span><span class="n">${esc(money(u.ars - u.orig))}</span>
        </div>
        <h3>Sensibilidad al tipo de cambio</h3>
        <div class="tw"><table><thead><tr><th>Variación del dólar</th><th class="n">Tipo de cambio</th><th class="n">Deuda en dólares, en pesos</th><th class="n">Diferencia</th></tr></thead><tbody>
        ${sens.map((s) => `<tr><td>${s.p > 0 ? '+' : ''}${s.p} %</td><td class="n">${esc(nf.format(u.tc * (1 + s.p / 100)))}</td><td class="n">${esc(money(s.ars))}</td><td class="n">${esc(money(s.ars - u.ars))}</td></tr>`).join('')}
        </tbody></table></div>
        <div class="tw"><table><thead><tr><th>Comprobante</th><th>División</th><th>Vence</th><th class="n">USD</th><th class="n">ARS al TC actual</th></tr></thead><tbody>
        ${u.rows.map((f) => `<tr><td>${esc(f.tipo)} <span class="num">${esc(f.nro)}</span></td><td>${esc(DIVS[f.div] || f.div)}</td><td>${dmy(f.fvenc)}</td><td class="n">${esc(nf.format(f.usd))}</td><td class="n">${esc(nf.format(f.ars))}</td></tr>`).join('')}
        </tbody></table></div>`;
    }
    if (r === 'pa') {
      return `<h3>Pendiente de acreditar por antigüedad</h3>
        ${hbars(m.paBuckets.map((b) => ({ label: b.l, value: b.co + b.pq, cls: /30/.test(b.l) && b.co + b.pq ? 'bad' : '', extra: b.n + ' valores' })), money)}
        <div class="tw"><table><thead><tr><th>Antigüedad</th><th class="n">Combustibles</th><th class="n">Petroquímicos</th><th class="n">Total</th><th class="n">Valores</th></tr></thead><tbody>
        ${m.paBuckets.map((b) => `<tr><td>${b.l}</td><td class="n">${esc(money(b.co))}</td><td class="n">${esc(money(b.pq))}</td><td class="n">${esc(money(b.co + b.pq))}</td><td class="n">${b.n}</td></tr>`).join('')}
        <tr class="tot"><td>Total</td><td class="n">${esc(money(v.paCO))}</td><td class="n">${esc(money(v.paPQ))}</td><td class="n">${esc(money(v.paCO + v.paPQ))}</td><td class="n">${v.pendientes.length}</td></tr></tbody></table></div>
        <div class="actions"><button data-a="irConc" data-t="pend">Ver el detalle de cada valor</button></div>`;
    }
    if (r === 'recl') {
      const tbl = (titulo, arr, extra) => arr.length ? `<h3>${titulo} <span class="muted small">(${arr.length})</span></h3><div class="tw"><table><thead><tr><th>Recibo</th><th>Concepto</th><th>Fecha</th><th class="n">Días</th><th class="n">Importe</th></tr></thead><tbody>
        ${arr.map((f) => `<tr><td class="num">${esc(f.nro)}</td><td>${esc(f.tpos)}${extra ? ' · ' + extra(f) : ''}</td><td>${dmy(f.fdoc)}</td><td class="n">${E.daysBetween(f.fdoc, m.hoy)}</td><td class="n">${esc(nf.format(f.rec.imp))}</td></tr>`).join('')}</tbody></table></div>` : '';
      const html = tbl('Retenciones duplicadas: solicitar anulación', v.retdup) + tbl('Posibles retenciones duplicadas', v.posibleDup) +
        tbl('Créditos sin identificar', v.sin) + tbl('Saldos parciales a favor', v.parciales, (f) => `pago de ${esc(money(f.pago.imp))}`);
      return html || `<h3>Reclamos a YPF</h3><p class="muted">No hay partidas para reclamar.</p>`;
    }
    if (r === 'nodum') {
      if (!v.noCargadas.length) return `<h3>Pendiente de registro en Nodum</h3><p class="muted">Todos los comprobantes del período están registrados${(S._corr.navs || []).length ? '' : ' (cargá el navegador contable para controlarlo)'}.</p>`;
      return `<h3>Comprobantes de YPF sin registrar en Nodum</h3><p class="muted small">Facturas MN y notas de débito en pesos del período del navegador contable que no figuran en la cuenta 21110101.</p>
        <div class="tw"><table><thead><tr><th>Comprobante</th><th>División</th><th>Fecha</th><th>Vence</th><th class="n">Importe</th></tr></thead><tbody>
        ${v.noCargadas.map((e) => `<tr><td>${esc(e.tipo)} <span class="num">${esc(e.nro)}</span></td><td>${esc(DIVS[e.div] || e.div)}</td><td>${dmy(e.fdoc)}</td><td>${dmy(e.fvenc)}</td><td class="n">${esc(nf.format(e.orig))}</td></tr>`).join('')}
        <tr class="tot"><td colspan="4">Total</td><td class="n">${esc(nf.format(sumBy(v.noCargadas, 'orig')))}</td></tr></tbody></table></div>`;
    }
    if (r === 'evol') {
      return `<h3>Evolución de la deuda</h3>
        <div class="legend"><span><i class="sw s1"></i>Deuda según YPF</span><span><i class="sw s2"></i>Deuda real</span></div>
        <div class="chartbox">${lineChart(m.evol, [{ k: 'ypf', name: 'Según YPF', cls: 's1' }, { k: 'real', name: 'Deuda real', cls: 's2' }])}</div>
        <div class="tw"><table><thead><tr><th>Conciliación</th><th class="n">Deuda según YPF</th><th class="n">Deuda real</th></tr></thead><tbody>
        ${m.evol.slice().reverse().map((p) => `<tr><td>${esc(fechaHora(p.fecha))}</td><td class="n">${esc(money(p.ypf))}</td><td class="n">${esc(money(p.real))}</td></tr>`).join('')}</tbody></table></div>`;
    }
    if (r === 'cajas') {
      const mesTxt = new Date(m.hoy + 'T12:00:00').toLocaleDateString('es-AR', { month: 'long', year: 'numeric' });
      return `<h3>Pagos a YPF de ${esc(mesTxt)}, por caja y concepto</h3>
        ${hbars(m.cajas.map((c) => ({ label: c.caja, value: c.tot, extra: c.n + ' valores' })), money)}
        <div class="tw"><table><thead><tr><th>Caja / concepto</th><th class="n">Valores</th><th class="n">Pagado</th><th class="n">Pendiente de acreditar</th></tr></thead><tbody>
        ${m.cajas.map((c) => `<tr><td>${esc(c.caja)}</td><td class="n">${c.n}</td><td class="n">${esc(money(c.tot))}</td><td class="n">${esc(money(c.pend))}</td></tr>`).join('')}
        <tr class="tot"><td>Total</td><td class="n">${sumBy(m.cajas, 'n')}</td><td class="n">${esc(money(sumBy(m.cajas, 'tot')))}</td><td class="n">${esc(money(sumBy(m.cajas, 'pend')))}</td></tr></tbody></table></div>
        <p class="muted small">La caja sale de la referencia del pago (CCO, CSL, CPQ…). Los pagos sin referencia se agrupan como “Otros pagos”.</p>`;
    }
    return '';
  }

  /* ---------- Conciliación (operativa) ---------- */
  function vistaConc(v, m, corr) {
    const hoyOk = corr && fechaCorr(corr) === m.hoy;
    const tareas = [
      ['Cargar los archivos del día', hoyOk, hoyOk ? 'Hecho' : 'Pendiente', 'carga'],
      ['Identificar créditos sin identificar', !v.sin.length, v.sin.length + ' recibos', 'exc', 'x-sin'],
      ['Revisar posibles retenciones duplicadas', !v.posibleDup.length, v.posibleDup.length + ' recibos', 'exc', 'x-dup'],
      ['Controlar pendientes de acreditar con más de 7 días', !v.viejos.length, v.viejos.length + ' valores', 'pend', 'x-viejo'],
      ['Solicitar a YPF la anulación de retenciones duplicadas', !v.retdup.length, v.retdup.length + ' retenciones', 'exc', 'x-retdup'],
      ['Informar comprobantes sin registrar en Nodum', !v.noCargadas.length, v.noCargadas.length + ' comprobantes', 'exc', 'x-nonodum'],
    ];
    const hechas = tareas.filter((t) => t[1]).length;
    const tabs = [['pend', 'Pendientes de acreditar', v.pendientes.length], ['ext', 'Extranet', v.filas.length], ['exc', 'Excepciones', v.sin.length + v.posibleDup.length + v.retdup.length + v.noCargadas.length + v.alertas.length + v.parciales.length], ['mov', 'Movimientos', v.mov ? v.mov.facturasSalieron.length + v.mov.recibosSalieron.length + v.mov.facturasNuevas.length + v.mov.recibosNuevos.length : 0], ['chr', 'Cheques rechazados', m.chrAb.length], ['hist', 'Historial', S.corridas.length]];
    return `
      <section class="panel card" aria-label="Tareas del día">
        <div class="sec-head"><h2>Tareas del día</h2><span class="small"><b>${hechas}</b> de ${tareas.length} completas</span></div>
        <div class="prog" role="progressbar" aria-valuemin="0" aria-valuemax="${tareas.length}" aria-valuenow="${hechas}"><span style="width:${(hechas / tareas.length * 100).toFixed(1)}%"></span></div>
        <ol class="tasks">${tareas.map((t) => `<li class="${t[1] ? 'done' : ''}"><span class="chk" aria-hidden="true">${t[1] ? '✓' : ''}</span><span class="tt">${esc(t[0])}</span><span class="small ${t[1] ? 'muted' : ''}">${esc(t[2])}</span>
          ${t[1] ? '' : t[3] === 'carga' ? `<button class="mini" data-a="toggleCarga">Cargar</button>` : `<button class="mini" data-a="tab" data-t="${t[3]}" data-go="${t[4] || ''}">Ver</button>`}</li>`).join('')}</ol>
      </section>
      <div class="tabs" role="tablist">${tabs.map((t) => `<button class="tab" role="tab" id="tab-${t[0]}" aria-selected="${S.tab === t[0]}" data-a="tab" data-t="${t[0]}">${t[1]}<span class="cnt">${t[2]}</span></button>`).join('')}</div>
      <section class="sec">${S.tab === 'pend' ? tabPend(v) : S.tab === 'ext' ? tabExt(v) : S.tab === 'exc' ? tabExc(v) : S.tab === 'mov' ? tabMov(v) : S.tab === 'chr' ? tabChr(m) : tabHist()}</section>`;
  }

  /* ---------- resumen ejecutivo ---------- */
  function hojaResumen(v, m, corr) {
    return `<section class="print-sheet" aria-hidden="true">
      <div class="brandbar"><img class="logo" src="logo-copsa.jpg" alt="" width="153" height="44"><span class="sep"></span><div class="who"><span class="lbl">Resumen ejecutivo</span><h1>Cuenta corriente con YPF</h1></div></div>
      <p class="small">Conciliación del ${esc(fechaHora(corr.fecha))}${corr.usuario ? ' · ' + esc(corr.usuario) : ''}</p>
      <table><thead><tr><th></th><th class="n">Combustibles</th><th class="n">Petroquímicos</th><th class="n">Total</th></tr></thead><tbody>
        <tr><td>Deuda según YPF</td><td class="n">${esc(money(m.porU.CO))}</td><td class="n">${esc(money(m.porU.PQ))}</td><td class="n">${esc(money(v.deudaYPF))}</td></tr>
        <tr><td>Pendiente de acreditar</td><td class="n">${esc(money(-v.paCO))}</td><td class="n">${esc(money(-v.paPQ))}</td><td class="n">${esc(money(-(v.paCO + v.paPQ)))}</td></tr>
        <tr class="tot"><td>Deuda real</td><td class="n">${esc(money(m.porU.CO - v.paCO))}</td><td class="n">${esc(money(m.porU.PQ - v.paPQ))}</td><td class="n">${esc(money(v.deudaReal))}</td></tr>
      </tbody></table>
      <h3>Vencimientos</h3><table><thead><tr><th></th><th class="n">Combustibles</th><th class="n">Petroquímicos</th><th class="n">Total</th></tr></thead><tbody>${m.vencG.map((b) => `<tr><td>${b.l}</td><td class="n">${esc(money(b.co))}</td><td class="n">${esc(money(b.pq))}</td><td class="n">${esc(money(b.tot))}</td></tr>`).join('')}
        <tr><td>Acreditado sin compensar</td><td class="n">${esc(money(m.acred.co))}</td><td class="n">${esc(money(m.acred.pq))}</td><td class="n">${esc(money(m.acred.tot))}</td></tr>
        <tr class="tot"><td>Exigible neto (vencido + vence hoy − acreditado)</td><td class="n">${esc(money(m.neto.co))}</td><td class="n">${esc(money(m.neto.pq))}</td><td class="n">${esc(money(m.neto.tot))}</td></tr></tbody></table>
      ${m.usd.usd ? `<h3>Exposición en dólares</h3><table><tbody><tr><td>Pendiente en dólares</td><td class="n">USD ${esc(nf.format(m.usd.usd))}</td></tr><tr><td>Tipo de cambio implícito</td><td class="n">${esc(nf.format(m.usd.tc))}</td></tr><tr><td>Equivalente en pesos</td><td class="n">${esc(money(m.usd.ars))}</td></tr></tbody></table>` : ''}
      <h3>Partidas a gestionar</h3><table><tbody>
        <tr><td>Retenciones duplicadas a anular por YPF (${v.retdup.length})</td><td class="n">${esc(money(v.retdupTotal))}</td></tr>
        <tr><td>Créditos sin identificar (${v.sin.length})</td><td class="n">${esc(money(sumBy(v.sin, (f) => f.rec.imp)))}</td></tr>
        <tr><td>Comprobantes sin registrar en Nodum (${v.noCargadas.length})</td><td class="n">${esc(money(sumBy(v.noCargadas, 'orig')))}</td></tr>
        <tr><td>Pendiente de acreditar con más de 7 días (${v.viejos.length})</td><td class="n">${esc(money(sumBy(v.viejos, 'imp')))}</td></tr>
        <tr><td>Cheques rechazados abiertos (${m.chrAb.length})</td><td class="n">${esc(money(sumBy(m.chrAb, 'imp')))}</td></tr>
      </tbody></table>
    </section>`;
  }
  function textoResumen(v, corr) {
    const m = S._m || metricas(v);
    const L = [`Cuenta corriente con YPF – conciliación del ${fechaHora(corr.fecha)}`, '',
      `Deuda según YPF: ${money(v.deudaYPF)} (Combustibles ${money(m.porU.CO)} · Petroquímicos ${money(m.porU.PQ)})`,
      `Pendiente de acreditar: ${money(v.paCO + v.paPQ)} (Combustibles ${money(v.paCO)} · Petroquímicos ${money(v.paPQ)})`,
      `Deuda real: ${money(v.deudaReal)} (Combustibles ${money(m.porU.CO - v.paCO)} · Petroquímicos ${money(m.porU.PQ - v.paPQ)})`, '',
      'Vencimientos:', ...m.vencG.map((b) => `  ${b.l}: ${money(b.tot)} (Combustibles ${money(b.co)} · Petroquímicos ${money(b.pq)})`),
      `  Exigible neto de lo acreditado: ${money(m.neto.tot)} (Combustibles ${money(m.neto.co)} · Petroquímicos ${money(m.neto.pq)})`];
    if (m.usd.usd) L.push('', `Exposición en dólares: USD ${nf.format(m.usd.usd)} (TC implícito ${nf.format(m.usd.tc)})`);
    L.push('', 'Partidas a gestionar:',
      `  Retenciones duplicadas a anular por YPF: ${v.retdup.length} · ${money(v.retdupTotal)}`,
      `  Créditos sin identificar: ${v.sin.length}`,
      `  Comprobantes sin registrar en Nodum: ${v.noCargadas.length} · ${money(sumBy(v.noCargadas, 'orig'))}`,
      `  Cheques rechazados abiertos: ${m.chrAb.length} · ${money(sumBy(m.chrAb, 'imp'))}`);
    return L.join('\n');
  }

  function tabPend(v) {
    if (!v.pendientes.length) return `<div class="panel empty">No hay pagos pendientes de acreditar.</div>`;
    let html = `<div class="sec-head"><p class="muted small">Pagos registrados en Nodum que YPF todavía no acreditó. Pasan solos a acreditados cuando el recibo aparece en la Extranet.</p></div><div class="tw"><table><thead><tr><th>Fecha</th><th class="n">Días</th><th>Trans</th><th>ANT</th><th>Referencia</th><th>Medio</th><th class="n">Número</th><th class="n">Importe</th><th>Marcar</th></tr></thead><tbody>`;
    for (const u of ['CO', 'PQ']) {
      const ps = v.pendientes.filter((p) => p.u === u); if (!ps.length) continue;
      const tot = ps.reduce((a, p) => a + p.imp, 0);
      html += `<tr class="grp"><td colspan="7">${u === 'CO' ? 'Combustibles · 3000266' : 'Petroquímicos · 0001003'} — ${ps.length} medios de pago</td><td class="n">${esc(money(tot))}</td><td></td></tr>`;
      for (const p of ps) {
        html += `<tr${p.dias > 7 ? ' class="row-old"' : ''}><td>${dmy(p.f)}</td><td class="n"><span class="age ${p.dias > 7 ? 'old' : ''}">${p.dias}</span></td><td class="num">${p.trans}</td><td class="num">${p.ant || ''}</td><td>${esc(p.ref || '')}</td><td>${esc(medio(p.tipo))}${p.serie && p.serie !== '0' ? ' <span class="muted small">' + esc(p.serie) + '</span>' : ''}</td><td class="n">${p.nro || ''}</td><td class="n">${esc(nf.format(p.imp))}</td>
          <td><div class="edit"><button class="mini" data-a="pago" data-op="ok" data-k="${esc(p.key)}">Acreditado</button><button class="mini" data-a="pago" data-op="chr" data-k="${esc(p.key)}">Rechazado</button></div></td></tr>`;
      }
    }
    html += `</tbody></table></div>`;
    const manuales = Object.entries(S.lineas).filter(([, l]) => l.okBy && (l.est === 'CHR' || l.okVia === 'manual')).sort((a, b) => (a[1].okF < b[1].okF ? 1 : -1)).slice(0, 30);
    if (manuales.length) {
      html += `<details><summary class="small">Marcados a mano (${manuales.length})</summary><div class="tw" style="margin-top:8px"><table><thead><tr><th>Unidad</th><th>Fecha</th><th>Trans</th><th>Medio</th><th class="n">Importe</th><th>Estado</th><th>Por</th><th></th></tr></thead><tbody>
        ${manuales.map(([k, l]) => `<tr><td>${l.u}</td><td>${dmy(l.f)}</td><td class="num">${l.trans}</td><td>${esc(medio(l.tipo))} ${l.nro || ''}</td><td class="n">${esc(nf.format(l.imp))}</td><td>${l.est === 'CHR' ? '<span class="pill p-bad">Rechazado</span>' : '<span class="pill p-ok">Acreditado</span>'} <span class="muted small">${dmy(l.okF)}</span></td><td class="small">${esc(l.okBy)}</td><td><button class="mini" data-a="pago" data-op="pa" data-k="${esc(k)}">Volver a pendiente</button></td></tr>`).join('')}
      </tbody></table></div></details>`;
    }
    return html;
  }

  function editorRecibo(f) {
    const nro = f.nro; const r = f.rec || {};
    const cands = Object.entries(S.lineas).filter(([, l]) => l.est === 'PA' && l.tipo !== 'redo').sort((a, b) => Math.abs(a[1].imp - r.imp) - Math.abs(b[1].imp - r.imp)).slice(0, 40);
    return `<tr><td colspan="9" style="background:var(--soft)"><div class="edit">
      <label class="small" for="etq-${esc(nro)}">Etiqueta</label>
      <select id="etq-${esc(nro)}"><option value="">Sin etiqueta</option>${Object.entries(ETQ).map(([k, t]) => `<option value="${k}" ${r.etq === k ? 'selected' : ''}>${t}</option>`).join('')}</select>
      <label class="small" for="nota-${esc(nro)}">Nota / cliente</label>
      <input type="text" id="nota-${esc(nro)}" value="${esc(r.nota || '')}" placeholder="Ej.: PD Álvarez" size="22">
      <label class="small" for="vinc-${esc(nro)}">o vincular a un pago pendiente</label>
      <select id="vinc-${esc(nro)}"><option value="">—</option>${cands.map(([k, l]) => `<option value="${esc(k)}">${l.u} ${dmy(l.f)} · ${esc(medio(l.tipo))} ${l.nro || ''} · ${nf.format(l.imp)}</option>`).join('')}</select>
      ${selUnidad(nro, r, f)}
      <button class="mini primary" data-a="guardarRec" data-n="${esc(nro)}" data-f="${esc(f.fdoc || '')}">Guardar</button>
      ${r.match ? `<button class="mini" data-a="desvincular" data-n="${esc(nro)}">Desvincular</button>` : ''}
      <button class="mini" data-a="cerrarEd">Cancelar</button></div></td></tr>`;
  }

  function selUnidad(nro, r, f) {
    const auto = f.recibo ? (f.pago ? f.pago.u : E.unidadDeDiv(f.div)) : E.unidadDeDiv(f.div);
    return `<label class="small" for="uni-${esc(nro)}">Unidad</label>
      <select id="uni-${esc(nro)}"><option value="">Automática${auto ? ' (' + uNombre(auto) + ')' : ' (sin asignar)'}</option><option value="CO" ${r.u === 'CO' ? 'selected' : ''}>Combustibles</option><option value="PQ" ${r.u === 'PQ' ? 'selected' : ''}>Petroquímicos</option></select>`;
  }
  function editorFactura(f) {
    const nro = f.nro; const r = f.rec || {};
    return `<tr><td colspan="9" style="background:var(--soft)"><div class="edit">
      <label class="small" for="etq-${esc(nro)}">Marca</label>
      <select id="etq-${esc(nro)}"><option value="">Sin marca</option>${Object.entries(ETQF).map(([k, t]) => `<option value="${k}" ${r.etq === k ? 'selected' : ''}>${t}</option>`).join('')}</select>
      <label class="small" for="nota-${esc(nro)}">Nota</label>
      <input type="text" id="nota-${esc(nro)}" value="${esc(r.nota || '')}" placeholder="Ej.: Tesorería confirmó el débito el 15/10" size="30">
      ${selUnidad(nro, r, f)}
      <button class="mini primary" data-a="guardarRec" data-n="${esc(nro)}" data-f="${esc(f.fdoc || '')}">Guardar</button>
      <button class="mini" data-a="cerrarEd">Cancelar</button></div></td></tr>`;
  }
  function tabChr(m) {
    const ver = S.chrVer || 'ab';
    const cs = ver === 'ab' ? m.chrAb : m.chrs;
    let html = `<div class="sec-head"><p class="muted small" style="margin:0">Cheques que YPF rechazó, desde que se detectan hasta que se recuperan y se cierra el circuito en Proveedores. Se cierra cuando se emite en Nodum el form “Cheques Devuelto Proveedor”. Cambiá el estado y dejá la justificación; los cerrados dejan de figurar como abiertos.</p>
      <div class="edit"><label class="small" for="f-chrver">Mostrar</label><select id="f-chrver"><option value="ab" ${ver === 'ab' ? 'selected' : ''}>Abiertos (${m.chrAb.length})</option><option value="todos" ${ver === 'todos' ? 'selected' : ''}>Todos (${m.chrs.length})</option></select>
      <button class="mini" data-a="chrForm">${S.chrForm ? 'Cerrar' : 'Agregar cheque rechazado anterior'}</button></div></div>`;
    if (S.chrForm) {
      html += `<div class="panel card"><h3>Agregar cheque rechazado</h3><p class="muted small" style="margin:0">Para cheques rechazados antes del punto cero o fuera de la query del mes.</p><div class="edit">
        <label class="small" for="nc-u">Unidad</label><select id="nc-u"><option value="CO">Combustibles</option><option value="PQ">Petroquímicos</option></select>
        <label class="small" for="nc-f">Fecha del pago</label><input type="date" id="nc-f">
        <label class="small" for="nc-tipo">Medio</label><select id="nc-tipo"><option value="chdifter">Cheque dif. terceros</option><option value="chdifele">Echeq diferido</option><option value="chdifrec">Cheque dif. recibido</option><option value="chrecter">Cheque común terceros</option></select>
        <label class="small" for="nc-nro">Número</label><input type="text" id="nc-nro" size="10" inputmode="numeric">
        <label class="small" for="nc-imp">Importe</label><input type="text" id="nc-imp" size="12" placeholder="2.000.000,00">
        <label class="small" for="nc-trans">Trans (opcional)</label><input type="text" id="nc-trans" size="9" inputmode="numeric">
        <label class="small" for="nc-ant">ANT (opcional)</label><input type="text" id="nc-ant" size="6" inputmode="numeric">
        <label class="small" for="nc-ref">Referencia</label><input type="text" id="nc-ref" size="12" placeholder="LP 07-11">
        <label class="small" for="nc-est">Estado</label><select id="nc-est">${Object.entries(CHRE).map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}</select>
        <label class="small" for="nc-nota">Justificación</label><input type="text" id="nc-nota" size="28">
        <button class="mini primary" data-a="agregarChr">Agregar</button></div></div>`;
    }
    if (!cs.length) return html + `<div class="panel empty">${ver === 'ab' ? 'No hay cheques rechazados abiertos.' : 'No hay cheques rechazados.'}</div>`;
    html += `<div class="tw"><table><thead><tr><th>Unidad</th><th>Fecha</th><th>Trans / ANT</th><th>Referencia</th><th>Cheque</th><th class="n">Importe</th><th class="n">Días</th><th>Estado</th><th>Justificación</th><th></th></tr></thead><tbody>`;
    for (const c of cs) {
      html += `<tr><td>${esc(uNombre(c.u))}</td><td>${dmy(c.f)}</td><td class="num small">${c.trans || ''}${c.ant ? ' / ' + c.ant : ''}</td><td>${esc(c.ref || '')}</td><td>${esc(medio(c.tipo))} <span class="num">${c.nro || ''}</span></td><td class="n">${esc(nf.format(c.imp))}</td><td class="n"><span class="age ${c.chrE !== 'cerr' && c.dias > 30 ? 'old' : ''}">${c.dias == null ? '' : c.dias}</span></td>
        <td><select data-chre="${esc(c.key)}" aria-label="Estado">${Object.entries(CHRE).map(([k, t]) => `<option value="${k}" ${c.chrE === k ? 'selected' : ''}>${t}</option>`).join('')}</select></td>
        <td><input type="text" data-chrn="${esc(c.key)}" value="${esc(c.chrNota || '')}" size="24" aria-label="Justificación">${c.chrBy ? `<div class="muted small">${esc(c.chrBy)} · ${dmy(c.chrF)}</div>` : ''}</td>
        <td><button class="mini primary" data-a="guardarChr" data-k="${esc(c.key)}">Guardar</button></td></tr>`;
    }
    return html + `<tr class="tot"><td colspan="5">Total ${ver === 'ab' ? 'abiertos' : ''}</td><td class="n">${esc(nf.format(sumBy(cs, 'imp')))}</td><td colspan="4"></td></tr></tbody></table></div>`;
  }
  function tabExt(v) {
    const F = S.filtroExt;
    const filas = v.filas.filter((f) => F === 'todo' || (F === 'rec' && f.recibo) || (F === 'deuda' && !f.recibo) || (F === 'revisar' && ['sin', 'posibleDup', 'parcial', 'retdup', 'etiquetado'].includes(f.estado)));
    const opt = [['revisar', 'A revisar'], ['rec', 'Recibos'], ['deuda', 'Facturas y notas de débito'], ['todo', 'Todo']];
    let html = `<div class="sec-head"><div class="edit"><label class="small" for="f-ext">Mostrar</label><select id="f-ext">${opt.map((o) => `<option value="${o[0]}" ${F === o[0] ? 'selected' : ''}>${o[1]}</option>`).join('')}</select></div><span class="muted small">${filas.length} de ${v.filas.length} partidas</span></div>`;
    if (!filas.length) return html + `<div class="panel empty">Nada para revisar en la Extranet.</div>`;
    html += `<div class="tw"><table><thead><tr><th>Estado</th><th>Tipo</th><th>Nro. legal</th><th>Div.</th><th>T. Pos.</th><th>F. Doc.</th><th>F. Venc.</th><th class="n">Original ARS</th><th>Cruza con / etiqueta</th></tr></thead><tbody>`;
    for (const f of filas) {
      const r = f.rec || {};
      const detalle = f.pago ? `${f.pago.u} · trans ${f.pago.trans}${f.pago.ant ? ' · ANT ' + f.pago.ant : ''} · ${esc(medio(f.pago.tipo))} ${f.pago.nro || ''}${f.estado === 'parcial' ? ` <span class="muted small">(pago de ${esc(nf.format(f.pago.imp))})</span>` : ''}${f.pago.ref ? ' · ' + esc(f.pago.ref) : ''}` :
        r.etq ? `<b>${esc(ETQ[r.etq] || r.etq)}</b>${r.nota ? ' · ' + esc(r.nota) : ''}` : f.estado === 'posibleDup' ? 'Ya hay otro recibo con esta retención' : f.recibo ? '<span class="muted">—</span>' : '';
      const editable = true;
      const uTxt = `<span class="muted small">${esc(uNombre(f.u))}${f.uManual ? ' (a mano)' : ''}</span>`;
      const marca = !f.recibo && ETQF[r.etq] ? ` <span class="pill p-info">${esc('Débito informado')}</span>${r.nota ? ' <span class="small">' + esc(r.nota) + '</span>' : ''}` : '';
      html += `<tr><td>${estadoPill(f)}${marca}</td><td>${esc(f.tipo)}</td><td class="num">${esc(f.nro)}</td><td>${esc(f.div)}<br>${uTxt}</td><td class="small">${esc(f.tpos)}</td><td>${dmy(f.fdoc)}</td><td>${dmy(f.fvenc)}</td><td class="n">${esc(nf.format(f.orig))}</td>
        <td><div class="edit"><span class="small">${detalle}</span>${editable ? `<button class="mini" data-a="editar" data-n="${esc(f.nro)}">${!f.recibo ? 'Marcar / unidad' : f.pago || r.etq ? 'Cambiar' : 'Identificar'}</button>` : ''}</div></td></tr>`;
      if (S.editando === f.nro) html += f.recibo ? editorRecibo(f) : editorFactura(f);
    }
    return html + `</tbody></table></div>`;
  }

  // Lleva la pantalla a la sección pedida y la resalta un momento
  function irA(go) {
    if (!go) return;
    const els = go === 'x-viejo' ? [...document.querySelectorAll('tr.row-old')] : [document.getElementById(go)].filter(Boolean);
    const el = els[0] || document.getElementById('tab-' + S.tab);
    if (!el) return;
    el.scrollIntoView({ block: 'start', behavior: 'smooth' });
    els.forEach((e) => { e.style.outline = '2px solid var(--accent)'; e.style.outlineOffset = '-2px'; setTimeout(() => { e.style.outline = ''; e.style.outlineOffset = ''; }, 2500); });
  }
  function lista(titulo, ayuda, rows, cols, id) {
    if (!rows.length) return '';
    return `<div class="panel card"${id ? ` id="${id}"` : ''}><div class="sec-head"><h3>${titulo} <span class="muted small">(${rows.length})</span></h3></div>${ayuda ? `<p class="muted small" style="margin:0">${ayuda}</p>` : ''}
      <div class="tw"><table><thead><tr>${cols.map((c) => `<th class="${c[2] || ''}">${c[0]}</th>`).join('')}</tr></thead><tbody>${rows.map((r) => `<tr>${cols.map((c) => `<td class="${c[2] || ''}">${c[1](r)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
  }
  function tabExc(v) {
    const recCols = [['Nro. legal', (f) => `<span class="num">${esc(f.nro)}</span>`], ['T. Pos.', (f) => esc(f.tpos)], ['Div.', (f) => esc(f.div)], ['F. Doc.', (f) => dmy(f.fdoc)], ['Importe', (f) => esc(nf.format(f.rec.imp)), 'n'], ['', (f) => `<button class="mini" data-a="irRec" data-n="${esc(f.nro)}">Identificar</button>`]];
    const html = [
      lista('Recibos sin identificar', 'Plata que YPF acreditó y no encontré en Nodum: pago directo de un cliente (PD), un pago todavía no contabilizado o algo para reclamar.', v.sin, recCols, 'x-sin'),
      lista('Posibles retenciones duplicadas', 'La retención ya está cruzada con otro recibo por el mismo importe.', v.posibleDup, recCols, 'x-dup'),
      lista('Retenciones duplicadas: pedir anulación a YPF', 'Mientras YPF no las anule, la deuda que muestra la Extranet es menor que la real.', v.retdup, [['Nro. legal', (f) => `<span class="num">${esc(f.nro)}</span>`], ['T. Pos.', (f) => esc(f.tpos)], ['F. Doc.', (f) => dmy(f.fdoc)], ['Días en la Extranet', (f) => `<span class="age old">${E.daysBetween(f.fdoc, hoyISO())}</span>`, 'n'], ['Importe', (f) => esc(nf.format(f.rec.imp)), 'n']], 'x-retdup'),
      lista('Saldos parciales', 'YPF aplicó una parte del pago y dejó el resto como saldo a favor.', v.parciales, [['Nro. legal', (f) => `<span class="num">${esc(f.nro)}</span>`], ['T. Pos.', (f) => esc(f.tpos)], ['Pago en Nodum', (f) => `${f.pago.u} trans ${f.pago.trans} · ${esc(nf.format(f.pago.imp))}`], ['Saldo en Extranet', (f) => esc(nf.format(f.rec.imp)), 'n']]),
      lista('En la Extranet y no cargados en Nodum', 'Facturas MN y notas de débito dentro del período del navegador contable que no aparecen en la cuenta 21110101.', v.noCargadas, [['Comprobante', (e) => `${esc(e.tipo)} <span class="num">${esc(e.nro)}</span>`], ['Div.', (e) => esc(e.div)], ['F. Doc.', (e) => dmy(e.fdoc)], ['Importe', (e) => esc(nf.format(e.orig)), 'n']], 'x-nonodum'),
      lista('Diferencias con Nodum', '', v.alertas, [['Unidad', (a) => esc(a.u)], ['Detalle', (a) => esc(a.texto)], ['Importe', (a) => esc(nf.format(a.imp)), 'n']]),
    ].join('');
    return html || `<div class="panel empty">No hay excepciones abiertas.</div>`;
  }

  function tabMov(v) {
    if (!v.mov) return `<div class="panel empty">Hace falta una conciliación anterior para comparar.</div>`;
    const m = v.mov;
    const cols = [['Comprobante', (e) => `${esc(e.tipo)} <span class="num">${esc(e.nro)}</span>`], ['Div.', (e) => esc(e.div)], ['F. Doc.', (e) => dmy(e.fdoc)], ['Importe', (e) => esc(nf.format(e.orig)), 'n']];
    return `<p class="muted small">Comparado con la conciliación del ${esc(fechaHora(m.desde))}. Lo que salió de la Extranet es lo que YPF compensó o cobró en el medio.</p>
      <div class="panel card"><div class="kv">
        <span>Facturas y notas de débito que salieron</span><span class="n">${esc(money(m.tFS))}</span>
        <span>Recibos y saldos que salieron</span><span class="n">${esc(money(m.tRS))}</span>
        <span>Facturas y notas de débito nuevas</span><span class="n">${esc(money(m.tFN))}</span>
        <span>Recibos nuevos</span><span class="n">${esc(money(m.tRN))}</span>
      </div></div>
      <div class="cards2">${lista('Salieron: facturas y ND', '', m.facturasSalieron, cols)}${lista('Salieron: recibos', '', m.recibosSalieron, cols)}${lista('Nuevas: facturas y ND', '', m.facturasNuevas, cols)}${lista('Nuevos: recibos', '', m.recibosNuevos, cols)}</div>`;
  }

  function tabHist() {
    if (!S.corridas.length) return `<div class="panel empty">Sin conciliaciones.</div>`;
    return `<div class="tw"><table><thead><tr><th>Fecha</th><th>Por</th><th class="n">Deuda YPF</th><th class="n">Pend. CO</th><th class="n">Pend. PQ</th><th class="n">Deuda real</th><th>Archivos</th></tr></thead><tbody>
      ${S.corridas.map((c) => { const r = c.resumen || {}; return `<tr><td>${esc(fechaHora(c.fecha))}</td><td class="small">${esc(c.usuario || c.origen || '')}</td><td class="n">${esc(mill(r.deudaYPF))}</td><td class="n">${esc(mill(r.paCO))}</td><td class="n">${esc(mill(r.paPQ))}</td><td class="n">${esc(mill(r.deudaReal))}</td><td class="small">${esc((c.archivos || []).join(' · '))}</td></tr>`; }).join('')}
    </tbody></table></div>`;
  }

  /* ---------- eventos ---------- */
  app.addEventListener('click', async (ev) => {
    const b = ev.target.closest('[data-a]'); if (!b) return;
    const a = b.dataset.a;
    if (a === 'vista') { S.vista = b.dataset.v; try { localStorage.setItem('ypf-vista', S.vista); } catch (e) { /* sin almacenamiento */ } render(); }
    else if (a === 'rep') { S.reporte = b.dataset.r; render(); const d = $('#detalle'); if (d) d.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
    else if (a === 'resumen') { window.print(); }
    else if (a === 'copiarRes') { const t = textoResumen(S._v, S._corr); try { await navigator.clipboard.writeText(t); toast('Resumen copiado'); } catch (e) { toast('No se pudo copiar'); } }
    else if (a === 'irConc') { S.vista = 'conc'; S.tab = b.dataset.t || 'pend'; render(); }
    else if (a === 'borrarLocal') { if (window.confirm('¿Borrar todas las conciliaciones guardadas en este navegador?')) { try { localStorage.removeItem('ypf-monitor-local'); } catch (e) { /* nada */ } S.pagosDocs = {}; S.recibosDocs = {}; S.corridas = []; rebuild(); render(); toast('Datos locales borrados'); } }
    else if (a === 'toggleCarga') { S.cargaAbierta = !S.cargaAbierta; render(); }
    else if (a === 'conciliar') conciliar();
    else if (a === 'limpiar') { S.files = { extranet: null, qCO: null, qPQ: null, nCO: null, nPQ: null }; S.fileMsgs = []; render(); }
    else if (a === 'excel') { if (S._v) { try { await descargarExcel(S._v, S._corr); } catch (e) { console.error(e); toast('No se pudo generar el Excel: ' + e.message); } } }
    else if (a === 'tab') { S.tab = b.dataset.t; if (S.tab === 'ext' && b.classList.contains('chip')) S.filtroExt = 'revisar'; render(); irA(b.dataset.go); }
    else if (a === 'pago') accionPago(b.dataset.k, b.dataset.op);
    else if (a === 'editar') { S.editando = S.editando === b.dataset.n ? null : b.dataset.n; render(); }
    else if (a === 'irRec') { S.tab = 'ext'; S.filtroExt = 'revisar'; S.editando = b.dataset.n; render(); }
    else if (a === 'cerrarEd') { S.editando = null; render(); }
    else if (a === 'guardarRec') guardarRecibo(b.dataset.n, b.dataset.f);
    else if (a === 'guardarChr') guardarChr(b.dataset.k);
    else if (a === 'agregarChr') agregarChr();
    else if (a === 'chrForm') { S.chrForm = !S.chrForm; render(); }
    else if (a === 'desvincular') desvincular(b.dataset.n);
    else if (a === 'login') { try { const prov = new firebase.auth.GoogleAuthProvider(); prov.setCustomParameters({ prompt: 'select_account' }); await fb.auth.signInWithPopup(prov); } catch (e) { toast('No se pudo ingresar: ' + e.message); } }
    else if (a === 'salir') { await fb.auth.signOut(); }
  });
  app.addEventListener('change', async (ev) => {
    const t = ev.target;
    if (t.id === 'f-in') { await leerArchivos([...t.files]); t.value = ''; }
    if (t.id === 'f-ext') { S.filtroExt = t.value; render(); }
    if (t.id === 'f-chrver') { S.chrVer = t.value; render(); }
    if (t.id === 'f-seed' && t.files[0]) {
      try { const seed = JSON.parse(await t.files[0].text()); await Store.importar(seed); toast('Estado inicial importado'); }
      catch (e) { toast('No se pudo importar: ' + e.message); }
    }
  });
  app.addEventListener('dragover', (ev) => { const d = ev.target.closest('#drop'); if (d) { ev.preventDefault(); d.classList.add('over'); } });
  app.addEventListener('dragleave', (ev) => { const d = ev.target.closest('#drop'); if (d) d.classList.remove('over'); });
  app.addEventListener('drop', (ev) => { const d = ev.target.closest('#drop'); if (d) { ev.preventDefault(); d.classList.remove('over'); leerArchivos([...ev.dataTransfer.files]); } });

  Store.init();
})();
