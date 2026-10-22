/* Motor de conciliación YPF – Copsa. Funciones puras: parseo de archivos, ingesta y análisis. */
(function (root) {
  const R2 = (x) => Math.round((Number(x) || 0) * 100) / 100;
  const CERT = 'certribr';
  const UNIDAD_TIT = { '3000266': 'CO', '0001003': 'PQ' };

  function pad(n, w) { return String(n).padStart(w, '0'); }
  function toISO(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      const ms = Date.UTC(1899, 11, 30) + Math.round(v * 86400000);
      return new Date(ms).toISOString().slice(0, 10);
    }
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return `${m[3]}-${pad(m[2], 2)}-${pad(m[1], 2)}`;
    return null;
  }
  function num(v) {
    if (typeof v === 'number') return v;
    if (v == null || v === '') return 0;
    let s = String(v).trim();
    if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s.replace(/[^\d.\-]/g, ''));
    return isNaN(n) ? 0 : n;
  }
  function daysBetween(a, b) {
    if (!a || !b) return null;
    return Math.round((Date.parse(b) - Date.parse(a)) / 86400000);
  }
  function lineKey(u, trans, tipo, nro, imp) {
    return `${u}-${parseInt(trans, 10)}-${tipo}-${parseInt(nro, 10) || 0}-${Math.round(R2(imp) * 100)}`;
  }
  function pagoDocId(line) { return `${line.u}-${(line.f || '0000-00').slice(0, 7)}`; }
  function reciboDocId(rec) { return (rec.f || '0000-00').slice(0, 7); }
  function unidadDeDiv(div) {
    if (['PQ', 'ME'].includes(div)) return 'PQ';
    if (['MN', 'MA', 'GR', 'YF'].includes(div)) return 'CO';
    return null;
  }

  /* ---------- Detección y parseo de archivos ---------- */
  function sheetRows(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  }
  function detect(XLSX, wb, filename) {
    for (const name of wb.SheetNames) {
      const rows = sheetRows(XLSX, wb.Sheets[name]);
      const head = rows.slice(0, 20).map((r) => (r || []).map((c) => (c == null ? '' : String(c))).join(' | ')).join('\n');
      if (/Medio pago por titular/i.test(head)) return parseQuery(rows, filename);
      if (/Estado de Cuenta/i.test(head) && /Titular:/i.test(head)) return parseNav(rows, filename);
      if (/Nro\. Legal/.test(head) && /T\. Pos\./.test(head)) return parseExtranet(rows, filename);
    }
    return { kind: 'desconocido', filename };
  }

  function parseQuery(rows, filename) {
    let tit = null, desde = null, hasta = null, h = -1;
    rows.forEach((r, i) => {
      const s = (r && r[0] != null) ? String(r[0]) : '';
      let m = s.match(/Titular\s*=\s*(\d+)/); if (m) tit = m[1];
      m = s.match(/Fecha_Desde\s*=\s*(.+)/); if (m) desde = toISO(m[1]);
      m = s.match(/Fecha_Hasta\s*=\s*(.+)/); if (m) hasta = toISO(m[1]);
      if (h < 0 && r && r.includes('fec_doc') && r.includes('nro_trans')) h = i;
    });
    if (h < 0) return { kind: 'error', filename, error: 'No encontré la fila de encabezados (fec_doc, nro_trans…).' };
    const H = rows[h]; const c = (n) => H.indexOf(n);
    const lines = [];
    for (const r of rows.slice(h + 1)) {
      if (!r) continue;
      const trans = r[c('nro_trans')];
      if (trans == null || isNaN(parseInt(trans, 10))) continue;
      const t = String(r[c('cod_tit')] || tit || '').trim();
      const u = UNIDAD_TIT[t] || (/PQ/.test(String(r[c('nom_tit')] || '')) ? 'PQ' : 'CO');
      lines.push({
        u, f: toISO(r[c('fec_doc')]), trans: parseInt(trans, 10),
        tipo: String(r[c('cod_doca')] || '').trim(), nom: String(r[c('nom_docum')] || '').trim(),
        serie: String(r[c('serie_doca')] == null ? '0' : r[c('serie_doca')]).trim(),
        nro: parseInt(r[c('nro_doca')], 10) || 0, imp: R2(num(r[c('imp_mov_mn')])),
      });
    }
    const u = UNIDAD_TIT[tit] || (lines[0] && lines[0].u) || null;
    if (!desde && lines.length) desde = lines.map((l) => l.f).sort()[0];
    if (!hasta && lines.length) hasta = lines.map((l) => l.f).sort().slice(-1)[0];
    return { kind: 'query', filename, u, desde, hasta, lines, total: R2(lines.reduce((a, l) => a + l.imp, 0)) };
  }

  function parseNav(rows, filename) {
    let tit = null, desde = null, hasta = null, h = -1;
    rows.forEach((r, i) => {
      const s = (r && r[0] != null) ? String(r[0]) : '';
      let m = s.match(/Titular:\s*(\d+)/); if (m) tit = m[1];
      m = s.match(/^Desde\s+(.+)/); if (m) desde = toISO(m[1]);
      m = s.match(/^Hasta\s+(.+)/); if (m) hasta = toISO(m[1]);
      if (h < 0 && r && String(r[0]) === 'Fecha' && String(r[1]) === 'Trans') h = i;
    });
    if (h < 0) return { kind: 'error', filename, error: 'No encontré la fila de encabezados (Fecha, Trans…).' };
    const pagos = {}, anticipos = {}, comprobantes = [];
    let saldoInicial = null, saldoFinal = null;
    for (const r of rows.slice(h + 1)) {
      if (!r) continue;
      const ref = String(r[2] || '').trim();
      if (ref === 'Saldo Inicial') { saldoInicial = num(r[11]); continue; }
      if (ref === 'TOTAL') { saldoFinal = num(r[11]); continue; }
      const trans = parseInt(r[1], 10); if (isNaN(trans)) continue;
      const docu = String(r[3] || '').trim();
      const debe = num(r[9]), haber = num(r[10]);
      if (docu === 'Pagos') {
        pagos[trans] = { ant: parseInt(r[8], 10) || null, debe: R2(debe), f: toISO(r[0]) };
      } else if (/^Anticipo/i.test(docu)) {
        anticipos[parseInt(r[5], 10)] = ref;
      } else if (/Factura|Nota de D[eé]bito|Nota de Cr[eé]dito/i.test(docu)) {
        const letra = (docu.match(/\s([A-Z])$/) || [, 'A'])[1];
        const nroLegal = letra + pad(String(r[4] || '').trim(), 4) + pad(parseInt(r[5], 10) || 0, 8);
        comprobantes.push({ nroLegal, docu, ref, f: toISO(r[0]), imp: R2(haber || debe) });
      }
      if (r[11] != null && ref !== 'TOTAL') saldoFinal = num(r[11]);
    }
    return { kind: 'nav', filename, u: UNIDAD_TIT[tit] || null, desde, hasta, pagos, anticipos, comprobantes, saldoInicial, saldoFinal };
  }

  function parseExtranet(rows, filename) {
    let h = rows.findIndex((r) => r && r.includes('Nro. Legal') && r.includes('T. Pos.'));
    const H = rows[h].map((x) => (x == null ? '' : String(x).trim()));
    const c = (pred) => H.findIndex(pred);
    const ix = {
      tipo: c((x) => x === 'Tipo'), nro: c((x) => x === 'Nro. Legal'), clase: c((x) => x === 'Clase Doc'),
      div: c((x) => x === 'Div.'), texto: c((x) => x === 'Texto/SGC'), tpos: c((x) => x === 'T. Pos.'),
      fdoc: c((x) => x === 'F. Doc.'), fvenc: c((x) => x === 'F. Venc.'), mon: c((x) => x === 'Moneda'),
      ars: c((x) => /^Pendiente ARS/.test(x)), usd: c((x) => x === 'Pendiente USD'), orig: c((x) => x === 'Original ARS'),
    };
    const out = [], omitidas = [];
    const esNum = (x) => x !== null && x !== '' && !isNaN(num(x));
    for (const r of rows.slice(h + 1)) {
      if (!r) continue;
      const nro = String(r[ix.nro] || '').trim();
      const g = (k) => (ix[k] >= 0 ? r[ix[k]] : null);
      // Se toma toda fila con número de documento e importe, cualquiera sea el formato del número
      // (facturas A…, recibos X…, compensaciones 153DP00-…, etc.)
      if (!nro || nro === 'Nro. Legal' || /^total/i.test(String(g('tipo') || '')) || !esNum(g('orig'))) {
        if (esNum(g('orig')) && !/^total/i.test(String(g('tipo') || ''))) omitidas.push({ tipo: String(g('tipo') || '').trim(), nro, orig: R2(num(g('orig'))) });
        continue;
      }
      out.push([
        String(g('tipo') || '').trim(), nro, String(g('clase') || '').trim(), String(g('div') || '').trim(),
        String(g('texto') || '').trim(), String(g('tpos') || '').trim(), toISO(g('fdoc')), toISO(g('fvenc')),
        String(g('mon') || '').trim(), R2(num(g('ars'))), R2(num(g('usd'))), R2(num(g('orig'))),
      ]);
    }
    const total = R2(out.reduce((a, x) => a + x[11], 0));
    // Control: filas con importe que no se pudieron leer (se ignora una fila de total igual a la suma)
    const sinLeer = omitidas.filter((o) => Math.abs(o.orig - total) > 0.05);
    return { kind: 'extranet', filename, rows: out, total, sinLeer };
  }

  function extObj(a) {
    return { tipo: a[0], nro: a[1], clase: a[2], div: a[3], texto: a[4], tpos: a[5], fdoc: a[6], fvenc: a[7], mon: a[8], ars: a[9], usd: a[10], orig: a[11] };
  }
  // Un "Saldo" con importe positivo es lo que queda por pagar de una factura compensada en parte: es deuda, no un crédito
  function esRecibo(e) { if (e.orig > 0 && /^Saldo/i.test(e.tipo)) return false; return e.clase === 'RC' || e.clase === 'CL' || /^(Recibo|Saldo)/i.test(e.tipo) || (e.orig < 0 && !/Factura|d[eé]bito/i.test(e.tipo)); }

  /* ---------- Ingesta: suma los archivos al estado guardado ---------- */
  function ingest(state, inputs, today, nowISO) {
    const lineas = JSON.parse(JSON.stringify(state.lineas || {}));
    const recibos = JSON.parse(JSON.stringify(state.recibos || {}));
    const chP = new Set(), chR = new Set(), alertas = [];
    const touchL = (k) => chP.add(pagoDocId(lineas[k]));
    const touchR = (n) => chR.add(reciboDocId(recibos[n]));

    // 1. Pagos de Nodum
    const queryInfo = [];
    for (const q of inputs.queries || []) {
      const keys = new Set();
      const porTrans = {};
      for (const ql of q.lines) {
        const k = lineKey(ql.u, ql.trans, ql.tipo, ql.nro, ql.imp);
        keys.add(k);
        porTrans[ql.trans] = R2((porTrans[ql.trans] || 0) + ql.imp);
        if (!lineas[k]) {
          lineas[k] = { u: ql.u, f: ql.f, trans: ql.trans, tipo: ql.tipo, nom: ql.nom, serie: ql.serie, nro: ql.nro, imp: ql.imp,
            est: ql.tipo === 'redo' ? 'R' : 'PA', ref: '', ant: null, pv: today };
          touchL(k);
        }
      }
      for (const [k, l] of Object.entries(lineas)) {
        if (l.u === q.u && l.est === 'PA' && l.f >= q.desde && l.f <= q.hasta && !keys.has(k)) {
          alertas.push({ tipo: 'noNodum', u: l.u, key: k, texto: `Pago PA que ya no figura en la query de Nodum (trans ${l.trans}, ${l.tipo} ${l.nro})`, imp: l.imp });
        }
      }
      queryInfo.push({ u: q.u, desde: q.desde, hasta: q.hasta, n: q.lines.length, total: q.total, porTrans });
    }

    // 2. Navegador contable: ANT, referencia y control contra la query
    const navInfo = [];
    for (const nv of inputs.navs || []) {
      for (const [k, l] of Object.entries(lineas)) {
        if (l.u !== nv.u) continue;
        const p = nv.pagos[l.trans];
        if (!p) continue;
        const ref = nv.anticipos[p.ant] || l.ref;
        if (l.ant !== p.ant || (!l.ref && ref)) { l.ant = p.ant; if (!l.ref) l.ref = ref; touchL(k); }
      }
      const qi = queryInfo.find((q) => q.u === nv.u);
      if (qi) {
        for (const [t, p] of Object.entries(nv.pagos)) {
          const qt = qi.porTrans[t];
          if (qt == null && p.f >= qi.desde && p.f <= qi.hasta) alertas.push({ tipo: 'difContable', u: nv.u, trans: +t, texto: `Trans ${t}: está en el contable ($${p.debe}) pero no en la query`, imp: p.debe });
          else if (qt != null && Math.abs(qt - p.debe) > 0.05) alertas.push({ tipo: 'difContable', u: nv.u, trans: +t, texto: `Trans ${t}: query $${qt} vs contable $${p.debe}`, imp: R2(qt - p.debe) });
        }
      }
      navInfo.push({ u: nv.u, desde: nv.desde, hasta: nv.hasta, comprobantes: nv.comprobantes.map((c) => c.nroLegal), saldoInicial: nv.saldoInicial, saldoFinal: nv.saldoFinal });
    }

    // 3. Recibos de la Extranet contra pagos de Nodum
    const linkedBy = {};
    for (const r of Object.values(recibos)) if (r.match) (linkedBy[r.match] = linkedBy[r.match] || []).push(r.nro);
    const link = (rec, k, mt) => {
      rec.match = k; rec.mt = mt; rec.mf = today; delete rec.sug;
      (linkedBy[k] = linkedBy[k] || []).push(rec.nro);
      const l = lineas[k];
      if (l.est === 'PA') { l.est = 'OK'; l.okVia = rec.nro; l.okF = today; touchL(k); }
    };
    const near = (a, b) => Math.abs(a - b) < 0.005;
    const byPref = (rec, prefU) => (a, b) => {
      const la = lineas[a], lb = lineas[b];
      return ((lb.est === 'PA') - (la.est === 'PA')) || ((prefU && lb.u === prefU) - (prefU && la.u === prefU)) ||
        Math.abs(daysBetween(la.f, rec.f) || 0) - Math.abs(daysBetween(lb.f, rec.f) || 0);
    };
    function tryMatch(rec) {
      const tp = String(rec.tpos || '').trim();
      const esRet = /^0092-/.test(tp);
      const m = tp.match(/^0*(\d{1,12})(\s+ECHEQ)?$/);
      if (m && !esRet) {
        const n = parseInt(m[1], 10);
        const c = Object.keys(lineas).filter((k) => lineas[k].tipo !== CERT && lineas[k].tipo !== 'efectivo' && lineas[k].nro === n);
        if (c.length) {
          c.sort((a, b) => ((linkedBy[a] ? 1 : 0) - (linkedBy[b] ? 1 : 0)) || Math.abs(lineas[a].imp - rec.imp) - Math.abs(lineas[b].imp - rec.imp));
          link(rec, c[0], near(lineas[c[0]].imp, rec.imp) ? 'cheque' : 'parcial');
          return;
        }
      }
      if (esRet) {
        const all = Object.keys(lineas).filter((k) => lineas[k].tipo === CERT && near(lineas[k].imp, rec.imp));
        const free = all.filter((k) => !linkedBy[k]);
        if (free.length) { free.sort(byPref(rec, unidadDeDiv(rec.div))); link(rec, free[0], 'retencion'); return; }
        if (all.length) rec.sug = 'RETDUP';
        return;
      }
      const c = Object.keys(lineas).filter((k) => lineas[k].tipo !== CERT && !linkedBy[k] && near(lineas[k].imp, rec.imp));
      if (c.length) { c.sort(byPref(rec, null)); link(rec, c[0], 'importe'); }
    }

    const ext = inputs.extranet ? inputs.extranet.rows : [];
    for (const a of ext) {
      const e = extObj(a);
      if (!esRecibo(e)) continue;
      let rec = recibos[e.nro];
      if (!rec) { rec = recibos[e.nro] = { nro: e.nro, f: e.fdoc, div: e.div, tpos: e.tpos, imp: R2(-e.orig), pv: today }; }
      rec.uv = today; rec.tipo = e.tipo; rec.imp = R2(-e.orig); rec.tpos = e.tpos; rec.div = e.div;
      if (rec.match && !lineas[rec.match]) { delete rec.match; delete rec.mt; }
      if (!rec.match && !rec.etq) tryMatch(rec);
      touchR(e.nro);
    }

    const corrida = {
      fecha: nowISO, hoy: today,
      archivos: [inputs.extranet, ...(inputs.queries || []), ...(inputs.navs || [])].filter(Boolean).map((x) => x.filename),
      extranet: ext, queries: queryInfo.map(({ porTrans, ...q }) => q), navs: navInfo, alertas,
    };
    return { lineas, recibos, chP: [...chP], chR: [...chR], corrida };
  }

  /* ---------- Análisis: arma la vista a partir del estado ---------- */
  function analizar(lineas, recibos, corrida, prev, today) {
    const ext = (corrida && corrida.extranet ? corrida.extranet : []).map(extObj);
    const deudaYPF = R2(ext.reduce((a, e) => a + e.orig, 0));
    const deudaTC = R2(ext.reduce((a, e) => a + e.ars, 0));
    const divs = {};
    for (const e of ext) {
      const d = divs[e.div] = divs[e.div] || { div: e.div, debitos: 0, creditos: 0, n: 0 };
      if (e.orig >= 0) d.debitos = R2(d.debitos + e.orig); else d.creditos = R2(d.creditos + e.orig);
      d.n++;
    }
    const filas = ext.map((e) => {
      const v = { ...e, recibo: esRecibo(e) };
      const r = recibos[e.nro] || {};
      // Unidad: la asignada a mano manda; si no, el pago de Nodum con el que cruzó; si no, la división
      v.uManual = r.u === 'CO' || r.u === 'PQ' ? r.u : null;
      if (v.recibo) {
        v.rec = r;
        const l = r.match ? lineas[r.match] : null;
        v.pago = l ? { ...l, key: r.match } : null;
        v.u = v.uManual || (l ? l.u : unidadDeDiv(e.div)) || '—';
        if (r.etq === 'RETDUP') v.estado = 'retdup';
        else if (r.etq) v.estado = 'etiquetado';
        else if (l && r.mt === 'parcial') v.estado = 'parcial';
        else if (l) v.estado = 'cruzado';
        else if (r.sug === 'RETDUP') v.estado = 'posibleDup';
        else v.estado = 'sin';
      } else {
        v.estado = 'deuda'; v.rec = r;
        v.debitado = r.etq === 'DEBITADO';
        v.u = v.uManual || unidadDeDiv(e.div) || '—';
      }
      return v;
    });
    const pendientes = Object.entries(lineas).filter(([, l]) => l.est === 'PA')
      .map(([k, l]) => ({ ...l, key: k, dias: daysBetween(l.f, today) }))
      .sort((a, b) => (a.u > b.u ? 1 : a.u < b.u ? -1 : 0) || (a.f < b.f ? -1 : a.f > b.f ? 1 : 0) || a.trans - b.trans);
    const paCO = R2(pendientes.filter((p) => p.u === 'CO').reduce((a, p) => a + p.imp, 0));
    const paPQ = R2(pendientes.filter((p) => p.u === 'PQ').reduce((a, p) => a + p.imp, 0));
    const retdup = filas.filter((f) => f.estado === 'retdup');
    const noCargadas = [];
    for (const nv of (corrida && corrida.navs) || []) {
      const set = new Set(nv.comprobantes);
      for (const e of ext) {
        const esDD = e.clase === 'DD', esFDMN = e.clase === 'FD' && e.div === 'MN';
        if (!(esDD || esFDMN) || e.mon !== 'ARS') continue;
        if (unidadDeDiv(e.div) !== nv.u) continue;
        if (nv.desde && e.fdoc < nv.desde) continue;
        if (nv.hasta && e.fdoc > nv.hasta) continue;
        if (!set.has(e.nro)) noCargadas.push({ ...e, u: nv.u });
      }
    }
    let mov = null;
    if (prev && prev.extranet) {
      const P = new Map(prev.extranet.map((a) => [a[1], extObj(a)]));
      const C = new Map(ext.map((e) => [e.nro, e]));
      const gone = [...P.values()].filter((e) => !C.has(e.nro));
      const nuevos = ext.filter((e) => !P.has(e.nro));
      const sum = (arr) => R2(arr.reduce((a, e) => a + e.orig, 0));
      mov = {
        desde: prev.fecha,
        facturasSalieron: gone.filter((e) => !esRecibo(e)), recibosSalieron: gone.filter(esRecibo),
        facturasNuevas: nuevos.filter((e) => !esRecibo(e)), recibosNuevos: nuevos.filter(esRecibo),
      };
      mov.tFS = sum(mov.facturasSalieron); mov.tRS = sum(mov.recibosSalieron);
      mov.tFN = sum(mov.facturasNuevas); mov.tRN = sum(mov.recibosNuevos);
    }
    return {
      deudaYPF, deudaTC, paCO, paPQ, deudaReal: R2(deudaYPF - paCO - paPQ),
      retdupTotal: R2(retdup.reduce((a, f) => a + f.rec.imp, 0)),
      divs: Object.values(divs).sort((a, b) => b.debitos - a.debitos),
      filas, pendientes, retdup,
      sin: filas.filter((f) => f.estado === 'sin'),
      posibleDup: filas.filter((f) => f.estado === 'posibleDup'),
      parciales: filas.filter((f) => f.estado === 'parcial'),
      etiquetados: filas.filter((f) => f.estado === 'etiquetado'),
      noCargadas, alertas: (corrida && corrida.alertas) || [], mov,
      viejos: pendientes.filter((p) => p.dias != null && p.dias > 7),
    };
  }

  root.YPFEngine = { detect, parseQuery, parseNav, parseExtranet, ingest, analizar, lineKey, pagoDocId, reciboDocId, toISO, R2, daysBetween, unidadDeDiv };
  if (typeof module !== 'undefined') module.exports = root.YPFEngine;
})(typeof window !== 'undefined' ? window : globalThis);
