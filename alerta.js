const admin = require('firebase-admin');
const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const sa = JSON.parse(process.env.FIREBASE_KEY);
admin.initializeApp({ credential: admin.credential.cert(sa) });
const db = admin.firestore();

const token = process.env.TG_TOKEN;
const chatId = process.env.TG_CHAT_ID;

function getDateStr(offsetDays) {
  const d = new Date();
  d.setUTCHours(d.getUTCHours() - 6);
  d.setUTCDate(d.getUTCDate() + offsetDays);
  return d.toISOString().split('T')[0];
}

function send(text) {
  return new Promise(resolve => {
    const body = JSON.stringify({ chat_id: chatId, text, parse_mode: 'Markdown' });
    const req = https.request({
      hostname: 'api.telegram.org',
      path: '/bot' + token + '/sendMessage',
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
    }, resolve);
    req.write(body);
    req.end();
  });
}

function sendDocument(filePath, caption) {
  return new Promise((resolve, reject) => {
    const fileContent = fs.readFileSync(filePath);
    const boundary = '----FormBoundary' + Date.now();
    const filename = path.basename(filePath);

    let body = '';
    body += '--' + boundary + '\r\n';
    body += 'Content-Disposition: form-data; name="chat_id"\r\n\r\n';
    body += chatId + '\r\n';
    body += '--' + boundary + '\r\n';
    body += 'Content-Disposition: form-data; name="caption"\r\n\r\n';
    body += (caption || '') + '\r\n';

    const bodyStart = Buffer.from(body);
    const fileHeader = Buffer.from(
      '--' + boundary + '\r\n' +
      'Content-Disposition: form-data; name="document"; filename="' + filename + '"\r\n' +
      'Content-Type: application/octet-stream\r\n\r\n'
    );
    const bodyEnd = Buffer.from('\r\n--' + boundary + '--\r\n');
    const fullBody = Buffer.concat([bodyStart, fileHeader, fileContent, bodyEnd]);

    const req = https.request({
      hostname: 'api.telegram.org',
      path: '/bot' + token + '/sendDocument',
      method: 'POST',
      headers: {
        'Content-Type': 'multipart/form-data; boundary=' + boundary,
        'Content-Length': fullBody.length
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    });
    req.on('error', reject);
    req.write(fullBody);
    req.end();
  });
}

function getCuotasFechas(p) {
  const esDiario    = p.frecuencia === 'diario';
  const esSemanal   = p.frecuencia === 'semanal';
  const esQuincenal = p.frecuencia === 'quincenal';
  const esMensual   = p.frecuencia === 'mensual';
  const totalCuotas = Math.round(p.total / p.cuota);
  const fechas = [];
  const festivos = global._festivos || [];

  function esDiaInhabil(d) {
    if (d.getDay() === 0) return true;
    return festivos.includes(d.toISOString().split('T')[0]);
  }

  if (esDiario || esSemanal) {
    let offsetExtra = 0;
    for (let i = 0; i < totalCuotas; i++) {
      const d = new Date(p.fecha + 'T12:00:00');
      if (esSemanal) {
        d.setDate(d.getDate() + (i + 1) * 7 + offsetExtra);
      } else {
        d.setDate(d.getDate() + (i + 1) + offsetExtra);
      }
      if (esDiaInhabil(d)) {
        d.setDate(d.getDate() + 1);
        offsetExtra += 1;
        while (esDiaInhabil(d)) {
          d.setDate(d.getDate() + 1);
          offsetExtra += 1;
        }
      }
      fechas.push({ num: i + 1, fecha: d.toISOString().split('T')[0] });
    }
  } else {
    for (let i = 0; i < totalCuotas; i++) {
      const d = new Date(p.fecha + 'T12:00:00');
      if (esMensual) d.setMonth(d.getMonth() + (i + 1));
      else if (esQuincenal) d.setDate(d.getDate() + (i + 1) * 15);
      else d.setDate(d.getDate() + (i + 1));
      while (esDiaInhabil(d)) d.setDate(d.getDate() + 1);
      fechas.push({ num: i + 1, fecha: d.toISOString().split('T')[0] });
    }
  }
  return fechas;
}

function getSaldo(p, pagos) {
  const pagado = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento)
    .reduce((a, pg) => a + (pg.monto || 0), 0);
  return (p.total || p.monto || 0) - pagado;
}

// Pagado real en efectivo (excluye pago de cierre por renovación)
function getPagadoReal(p, pagos) {
  return pagos
    .filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento && !pg.esRenovacion)
    .reduce((a, pg) => a + (pg.monto || 0), 0);
}

function fmtQ(n) { return 'Q ' + (n || 0).toLocaleString('es-GT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function generarHTML(prestamos, pagos, hoy, deudas, gastos, papeleria) {
  const fechaLabel = new Date(hoy + 'T12:00:00').toLocaleDateString('es-GT', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });

  const deudaAcumulada = (deudas || []).reduce((a, d) => a + (d.monto || 0), 0);

  // Posición de Caja Real (igual que la app)
  const cuotaBanco = (deudas || []).reduce((a, d) => a + (d.cuota || d.monto || 0), 0);
  const efectivoRenov = prestamos.filter(p => p.esRenovacion).reduce((a, p) => a + (p.efectivoEntregado || 0), 0);
  const papeleriaRenov = (papeleria || []).filter(p => p.notas && p.notas.includes('Renovación')).reduce((a, p) => a + (p.monto || 0), 0);
  const totalGastos = (gastos || []).reduce((a, g) => a + (g.monto || 0), 0);

  let totalMonto = 0, totalPagado = 0, totalSaldo = 0, countAlDia = 0, countAtrasado = 0, countPagado = 0;
  const filas = prestamos.map((p, i) => {
    const saldo = getSaldo(p, pagos);
    // Pagado real: excluye pago sintético de cierre por renovación (igual que la app)
    const pagado = getPagadoReal(p, pagos);
    const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
    const fechasCuotas = getCuotasFechas(p);
    const cuotasVencidas = fechasCuotas.filter((fc, idx) => fc.fecha < hoy && idx >= pagosRealizados);

    let estado = 'Al día';
    let diasAtraso = 0;
    if (saldo <= 0) { estado = 'Pagado'; countPagado++; }
    else if (cuotasVencidas.length > 0) {
      estado = 'Atrasado';
      diasAtraso = Math.floor((new Date(hoy + 'T12:00:00') - new Date(cuotasVencidas[0].fecha + 'T12:00:00')) / 86400000);
      countAtrasado++;
    } else { countAlDia++; }

    const vence = fechasCuotas[fechasCuotas.length - 1]?.fecha || '—';
    // Solo suma monto en préstamos activos (igual que la app — excluye pagados/renovados)
    if (saldo > 0) totalMonto += p.monto || 0;
    totalPagado += pagado;
    totalSaldo += Math.max(0, saldo);

    const estadoColor = estado === 'Pagado' ? '#27ae60' : estado === 'Atrasado' ? '#e74c3c' : '#2980b9';
    return `<tr>
      <td>${i + 1}</td>
      <td><strong>${p.nombre}</strong></td>
      <td>${p.fecha}</td>
      <td style="color:${estado === 'Atrasado' ? '#e74c3c' : '#555'}">${vence}</td>
      <td style="color:${diasAtraso > 0 ? '#e74c3c' : '#555'}">${diasAtraso > 0 ? diasAtraso + 'd' : '—'}</td>
      <td>${fmtQ(p.monto)}</td>
      <td>${fmtQ(p.total)}</td>
      <td style="color:#27ae60">${fmtQ(pagado)}</td>
      <td style="color:#e74c3c">${fmtQ(saldo)}</td>
      <td>0.00</td>
      <td><span style="background:${estadoColor};color:#fff;padding:2px 8px;border-radius:4px;font-size:11px">${estado}</span></td>
    </tr>`;
  }).join('');

  const atrasadosList = prestamos.filter((p) => {
    const saldo = getSaldo(p, pagos);
    if (saldo <= 0) return false;
    const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
    const fechasCuotas = getCuotasFechas(p);
    const cuotasVencidas = fechasCuotas.filter((fc, idx) => fc.fecha < hoy && idx >= pagosRealizados);
    return cuotasVencidas.length > 0;
  });

  const filasAtrasados = atrasadosList.map((p, i) => {
    const saldo = getSaldo(p, pagos);
    const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
    const fechasCuotas = getCuotasFechas(p);
    const cuotasVencidas = fechasCuotas.filter((fc, idx) => fc.fecha < hoy && idx >= pagosRealizados);
    const fechaAntigua = cuotasVencidas[0].fecha;
    const dias = Math.floor((new Date(hoy + 'T12:00:00') - new Date(fechaAntigua + 'T12:00:00')) / 86400000);
    return `<tr>
      <td>${i + 1}</td>
      <td>${p.nombre}</td>
      <td>${p.telefono || '—'}</td>
      <td style="color:#e74c3c">${fechaAntigua}</td>
      <td style="color:#e74c3c">${dias} día${dias !== 1 ? 's' : ''}</td>
      <td style="color:#e74c3c">${fmtQ(saldo)}</td>
      <td>Q 0.00</td>
    </tr>`;
  }).join('');

  const posicionCajaReal = totalPagado - cuotaBanco - efectivoRenov - papeleriaRenov - totalGastos;
  const cajaColor = posicionCajaReal >= 0 ? '#27ae60' : '#e74c3c';
  const cajaDesc = [
    `Recuperado ${fmtQ(totalPagado)}`,
    cuotaBanco > 0 ? `Banco ${fmtQ(cuotaBanco)}` : '',
    efectivoRenov > 0 ? `Entregado renovac. ${fmtQ(efectivoRenov)}` : '',
    papeleriaRenov > 0 ? `Papelería renovac. ${fmtQ(papeleriaRenov)}` : '',
    totalGastos > 0 ? `Gastos ${fmtQ(totalGastos)}` : '',
  ].filter(Boolean).join(' – ');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; font-size: 12px; color: #333; margin: 20px; }
    .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #c8102e; padding-bottom: 10px; margin-bottom: 16px; }
    .logo { font-size: 22px; font-weight: 900; color: #c8102e; }
    .logo span { color: #333; }
    .report-info { text-align: right; font-size: 11px; color: #666; }
    .kpis { display: flex; gap: 10px; margin-bottom: 16px; }
    .kpi { flex: 1; border: 1px solid #ddd; border-radius: 6px; padding: 10px; }
    .kpi-label { font-size: 10px; color: #888; text-transform: uppercase; }
    .kpi-value { font-size: 15px; font-weight: 700; margin-top: 4px; }
    .badges { margin-bottom: 12px; }
    .badge { display: inline-block; padding: 3px 10px; border-radius: 4px; font-size: 11px; margin-right: 6px; }
    h3 { font-size: 13px; margin: 16px 0 8px; color: #333; border-left: 3px solid #c8102e; padding-left: 8px; }
    table { width: 100%; border-collapse: collapse; font-size: 11px; }
    th { background: #222; color: #fff; padding: 6px 8px; text-align: left; }
    td { padding: 5px 8px; border-bottom: 1px solid #eee; }
    tr:nth-child(even) { background: #f9f9f9; }
    .totals td { font-weight: 700; background: #f0f0f0; border-top: 2px solid #ccc; }
  </style></head><body>
  <div class="header">
    <div class="logo"><span>CREDIT</span>X<br><small style="font-size:11px;font-weight:400;color:#666">Soluciones Financieras</small></div>
    <div class="report-info">
      <strong>REPORTE DE PRÉSTAMOS</strong><br>
      Generado: ${fechaLabel}<br>
      Total registros: ${prestamos.length}
    </div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="kpi-label">Total Prestado</div><div class="kpi-value" style="color:#c8102e">${fmtQ(totalMonto)}</div></div>
    <div class="kpi"><div class="kpi-label">Total Pagado</div><div class="kpi-value" style="color:#27ae60">${fmtQ(totalPagado)}</div></div>
    <div class="kpi"><div class="kpi-label">🏦 Deuda Banco</div><div class="kpi-value" style="color:#e74c3c">- ${fmtQ(deudaAcumulada)}</div></div>
    <div class="kpi"><div class="kpi-label">💲 Posición de Caja Real</div><div class="kpi-value" style="color:${cajaColor}">${fmtQ(posicionCajaReal)}</div><div style="font-size:8px;color:#888;margin-top:4px;">${cajaDesc}</div></div>
    <div class="kpi"><div class="kpi-label">Saldo Pendiente</div><div class="kpi-value" style="color:#e74c3c">${fmtQ(totalSaldo)}</div></div>
  </div>
  ${deudas && deudas.length > 0 ? `
  <h3>🏦 Historial de Deudas con Banco</h3>
  <table>
    <thead><tr><th>#</th><th>Fecha</th><th>Hora</th><th>Monto Deuda</th><th>Disponible</th></tr></thead>
    <tbody>${deudas.map((d, i) => `<tr>
      <td>${i + 1}</td>
      <td>${d.fecha || '—'}</td>
      <td>${d.hora || '—'}</td>
      <td style="color:#e74c3c">- ${fmtQ(d.monto)}</td>
      <td style="color:${(d.disponible || 0) >= 0 ? '#27ae60' : '#e74c3c'}">${fmtQ(d.disponible)}</td>
    </tr>`).join('')}</tbody>
    <tfoot><tr class="totals"><td colspan="3">TOTAL DEUDA BANCO →</td><td style="color:#e74c3c">- ${fmtQ(deudaAcumulada)}</td><td style="color:${cajaColor}">${fmtQ(posicionCajaReal)}</td></tr></tfoot>
  </table>` : ''}
  <h3>Detalle de Préstamos</h3>
  <table>
    <thead><tr><th>#</th><th>Cliente</th><th>Fecha Préstamo</th><th>Vencimiento</th><th>Días Atraso</th><th>Monto</th><th>Total</th><th>Pagado</th><th>Saldo</th><th>Mora</th><th>Estado</th></tr></thead>
    <tbody>${filas}</tbody>
    <tfoot><tr class="totals"><td colspan="5">TOTALES →</td><td>${fmtQ(totalMonto)}</td><td></td><td style="color:#27ae60">${fmtQ(totalPagado)}</td><td style="color:#e74c3c">${fmtQ(totalSaldo)}</td><td>Q 0.00</td><td></td></tr></tfoot>
  </table>
  ${atrasadosList.length > 0 ? `
  <h3>⚠ Detalle de Clientes con Atraso</h3>
  <table>
    <thead><tr><th>#</th><th>Cliente</th><th>Teléfono</th><th>Venció</th><th>Días Atraso</th><th>Saldo</th><th>Mora</th></tr></thead>
    <tbody>${filasAtrasados}</tbody>
    <tfoot><tr class="totals"><td colspan="5">TOTAL ATRASOS →</td><td style="color:#e74c3c">${fmtQ(atrasadosList.reduce((a, p) => a + getSaldo(p, pagos), 0))}</td><td>Q 0.00</td></tr></tfoot>
  </table>` : ''}
  </body></html>`;
}

async function generarExcel(prestamos, pagos, hoy) {
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CreditX';

  // Clasificar préstamos
  const activos = [], renovaciones = [], cerrados = [];
  for (const p of prestamos) {
    const saldo = getSaldo(p, pagos);
    const pagado = getPagadoReal(p, pagos);
    const fechasCuotas = getCuotasFechas(p);
    const vence = fechasCuotas[fechasCuotas.length - 1]?.fecha || '—';
    const row = { ...p, saldo, pagado, vence, totalCalc: p.total || 0 };
    if (saldo <= 0) cerrados.push(row);
    else if (p.fecha >= '2026-04-01') renovaciones.push(row);
    else activos.push(row);
  }

  const ws = wb.addWorksheet('Cartera Completa', { properties: { tabColor: { argb: '2E4057' } } });

  // Title
  ws.mergeCells('A1:L1');
  const t = ws.getCell('A1');
  t.value = 'CREDITX — REPORTE FINANCIERO DE CARTERA';
  t.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1B3A4B' } };
  t.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  ws.mergeCells('A2:L2');
  const sub = ws.getCell('A2');
  sub.value = `Generado: ${hoy} | Incluye préstamos activos y renovaciones`;
  sub.font = { size: 9, italic: true, color: { argb: '666666' } };
  sub.alignment = { horizontal: 'center' };

  // Headers
  const headers = ['#','Cliente','Categoría','Fecha Préstamo','Vencimiento','Días Atraso','Monto Prestado','Total c/Interés','Total Pagado','Saldo Pendiente','Mora','Estado'];
  const hdrRow = ws.addRow(headers);
  hdrRow.number; // row 3
  // Add blank row 3, headers at row 4
  ws.spliceRows(3, 0, []);
  const hr = ws.getRow(4);
  headers.forEach((h, i) => {
    const c = hr.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: 'FFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E4057' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });
  ws.getRow(4).height = 26;

  const fills = {
    marzo: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E8F5E9' } },
    renov: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E0' } },
    cerrado: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E3F2FD' } },
    marzoHdr: { type: 'pattern', pattern: 'solid', fgColor: { argb: '43A047' } },
    renovHdr: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FB8C00' } },
    cerradoHdr: { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E88E5' } },
  };
  const thinBorder = { top: { style: 'thin', color: { argb: 'AAAAAA' } }, bottom: { style: 'thin', color: { argb: 'AAAAAA' } }, left: { style: 'thin', color: { argb: 'AAAAAA' } }, right: { style: 'thin', color: { argb: 'AAAAAA' } } };

  let num = 0;
  function addSection(label, items, catLabel, fillRow, fillHdr) {
    const secRow = ws.addRow([]);
    ws.mergeCells(secRow.number, 1, secRow.number, 12);
    const sc = secRow.getCell(1);
    sc.value = `  ● ${label}`;
    sc.font = { bold: true, size: 11, color: { argb: 'FFFFFF' } };
    sc.fill = fillHdr;
    const startRow = secRow.number + 1;
    for (const p of items) {
      num++;
      const saldo = p.saldo;
      const estado = saldo <= 0 ? 'Pagado' : 'Al día';
      const r = ws.addRow([num, p.nombre, catLabel, p.fecha, p.vence, 0, p.monto, p.totalCalc, p.pagado, Math.max(0, saldo), 0, estado]);
      for (let c = 1; c <= 12; c++) {
        const cell = r.getCell(c);
        cell.fill = fillRow;
        cell.border = thinBorder;
        cell.font = { name: 'Arial', size: 10 };
        if ([7,8,9,10,11].includes(c)) cell.numFmt = '#,##0.00';
        if ([1,6,12].includes(c)) cell.alignment = { horizontal: 'center' };
      }
    }
    return { startRow, endRow: ws.lastRow.number };
  }

  const s1 = addSection('PRÉSTAMOS ACTIVOS — ORIGINADOS EN MARZO', activos, 'Activo Marzo', fills.marzo, fills.marzoHdr);
  const s2 = addSection('PRÉSTAMOS ACTIVOS — RENOVACIONES ABRIL', renovaciones, 'Renovación', fills.renov, fills.renovHdr);
  const s3 = addSection('PRÉSTAMOS CERRADOS POR RENOVACIÓN', cerrados, 'Cerrado x Renov.', fills.cerrado, fills.cerradoHdr);

  // Totals
  ws.addRow([]);
  const totHdrRow = ws.addRow([]);
  ws.mergeCells(totHdrRow.number, 1, totHdrRow.number, 12);
  const th = totHdrRow.getCell(1);
  th.value = '  TOTALES GENERALES';
  th.font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  th.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1B3A4B' } };

  const grayFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D5D5D5' } };
  const sections = [
    { label: `Activos Marzo (${activos.length})`, ...s1, fill: fills.marzo },
    { label: `Renovaciones Abril (${renovaciones.length})`, ...s2, fill: fills.renov },
    { label: `Cerrados x Renov. (${cerrados.length})`, ...s3, fill: fills.cerrado },
  ];
  const sumRows = [];
  for (const sec of sections) {
    const r = ws.addRow([]);
    r.getCell(2).value = sec.label;
    r.getCell(2).font = { bold: true, name: 'Arial', size: 10 };
    r.getCell(2).fill = sec.fill;
    for (const col of [7,8,9,10,11]) {
      const letter = String.fromCharCode(64 + col);
      const c = r.getCell(col);
      c.value = { formula: `SUM(${letter}${sec.startRow}:${letter}${sec.endRow})` };
      c.numFmt = '#,##0.00';
      c.font = { bold: true, name: 'Arial', size: 10 };
      c.border = thinBorder;
      c.fill = sec.fill;
    }
    sumRows.push(r.number);
  }
  // Gran total
  const gt = ws.addRow([]);
  gt.getCell(2).value = 'GRAN TOTAL';
  gt.getCell(2).font = { bold: true, name: 'Arial', size: 11 };
  for (let col = 1; col <= 12; col++) {
    gt.getCell(col).fill = grayFill;
    gt.getCell(col).border = { top: { style: 'medium' }, bottom: { style: 'medium' }, left: { style: 'thin' }, right: { style: 'thin' } };
  }
  for (const col of [7,8,9,10,11]) {
    const letter = String.fromCharCode(64 + col);
    const c = gt.getCell(col);
    c.value = { formula: `${letter}${sumRows[0]}+${letter}${sumRows[1]}+${letter}${sumRows[2]}` };
    c.numFmt = '#,##0.00';
    c.font = { bold: true, name: 'Arial', size: 11 };
  }
  const gtRow = gt.number;

  // Rentabilidad section
  ws.addRow([]);
  const rentHdr = ws.addRow([]);
  ws.mergeCells(rentHdr.number, 1, rentHdr.number, 12);
  const rh = rentHdr.getCell(1);
  rh.value = '  ANÁLISIS DE RENTABILIDAD';
  rh.font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  rh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1B3A4B' } };

  const metrics = [
    ['Capital Total Colocado', `G${gtRow}`],
    ['Interés Total Generado', `H${gtRow}-G${gtRow}`],
    ['Tasa de Interés Promedio', `(H${gtRow}-G${gtRow})/G${gtRow}`],
    ['Total Cobrado', `I${gtRow}`],
    ['Saldo por Cobrar', `J${gtRow}`],
    ['% Recuperación', `I${gtRow}/H${gtRow}`],
  ];
  for (const [lbl, formula] of metrics) {
    const r = ws.addRow([]);
    ws.mergeCells(r.number, 2, r.number, 4);
    ws.mergeCells(r.number, 5, r.number, 7);
    r.getCell(2).value = lbl;
    r.getCell(2).font = { bold: true, name: 'Arial', size: 10 };
    r.getCell(2).border = thinBorder;
    const c = r.getCell(5);
    c.value = { formula };
    c.font = { bold: true, name: 'Arial', size: 11, color: { argb: '1B3A4B' } };
    c.border = thinBorder;
    c.numFmt = lbl.includes('%') || lbl.includes('Tasa') ? '0.0%' : '#,##0.00';
  }

  // Leyenda
  ws.addRow([]);
  const legRow = ws.addRow([]);
  legRow.getCell(2).value = 'LEYENDA DE COLORES:';
  legRow.getCell(2).font = { bold: true, name: 'Arial', size: 10 };
  for (const [lbl, fill] of [['Verde — Activos originados en Marzo', fills.marzo], ['Naranja — Renovaciones de Abril', fills.renov], ['Azul — Cerrados por Renovación', fills.cerrado]]) {
    const r = ws.addRow([]);
    ws.mergeCells(r.number, 2, r.number, 5);
    r.getCell(2).value = lbl;
    r.getCell(2).font = { name: 'Arial', size: 9 };
    r.getCell(2).fill = fill;
  }

  // Column widths
  [4,22,16,15,14,12,16,16,15,16,10,10].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ========== HOJA 2: DETALLE RENOVACIONES ==========
  const wr = wb.addWorksheet('Detalle Renovaciones', { properties: { tabColor: { argb: 'FB8C00' } } });
  wr.mergeCells('A1:H1');
  const wrTitle = wr.getCell('A1');
  wrTitle.value = 'DETALLE DE RENOVACIONES';
  wrTitle.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
  wrTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E65100' } };
  wrTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  wr.getRow(1).height = 32;

  const wrHeaders = ['Cliente','Préstamo Anterior','Monto Anterior','Pagado Anterior','Nuevo Préstamo','Monto Nuevo','Total Nuevo','Fecha Renovación'];
  const wrHr = wr.addRow(wrHeaders);
  wrHr.eachCell((c) => {
    c.font = { bold: true, size: 10, color: { argb: 'FFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2E4057' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border = thinBorder;
  });

  // Match renovaciones with their cerrados by nombre
  for (const ren of renovaciones) {
    const cerrado = cerrados.find(c => c.nombre.toLowerCase() === ren.nombre.toLowerCase());
    const r = wr.addRow([
      ren.nombre,
      cerrado ? `${cerrado.fecha} → ${cerrado.vence}` : '—',
      cerrado ? cerrado.monto : 0,
      cerrado ? cerrado.pagado : 0,
      `${ren.fecha} → ${ren.vence}`,
      ren.monto,
      ren.totalCalc,
      ren.fecha
    ]);
    r.eachCell((c, ci) => {
      c.font = { name: 'Arial', size: 10 };
      c.border = thinBorder;
      c.fill = fills.renov;
      if ([3,4,6,7].includes(ci)) c.numFmt = '#,##0.00';
    });
  }

  // Totals row
  const wrDataStart = 3;
  const wrDataEnd = wrDataStart + renovaciones.length - 1;
  const wrTot = wr.addRow(['TOTALES']);
  wrTot.getCell(1).font = { bold: true, name: 'Arial', size: 10 };
  for (const [ci, letter] of [[3,'C'],[4,'D'],[6,'F'],[7,'G']]) {
    const c = wrTot.getCell(ci);
    c.value = { formula: `SUM(${letter}${wrDataStart}:${letter}${wrDataEnd})` };
    c.numFmt = '#,##0.00';
    c.font = { bold: true, name: 'Arial', size: 10 };
    c.border = { top: { style: 'medium' }, bottom: { style: 'medium' }, left: { style: 'thin' }, right: { style: 'thin' } };
  }
  [22,24,16,16,24,16,16,16].forEach((w, i) => { wr.getColumn(i + 1).width = w; });

  // ========== HOJA 3: RESUMEN EJECUTIVO ==========
  const rs = wb.addWorksheet('Resumen Ejecutivo', { properties: { tabColor: { argb: '1B3A4B' } } });
  rs.mergeCells('A1:D1');
  const rsTitle = rs.getCell('A1');
  rsTitle.value = 'RESUMEN EJECUTIVO — CREDITX';
  rsTitle.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
  rsTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1B3A4B' } };
  rsTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  rs.getRow(1).height = 35;

  const sectionFont = { bold: true, size: 12, color: { argb: '1B3A4B' } };
  const valFont = { bold: true, name: 'Arial', size: 11 };
  const lblFont = { name: 'Arial', size: 10 };
  const botBorder = { bottom: { style: 'medium' } };

  // Cartera Activa
  let rr = 3;
  rs.mergeCells(`A${rr}:D${rr}`);
  rs.getCell(`A${rr}`).value = 'CARTERA ACTIVA'; rs.getCell(`A${rr}`).font = sectionFont; rs.getCell(`A${rr}`).border = botBorder;
  rr++;
  const activosMonto = activos.reduce((a, p) => a + (p.monto || 0), 0);
  const renovMonto = renovaciones.reduce((a, p) => a + (p.monto || 0), 0);
  rs.getCell(`A${rr}`).value = 'Préstamos activos totales'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = activosMonto + renovMonto; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rr++;
  rs.getCell(`A${rr}`).value = 'Préstamos activos marzo'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = activosMonto; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rs.getCell(`C${rr}`).value = activos.length; rs.getCell(`D${rr}`).value = 'préstamos'; rs.getCell(`D${rr}`).font = lblFont;
  rr++;
  rs.getCell(`A${rr}`).value = 'Préstamos renovaciones abril'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = renovMonto; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rs.getCell(`C${rr}`).value = renovaciones.length; rs.getCell(`D${rr}`).value = 'préstamos'; rs.getCell(`D${rr}`).font = lblFont;
  rr += 2;

  // Cobranza
  rs.mergeCells(`A${rr}:D${rr}`);
  rs.getCell(`A${rr}`).value = 'COBRANZA'; rs.getCell(`A${rr}`).font = sectionFont; rs.getCell(`A${rr}`).border = botBorder;
  rr++;
  const totalCobrado = prestamos.reduce((a, p) => a + getPagadoReal(p, pagos), 0);
  const totalSaldoPend = prestamos.reduce((a, p) => a + Math.max(0, getSaldo(p, pagos)), 0);
  rs.getCell(`A${rr}`).value = 'Total cobrado a la fecha'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = totalCobrado; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rr++;
  rs.getCell(`A${rr}`).value = 'Saldo pendiente por cobrar'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = totalSaldoPend; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rr += 2;

  // Cartera Cerrada
  rs.mergeCells(`A${rr}:D${rr}`);
  rs.getCell(`A${rr}`).value = 'CARTERA CERRADA'; rs.getCell(`A${rr}`).font = sectionFont; rs.getCell(`A${rr}`).border = botBorder;
  rr++;
  const cerradosMonto = cerrados.reduce((a, p) => a + (p.monto || 0), 0);
  const cerradosCobrado = cerrados.reduce((a, p) => a + (p.pagado || 0), 0);
  rs.getCell(`A${rr}`).value = 'Préstamos cerrados por renovación'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = cerradosMonto; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;
  rs.getCell(`C${rr}`).value = cerrados.length; rs.getCell(`D${rr}`).value = 'préstamos'; rs.getCell(`D${rr}`).font = lblFont;
  rr++;
  rs.getCell(`A${rr}`).value = 'Cobrado de préstamos cerrados'; rs.getCell(`A${rr}`).font = lblFont;
  rs.getCell(`B${rr}`).value = cerradosCobrado; rs.getCell(`B${rr}`).numFmt = '#,##0.00'; rs.getCell(`B${rr}`).font = valFont;

  rs.getColumn(1).width = 34; rs.getColumn(2).width = 18; rs.getColumn(3).width = 8; rs.getColumn(4).width = 12;

  const xlsxPath = '/tmp/reporte_creditx.xlsx';
  await wb.xlsx.writeFile(xlsxPath);
  return xlsxPath;
}

(async () => {
  try {
    const snap = await db.collection('datos').doc('principal').get();
    if (!snap.exists) { await send('Sin datos en CreditX'); process.exit(0); return; }

    const data = snap.data();
    const prestamos = data.prestamos || [];
    const pagos = data.pagos || [];
    const deudas = data.deudas || [];
    const gastos = data.gastos || [];
    const papeleria = data.papeleria || [];
    const festivos = data.festivos || [];

    // Set global for getCuotasFechas
    global._festivos = festivos;
    const hoy = getDateStr(0);
    const manana = getDateStr(1);

    console.log('HOY:', hoy);
    console.log('FESTIVOS en Firestore:', JSON.stringify(festivos));
    console.log('Total festivos:', festivos.length);

    const fechaLabel = new Date(hoy + 'T12:00:00').toLocaleDateString('es-GT', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });

    // Generar PDF con Puppeteer
    const puppeteer = require('puppeteer');
    const htmlContent = generarHTML(prestamos, pagos, hoy, deudas, gastos, papeleria);
    const htmlPath = '/tmp/reporte.html';
    const pdfPath = '/tmp/reporte.pdf';
    fs.writeFileSync(htmlPath, htmlContent);

    const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    const page = await browser.newPage();
    await page.goto('file://' + htmlPath, { waitUntil: 'networkidle0' });
    await page.pdf({ path: pdfPath, format: 'Letter', margin: { top: '10mm', bottom: '10mm', left: '10mm', right: '10mm' } });
    await browser.close();

    // Enviar PDF por Telegram
    await sendDocument(pdfPath, 'Reporte CreditX — ' + fechaLabel);
    console.log('PDF enviado OK');

    // Generar y enviar Excel
    const xlsxPath = await generarExcel(prestamos, pagos, hoy);
    await sendDocument(xlsxPath, 'Excel CreditX — ' + fechaLabel);
    console.log('Excel enviado OK');

    // También enviar resumen de cobros del día
    const cobrosHoy = [], cobrosManana = [], atrasados = [];
    for (const p of prestamos) {
      const saldo = getSaldo(p, pagos);
      if (saldo <= 0) continue;
      const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
      const fechasCuotas = getCuotasFechas(p);
      const proximaCuota = fechasCuotas[pagosRealizados];
      const cuotasVencidas = fechasCuotas.filter((fc, idx) => fc.fecha < hoy && idx >= pagosRealizados);
      if (!proximaCuota) continue;
      const info = p.nombre + (p.telefono ? ' | ' + p.telefono : '') + '\n   Q ' + (p.cuota || 0).toFixed(2) + ' | Saldo: ' + fmtQ(saldo);
      if (cuotasVencidas.length > 0) atrasados.push(p.nombre + ' | Saldo: ' + fmtQ(saldo));
      else if (proximaCuota.fecha === hoy) cobrosHoy.push(info);
      else if (proximaCuota.fecha === manana) cobrosManana.push(info);
    }

    if (cobrosHoy.length || cobrosManana.length || atrasados.length) {
      let msg = 'RESUMEN DEL DIA - CreditX\n' + fechaLabel + '\n\n';
      if (cobrosHoy.length) {
        const total = cobrosHoy.reduce((a, info) => { const m = info.match(/Q ([\d,.]+) \|/); return a + (m ? parseFloat(m[1].replace(',','')) : 0); }, 0);
        msg += 'COBRAR HOY (' + cobrosHoy.length + ')\n';
        cobrosHoy.forEach((info, i) => { msg += (i + 1) + '. ' + info + '\n'; });
        msg += 'Total hoy: Q ' + total.toFixed(2) + '\n\n';
      }
      if (cobrosManana.length) {
        const total = cobrosManana.reduce((a, info) => { const m = info.match(/Q ([\d,.]+) \|/); return a + (m ? parseFloat(m[1].replace(',','')) : 0); }, 0);
        msg += 'COBRAR MANANA (' + cobrosManana.length + ')\n';
        cobrosManana.forEach((info, i) => { msg += (i + 1) + '. ' + info + '\n'; });
        msg += 'Total manana: Q ' + total.toFixed(2) + '\n\n';
      }
      if (atrasados.length) {
        msg += 'ATRASADOS (' + atrasados.length + ')\n';
        atrasados.forEach((info, i) => { msg += (i + 1) + '. ' + info + '\n'; });
      }
      msg += '\nCreditX Soluciones Financieras';
      await send(msg);
    }

  } catch (e) {
    console.error(e);
    await send('Error CreditX: ' + e.message);
  }
  process.exit(0);
})();
