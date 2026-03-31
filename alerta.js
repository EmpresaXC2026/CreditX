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
      'Content-Type: application/pdf\r\n\r\n'
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
  if (esDiario || esSemanal) {
    let offsetExtra = 0;
    for (let i = 0; i < totalCuotas; i++) {
      const d = new Date(p.fecha + 'T12:00:00');
      if (esSemanal) d.setDate(d.getDate() + (i + 1) * 7 + offsetExtra);
      else d.setDate(d.getDate() + (i + 1) + offsetExtra);
      if (d.getDay() === 0) { d.setDate(d.getDate() + 1); offsetExtra++; }
      fechas.push({ num: i + 1, fecha: d.toISOString().split('T')[0] });
    }
  } else {
    for (let i = 0; i < totalCuotas; i++) {
      const d = new Date(p.fecha + 'T12:00:00');
      if (esMensual) d.setMonth(d.getMonth() + (i + 1));
      else if (esQuincenal) d.setDate(d.getDate() + (i + 1) * 15);
      else d.setDate(d.getDate() + (i + 1));
      if (d.getDay() === 0) d.setDate(d.getDate() + 1);
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

function fmtQ(n) { return 'Q ' + (n || 0).toLocaleString('es-GT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function generarHTML(prestamos, pagos, hoy, deudas) {
  const fechaLabel = new Date(hoy + 'T12:00:00').toLocaleDateString('es-GT', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });

  const deudaAcumulada = (deudas || []).reduce((a, d) => a + (d.monto || 0), 0);

  let totalMonto = 0, totalPagado = 0, totalSaldo = 0, countAlDia = 0, countAtrasado = 0, countPagado = 0;
  const filas = prestamos.map((p, i) => {
    const saldo = getSaldo(p, pagos);
    const pagado = (p.total || 0) - saldo;
    const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
    const fechasCuotas = getCuotasFechas(p);
    const proximaCuota = fechasCuotas[pagosRealizados];

    let estado = 'Al día';
    let diasAtraso = 0;
    if (saldo <= 0) { estado = 'Pagado'; countPagado++; }
    else if (proximaCuota && proximaCuota.fecha < hoy) {
      estado = 'Atrasado';
      diasAtraso = Math.floor((new Date(hoy) - new Date(proximaCuota.fecha)) / 86400000);
      countAtrasado++;
    } else { countAlDia++; }

    const vence = fechasCuotas[fechasCuotas.length - 1]?.fecha || '—';
    totalMonto += p.monto || 0;
    totalPagado += pagado;
    totalSaldo += saldo;

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
    const proximaCuota = fechasCuotas[pagosRealizados];
    return proximaCuota && proximaCuota.fecha < hoy;
  });

  const filasAtrasados = atrasadosList.map((p, i) => {
    const saldo = getSaldo(p, pagos);
    const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
    const fechasCuotas = getCuotasFechas(p);
    const proximaCuota = fechasCuotas[pagosRealizados];
    const dias = Math.floor((new Date(hoy) - new Date(proximaCuota.fecha)) / 86400000);
    return `<tr>
      <td>${i + 1}</td>
      <td>${p.nombre}</td>
      <td>${p.telefono || '—'}</td>
      <td style="color:#e74c3c">${proximaCuota.fecha}</td>
      <td style="color:#e74c3c">${dias} día${dias !== 1 ? 's' : ''}</td>
      <td style="color:#e74c3c">${fmtQ(saldo)}</td>
      <td>Q 0.00</td>
    </tr>`;
  }).join('');

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
    <div class="kpi"><div class="kpi-label">✅ Disponible</div><div class="kpi-value" style="color:${(totalPagado - deudaAcumulada) >= 0 ? '#27ae60' : '#e74c3c'}">${fmtQ(totalPagado - deudaAcumulada)}</div></div>
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
    <tfoot><tr class="totals"><td colspan="3">TOTAL DEUDA BANCO →</td><td style="color:#e74c3c">- ${fmtQ(deudaAcumulada)}</td><td style="color:${(totalPagado - deudaAcumulada) >= 0 ? '#27ae60' : '#e74c3c'}">${fmtQ(totalPagado - deudaAcumulada)}</td></tr></tfoot>
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

(async () => {
  try {
    const snap = await db.collection('datos').doc('principal').get();
    if (!snap.exists) { await send('Sin datos en CreditX'); process.exit(0); return; }

    const data = snap.data();
    const prestamos = data.prestamos || [];
    const pagos = data.pagos || [];
    const deudas = data.deudas || [];
    const hoy = getDateStr(0);
    const manana = getDateStr(1);

    const fechaLabel = new Date(hoy + 'T12:00:00').toLocaleDateString('es-GT', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });

    // Generar PDF con Puppeteer
    const puppeteer = require('puppeteer');
    const htmlContent = generarHTML(prestamos, pagos, hoy, deudas);
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

    // También enviar resumen de cobros del día
    const cobrosHoy = [], cobrosManana = [], atrasados = [];
    for (const p of prestamos) {
      const saldo = getSaldo(p, pagos);
      if (saldo <= 0) continue;
      const pagosRealizados = pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
      const fechasCuotas = getCuotasFechas(p);
      const proximaCuota = fechasCuotas[pagosRealizados];
      if (!proximaCuota) continue;
      const info = p.nombre + (p.telefono ? ' | ' + p.telefono : '') + '\n   Q ' + (p.cuota || 0).toFixed(2) + ' | Saldo: ' + fmtQ(saldo);
      if (proximaCuota.fecha === hoy) cobrosHoy.push(info);
      else if (proximaCuota.fecha === manana) cobrosManana.push(info);
      else if (proximaCuota.fecha < hoy) atrasados.push(p.nombre + ' | Saldo: ' + fmtQ(saldo));
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
