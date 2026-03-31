const admin = require('firebase-admin');
const https = require('https');

const sa = JSON.parse(process.env.FIREBASE_KEY);
admin.initializeApp({ credential: admin.credential.cert(sa) });
const db = admin.firestore();

const token = process.env.TG_TOKEN;
const chatId = process.env.TG_CHAT_ID;

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

function getDateStr(offsetDays) {
  const d = new Date();
  d.setUTCHours(d.getUTCHours() - 6); // Guatemala UTC-6
  d.setUTCDate(d.getUTCDate() + offsetDays);
  return d.toISOString().split('T')[0];
}

// Replica exacta de getCuotasFechas del app
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
      if (d.getDay() === 0) { d.setDate(d.getDate() + 1); offsetExtra += 1; }
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
  const pagado = pagos
    .filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento)
    .reduce((a, pg) => a + (pg.monto || 0), 0);
  return (p.total || p.monto || 0) - pagado;
}

function getPagosRealizados(p, pagos) {
  return pagos.filter(pg => pg.prestamoId === p.id && !pg.soloIncumplimiento).length;
}

(async () => {
  try {
    const snap = await db.collection('datos').doc('principal').get();
    if (!snap.exists) { await send('Sin datos en CreditX'); return; }

    const data = snap.data();
    const prestamos = data.prestamos || [];
    const pagos = data.pagos || [];
    const hoy = getDateStr(0);
    const manana = getDateStr(1);

    const fechaLabel = new Date(hoy + 'T12:00:00').toLocaleDateString('es-GT', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });

    const cobrosHoy = [];
    const cobrosManana = [];
    const atrasados = [];

    for (const p of prestamos) {
      const saldo = getSaldo(p, pagos);
      if (saldo <= 0) continue;

      const pagosRealizados = getPagosRealizados(p, pagos);
      const fechasCuotas = getCuotasFechas(p);
      const proximaCuota = fechasCuotas[pagosRealizados];

      if (!proximaCuota) continue;

      const info = `${p.nombre}${p.telefono ? ' | ' + p.telefono : ''}\n   Q ${(p.cuota || 0).toFixed(2)} | Saldo: Q ${saldo.toFixed(2)}`;

      if (proximaCuota.fecha === hoy) {
        cobrosHoy.push(info);
      } else if (proximaCuota.fecha === manana) {
        cobrosManana.push(info);
      } else if (proximaCuota.fecha < hoy) {
        atrasados.push(`${p.nombre}${p.telefono ? ' | ' + p.telefono : ''}\n   Saldo: Q ${saldo.toFixed(2)}`);
      }
    }

    if (!cobrosHoy.length && !cobrosManana.length && !atrasados.length) {
      await send('CreditX - Sin cobros pendientes\n' + fechaLabel);
      return;
    }

    let msg = 'RESUMEN DEL DIA - CreditX\n' + fechaLabel + '\n\n';

    if (cobrosHoy.length) {
      const total = cobrosHoy.reduce((a, info) => {
        const match = info.match(/Q ([\d.]+) \|/);
        return a + (match ? parseFloat(match[1]) : 0);
      }, 0);
      msg += 'COBRAR HOY (' + cobrosHoy.length + ')\n';
      cobrosHoy.forEach((info, i) => { msg += (i + 1) + '. ' + info + '\n'; });
      msg += 'Total hoy: Q ' + total.toFixed(2) + '\n\n';
    }

    if (cobrosManana.length) {
      const total = cobrosManana.reduce((a, info) => {
        const match = info.match(/Q ([\d.]+) \|/);
        return a + (match ? parseFloat(match[1]) : 0);
      }, 0);
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
    console.log('Enviado OK');

  } catch (e) {
    console.error(e);
    await send('Error CreditX: ' + e.message);
  }
  process.exit(0);
})();
