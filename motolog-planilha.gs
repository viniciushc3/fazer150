/**
 * Motolog: API da planilha Google (Google Apps Script)
 *
 * Crie uma planilha nova em branco, abra Extensões > Apps Script nela e cole este código.
 * (Não cole na planilha do Central da Moto: ela já tem outro script.)
 * Troque o TOKEN abaixo por uma senha só sua (a mesma que você vai digitar no app).
 *
 * O script só lê e grava as abas que começam com "Motolog_". Ele nunca lê nem
 * devolve as outras abas (Usuarios, Página1, Manutencao, etc.).
 * As fotos de recibo vão para a pasta "Motolog Recibos" no seu Google Drive (só você vê).
 *
 * Resumo semanal por e-mail: depois de colar, escolha a função "instalarResumoSemanal"
 * no topo do editor e clique em Executar (uma vez só). Ele pede permissão para enviar
 * e-mail e passa a mandar o resumo toda segunda às 8h, se estiver ligado nos Ajustes do app.
 */
const TOKEN = 'TROQUE-ESTA-SENHA';

// Deixe vazio para usar a planilha onde o script foi colado.
// (Só preencha com um ID se o script estiver num projeto separado.)
const PLANILHA_ID = '';
function planilha_() { return PLANILHA_ID ? SpreadsheetApp.openById(PLANILHA_ID) : SpreadsheetApp.getActive(); }

const ABAS = {
  abastecimentos: { nome: 'Motolog_Abastecimentos', campos: ['id', 'data', 'hora', 'km', 'litros', 'valor', 'comb', 'posto', 'cheio', 'foto', 'gps'] },
  manutencoes:    { nome: 'Motolog_Manutencoes',    campos: ['id', 'data', 'hora', 'km', 'servico', 'outro', 'valor', 'oficina', 'obs', 'foto'] },
  gastos:         { nome: 'Motolog_Gastos',         campos: ['id', 'data', 'hora', 'valor', 'cat', 'desc', 'foto'] },
};
const ABA_CONFIG = 'Motolog_Config';
const PASTA_RECIBOS = 'Motolog Recibos';

function doGet(e) {
  if (!tokenOk_(e.parameter.token)) return saida_({ ok: false, erro: 'token' });
  return saida_({ ok: true, rev: rev_(), db: ler_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return saida_({ ok: false, erro: 'json' }); }
  if (!tokenOk_(body.token)) return saida_({ ok: false, erro: 'token' });
  if (body.acao === 'foto') return saida_(salvarFoto_(body));
  if (!body.db || !Array.isArray(body.db.abastecimentos) || !Array.isArray(body.db.manutencoes)) {
    return saida_({ ok: false, erro: 'formato' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const atual = rev_();
    // Evita que um aparelho desatualizado apague dados gravados por outro.
    if (body.baseRev !== 'force' && Number(body.baseRev) !== atual) {
      return saida_({ ok: false, erro: 'conflito', rev: atual, db: ler_() });
    }
    gravar_(body.db);
    const nova = atual + 1;
    setConfig_('rev', nova);
    setConfig_('atualizado', new Date().toISOString());
    return saida_({ ok: true, rev: nova });
  } finally {
    lock.releaseLock();
  }
}

// ---------- leitura e gravação ----------
function ler_() {
  const ss = planilha_();
  const tz = ss.getSpreadsheetTimeZone();
  const db = { versao: 1, abastecimentos: [], manutencoes: [], gastos: [] };
  Object.keys(ABAS).forEach(function (k) {
    const sh = ss.getSheetByName(ABAS[k].nome);
    if (!sh || sh.getLastRow() < 2) return;
    const valores = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const cab = valores.shift();
    db[k] = valores.filter(function (l) { return l[0] !== ''; }).map(function (l) {
      const o = {};
      cab.forEach(function (c, i) {
        let v = l[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, c === 'hora' ? 'HH:mm' : 'yyyy-MM-dd');
        o[c] = v;
      });
      if ('cheio' in o) o.cheio = o.cheio === true || String(o.cheio).toUpperCase() === 'TRUE';
      return o;
    });
  });
  const cfg = getConfig_('config');
  if (cfg) db.config = JSON.parse(cfg);
  return db;
}

function gravar_(db) {
  const ss = planilha_();
  Object.keys(ABAS).forEach(function (k) {
    const def = ABAS[k];
    const sh = ss.getSheetByName(def.nome) || ss.insertSheet(def.nome);
    sh.clearContents();
    const linhas = [def.campos].concat((db[k] || []).map(function (o) {
      return def.campos.map(function (c) { return o[c] === undefined || o[c] === null ? '' : o[c]; });
    }));
    // id, data, hora, foto e gps como texto puro, para o Sheets não converter.
    def.campos.forEach(function (c, i) {
      if (['id', 'data', 'hora', 'foto', 'gps'].indexOf(c) >= 0) sh.getRange(1, i + 1, Math.max(linhas.length, 1), 1).setNumberFormat('@');
    });
    sh.getRange(1, 1, linhas.length, def.campos.length).setValues(linhas);
    sh.setFrozenRows(1);
  });
  if (db.config) setConfig_('config', JSON.stringify(db.config));
}

// ---------- fotos de recibo (Google Drive) ----------
function salvarFoto_(b) {
  if (!b.dados || b.dados.length > 8000000) return { ok: false, erro: 'foto grande demais' };
  const it = DriveApp.getFoldersByName(PASTA_RECIBOS);
  const pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_RECIBOS);
  const nome = String(b.nome || 'recibo.jpg').replace(/[^\w.\-]/g, '_');
  const blob = Utilities.newBlob(Utilities.base64Decode(b.dados), b.mime || 'image/jpeg', nome);
  const arq = pasta.createFile(blob);
  return { ok: true, id: arq.getId(), url: arq.getUrl() };
}

// ---------- configuração (aba Motolog_Config: chave | valor) ----------
function abaConfig_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_CONFIG);
  if (!sh) { sh = ss.insertSheet(ABA_CONFIG); sh.appendRow(['chave', 'valor']); }
  return sh;
}
function getConfig_(chave) {
  const dados = abaConfig_().getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) if (dados[i][0] === chave) return dados[i][1];
  return null;
}
function setConfig_(chave, valor) {
  const sh = abaConfig_();
  const dados = sh.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === chave) { sh.getRange(i + 1, 2).setValue(valor); return; }
  }
  sh.appendRow([chave, valor]);
}
function rev_() { return Number(getConfig_('rev')) || 0; }

function tokenOk_(t) { return TOKEN !== 'TROQUE-ESTA-SENHA' && t === TOKEN; }
function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- resumo semanal por e-mail ----------
// Rode esta função uma vez no editor (Executar) para criar o gatilho de toda segunda às 8h.
function instalarResumoSemanal() {
  desinstalarResumoSemanal();
  ScriptApp.newTrigger('resumoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Logger.log('Gatilho criado. O e-mail vai para ' + Session.getEffectiveUser().getEmail() + ' toda segunda às 8h (se ligado nos Ajustes do app).');
}
function desinstalarResumoSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'resumoSemanal') ScriptApp.deleteTrigger(t); });
}
// Para testar na hora: rode esta função no editor (envia mesmo com o aviso desligado).
function testarResumoSemanal() { resumoSemanal(true); }

function resumoSemanal(forcar) {
  const db = ler_(), cfg = db.config || {};
  if (forcar !== true && !cfg.emailSemanal) return;
  const email = Session.getEffectiveUser().getEmail();
  if (!email) return;
  const tz = planilha_().getSpreadsheetTimeZone();
  const fmtD = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
  const hoje = new Date(), ini7 = fmtD(new Date(hoje.getTime() - 7 * 864e5)), mes = fmtD(hoje).slice(0, 7);
  const n = function (v) { return Number(v) || 0; };
  const brl = function (v) { return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); };
  const kmf = function (v) { return Math.round(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); };
  const todos = [].concat(db.abastecimentos, db.manutencoes, db.gastos || []);
  let semana = 0, mesTot = 0, combMes = 0;
  todos.forEach(function (x) { const d = String(x.data || ''); if (d > ini7) semana += n(x.valor); if (d.slice(0, 7) === mes) mesTot += n(x.valor); });
  db.abastecimentos.forEach(function (a) { if (String(a.data || '').slice(0, 7) === mes) combMes += n(a.valor); });

  // Consumo médio entre tanques cheios (igual ao app).
  const ab = db.abastecimentos.slice().sort(function (a, b) { return n(a.km) - n(b.km); });
  let kmTot = 0, lTot = 0, ult = null, acum = 0, ultimoKml = null;
  ab.forEach(function (a) {
    if (ult) { acum += n(a.litros); if (a.cheio) { const d = n(a.km) - n(ult.km); if (d > 0 && acum > 0) { kmTot += d; lTot += acum; ultimoKml = d / acum; } ult = a; acum = 0; } }
    else if (a.cheio) { ult = a; acum = 0; }
  });
  const media = lTot ? kmTot / lTot : null;
  const kmAtual = Math.max.apply(null, [n(cfg.km0)].concat(db.abastecimentos.map(function (a) { return n(a.km); }), db.manutencoes.map(function (m) { return n(m.km); })));

  // Próximas trocas (km ou meses, o que vencer primeiro).
  const intv = cfg.intervalos || {}, meses = cfg.meses || {}, trocas = [];
  Object.keys(Object.assign({}, intv, meses)).forEach(function (nome) {
    const ik = n(intv[nome]), im = n(meses[nome]);
    if ((!ik && !im) || nome === 'Outro') return;
    const feitas = db.manutencoes.filter(function (m) { return m.servico === nome; }).sort(function (a, b) { return n(b.km) - n(a.km); });
    if (!feitas.length) return;
    const u = feitas[0];
    const frKm = ik ? (kmAtual - n(u.km)) / ik : -1;
    let frDt = -1, dias = null;
    if (im && u.data) {
      const d0 = new Date(u.data + 'T12:00:00'), lim = new Date(d0); lim.setMonth(lim.getMonth() + im);
      frDt = (hoje - d0) / (lim - d0); dias = Math.round((lim - hoje) / 864e5);
    }
    const fr = Math.max(frKm, frDt), porKm = frKm >= frDt, falta = n(u.km) + ik - kmAtual;
    const txt = fr >= 1 ? (porKm ? 'atrasada em ' + kmf(-falta) + ' km' : 'venceu') : (porKm ? 'faltam ' + kmf(falta) + ' km' : 'faltam ' + dias + ' dias');
    trocas.push({ nome: nome, fr: fr, txt: txt });
  });
  trocas.sort(function (a, b) { return b.fr - a.fr; });
  const cor = function (fr) { return fr >= 1 ? '#dc2626' : fr >= 0.85 ? '#d97706' : '#16a34a'; };

  const linhas = [
    ['Gasto nos últimos 7 dias', brl(semana)],
    ['Gasto no mês', brl(mesTot) + ' (combustível ' + brl(combMes) + ')'],
    ['Consumo médio', media ? media.toFixed(1).replace('.', ',') + ' km/l' : '–'],
    ['Último tanque', ultimoKml ? ultimoKml.toFixed(1).replace('.', ',') + ' km/l' : '–'],
    ['Km atual', kmAtual ? kmf(kmAtual) + ' km' : '–'],
  ];
  const html = '<div style="font-family:Arial,sans-serif;max-width:520px">' +
    '<h2 style="margin:0 0 12px">🏍️ Motolog · resumo da semana</h2>' +
    '<table style="border-collapse:collapse;width:100%">' + linhas.map(function (l) {
      return '<tr><td style="padding:6px 0;color:#555">' + l[0] + '</td><td style="padding:6px 0;text-align:right"><b>' + l[1] + '</b></td></tr>';
    }).join('') + '</table>' +
    '<h3 style="margin:18px 0 6px">Próximas trocas</h3>' +
    (trocas.length ? '<ul style="padding-left:18px;margin:0">' + trocas.slice(0, 6).map(function (t) {
      return '<li style="margin:4px 0"><b>' + t.nome + '</b>: <span style="color:' + cor(t.fr) + '">' + t.txt + '</span></li>';
    }).join('') + '</ul>' : '<p style="color:#555">Nenhuma manutenção registrada ainda.</p>') +
    '<p style="color:#888;font-size:12px;margin-top:18px">Para parar de receber, desligue em Ajustes no app.</p></div>';
  const alerta = trocas.filter(function (t) { return t.fr >= 0.85; }).length;
  MailApp.sendEmail({ to: email, subject: 'Motolog: ' + brl(semana) + ' na semana' + (alerta ? ' · ' + alerta + ' troca(s) chegando' : ''), htmlBody: html });
}
