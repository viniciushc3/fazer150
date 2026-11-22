/**
 * Motolog: API da planilha Google (Google Apps Script), com login por pessoa.
 *
 * Cole este código no Apps Script da planilha do Motolog (Extensões > Apps Script),
 * salve e reimplante como NOVA VERSÃO (Implantar > Gerenciar implantações > lápis >
 * Versão: Nova versão > Implantar). A URL /exec continua a mesma.
 *
 * Como funciona:
 * - Cada pessoa cria a conta no app (nome, e-mail e senha). A senha é guardada
 *   embaralhada (hash com sal), nunca em texto.
 * - Todo cadastro novo fica "pendente". O dono da planilha recebe um e-mail e aprova
 *   trocando a coluna "status" para "ativo" na aba Motolog_Usuarios.
 *   (Para bloquear alguém, troque para "bloqueado".)
 * - Os registros ficam nas mesmas abas, com a coluna "usuario" (o e-mail da pessoa).
 *   Cada um só vê e altera os próprios dados.
 * - Os registros antigos, sem usuário, pertencem ao dono da planilha.
 * - O script só lê e grava as abas que começam com "Motolog_".
 * - As fotos de recibo vão para a pasta "Motolog Recibos" no Drive do dono.
 *
 * Resumo semanal por e-mail: rode "instalarResumoSemanal" uma vez no editor
 * (escolha a função no topo e clique em Executar). Cada pessoa liga ou desliga
 * o próprio resumo nos Ajustes do app.
 */

// Deixe vazio para usar a planilha onde o script foi colado.
const PLANILHA_ID = '';
function planilha_() { return PLANILHA_ID ? SpreadsheetApp.openById(PLANILHA_ID) : SpreadsheetApp.getActive(); }

const ABAS = {
  abastecimentos: { nome: 'Motolog_Abastecimentos', campos: ['usuario', 'id', 'data', 'hora', 'km', 'litros', 'valor', 'comb', 'posto', 'cheio', 'foto', 'gps'] },
  manutencoes:    { nome: 'Motolog_Manutencoes',    campos: ['usuario', 'id', 'data', 'hora', 'km', 'servico', 'outro', 'valor', 'oficina', 'obs', 'foto'] },
  gastos:         { nome: 'Motolog_Gastos',         campos: ['usuario', 'id', 'data', 'hora', 'valor', 'cat', 'desc', 'foto'] },
};
const ABA_CONFIG = 'Motolog_Config';
const ABA_USUARIOS = 'Motolog_Usuarios';
const CAMPOS_USUARIOS = ['email', 'nome', 'status', 'criado', 'ultimo_acesso', 'sal', 'hash'];
const PASTA_RECIBOS = 'Motolog Recibos';
const SESSAO_DIAS = 365;

function dono_() { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); }
function normEmail_(e) { return String(e || '').trim().toLowerCase(); }

// ---------- entrada ----------
function doGet(e) {
  const u = sessao_(e.parameter.sessao);
  if (!u) return saida_({ ok: false, erro: 'sessao' });
  return saida_({ ok: true, rev: rev_(u.email), db: ler_(u.email), usuario: { email: u.email, nome: u.nome } });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return saida_({ ok: false, erro: 'json' }); }
  if (body.acao === 'cadastro') return saida_(cadastro_(body));
  if (body.acao === 'login') return saida_(login_(body));

  const u = sessao_(body.sessao);
  if (!u) return saida_({ ok: false, erro: 'sessao' });
  if (body.acao === 'sair') { PropertiesService.getScriptProperties().deleteProperty('s:' + body.sessao); return saida_({ ok: true }); }
  if (body.acao === 'foto') return saida_(salvarFoto_(body, u.email));
  if (!body.db || !Array.isArray(body.db.abastecimentos) || !Array.isArray(body.db.manutencoes)) {
    return saida_({ ok: false, erro: 'formato' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const atual = rev_(u.email);
    // Evita que um aparelho desatualizado apague dados gravados por outro.
    if (body.baseRev !== 'force' && Number(body.baseRev) !== atual) {
      return saida_({ ok: false, erro: 'conflito', rev: atual, db: ler_(u.email) });
    }
    gravar_(u.email, body.db);
    const nova = atual + 1;
    setConfig_(chaveRev_(u.email), nova);
    return saida_({ ok: true, rev: nova });
  } finally {
    lock.releaseLock();
  }
}

// ---------- contas ----------
function abaUsuarios_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_USUARIOS);
  if (!sh) {
    sh = ss.insertSheet(ABA_USUARIOS);
    sh.getRange('A:G').setNumberFormat('@');
    sh.getRange(1, 1, 1, CAMPOS_USUARIOS.length).setValues([CAMPOS_USUARIOS]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function usuarios_() {
  const sh = abaUsuarios_(), v = sh.getDataRange().getValues(), cab = v.shift();
  return v.map(function (l, i) {
    const o = { linha: i + 2 };
    cab.forEach(function (c, j) { o[c] = l[j]; });
    o.email = normEmail_(o.email); o.status = String(o.status || '').trim().toLowerCase();
    return o;
  }).filter(function (o) { return o.email; });
}
function hash_(sal, senha) {
  let h = sal + '|' + senha;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + sal, Utilities.Charset.UTF_8));
  }
  return h;
}
function cadastro_(b) {
  const email = normEmail_(b.email), nome = String(b.nome || '').trim().slice(0, 60), senha = String(b.senha || '');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, erro: 'email' };
  if (senha.length < 6) return { ok: false, erro: 'senha_curta' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (usuarios_().some(function (u) { return u.email === email; })) return { ok: false, erro: 'existe' };
    const sal = Utilities.getUuid();
    abaUsuarios_().appendRow([email, nome, 'pendente', new Date().toISOString(), '', sal, hash_(sal, senha)]);
  } finally { lock.releaseLock(); }
  try {
    MailApp.sendEmail({
      to: dono_(),
      subject: 'Motolog: novo cadastro de ' + (nome || email),
      htmlBody: '<p><b>' + (nome || '(sem nome)') + '</b> (' + email + ') pediu acesso ao Motolog.</p>' +
        '<p>Para aprovar, abra a planilha, aba <b>' + ABA_USUARIOS + '</b>, e troque o status de "pendente" para <b>ativo</b>.</p>' +
        '<p><a href="' + planilha_().getUrl() + '">Abrir a planilha</a></p>'
    });
  } catch (err) { /* sem permissão de e-mail ainda: o cadastro fica pendente mesmo assim */ }
  return { ok: true, status: 'pendente' };
}
function login_(b) {
  const email = normEmail_(b.email), senha = String(b.senha || '');
  const cache = CacheService.getScriptCache(), kf = 'falhas:' + email;
  const falhas = Number(cache.get(kf)) || 0;
  if (falhas >= 6) return { ok: false, erro: 'bloqueio' };
  const u = usuarios_().filter(function (x) { return x.email === email; })[0];
  if (!u || hash_(String(u.sal), senha) !== String(u.hash)) {
    cache.put(kf, String(falhas + 1), 600);
    return { ok: false, erro: 'credenciais' };
  }
  cache.remove(kf);
  if (u.status === 'pendente') return { ok: false, erro: 'pendente' };
  if (u.status !== 'ativo') return { ok: false, erro: 'bloqueado' };
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('s:' + token, JSON.stringify({ email: email, criado: Date.now() }));
  abaUsuarios_().getRange(u.linha, CAMPOS_USUARIOS.indexOf('ultimo_acesso') + 1).setValue(new Date().toISOString());
  return { ok: true, sessao: token, usuario: { email: email, nome: u.nome || '' } };
}
// Confere a sessão e se a conta continua ativa (bloquear na planilha derruba o acesso).
function sessao_(token) {
  if (!token || String(token).length < 40) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s:' + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (Date.now() - s.criado > SESSAO_DIAS * 864e5) { props.deleteProperty('s:' + token); return null; }
  const u = usuarios_().filter(function (x) { return x.email === s.email; })[0];
  if (!u || u.status !== 'ativo') return null;
  return { email: u.email, nome: u.nome || '' };
}

// ---------- leitura e gravação (só as linhas da pessoa) ----------
// Linhas sem usuário (de antes do login) pertencem ao dono da planilha.
function dele_(linhaUsuario, email) {
  const u = normEmail_(linhaUsuario);
  return u ? u === email : email === dono_();
}
function ler_(email) {
  const ss = planilha_();
  const tz = ss.getSpreadsheetTimeZone();
  const db = { versao: 1, abastecimentos: [], manutencoes: [], gastos: [] };
  Object.keys(ABAS).forEach(function (k) {
    const sh = ss.getSheetByName(ABAS[k].nome);
    if (!sh || sh.getLastRow() < 2) return;
    const valores = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const cab = valores.shift();
    const iu = cab.indexOf('usuario'), iid = cab.indexOf('id');
    db[k] = valores.filter(function (l) {
      return l[iid] !== '' && dele_(iu >= 0 ? l[iu] : '', email);
    }).map(function (l) {
      const o = {};
      cab.forEach(function (c, i) {
        if (c === 'usuario' || c === '') return;
        let v = l[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, c === 'hora' ? 'HH:mm' : 'yyyy-MM-dd');
        o[c] = v;
      });
      if ('cheio' in o) o.cheio = o.cheio === true || String(o.cheio).toUpperCase() === 'TRUE';
      return o;
    });
  });
  const cfg = getConfig_(chaveConfig_(email));
  if (cfg) db.config = JSON.parse(cfg);
  return db;
}
function gravar_(email, db) {
  const ss = planilha_();
  Object.keys(ABAS).forEach(function (k) {
    const def = ABAS[k];
    const sh = ss.getSheetByName(def.nome) || ss.insertSheet(def.nome);
    // Guarda as linhas das outras pessoas como estão.
    let outras = [];
    if (sh.getLastRow() >= 2) {
      const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), cab = v.shift();
      const iu = cab.indexOf('usuario'), iid = cab.indexOf('id');
      outras = v.filter(function (l) { return l[iid] !== '' && !dele_(iu >= 0 ? l[iu] : '', email); })
        .map(function (l) { return def.campos.map(function (c) { const i = cab.indexOf(c); return i >= 0 ? l[i] : ''; }); });
    }
    const minhas = (db[k] || []).map(function (o) {
      return def.campos.map(function (c) { if (c === 'usuario') return email; return o[c] === undefined || o[c] === null ? '' : o[c]; });
    });
    const linhas = [def.campos].concat(outras, minhas);
    sh.clearContents();
    // usuario, id, data, hora, foto e gps como texto puro, para o Sheets não converter.
    def.campos.forEach(function (c, i) {
      if (['usuario', 'id', 'data', 'hora', 'foto', 'gps'].indexOf(c) >= 0) sh.getRange(1, i + 1, Math.max(linhas.length, 1), 1).setNumberFormat('@');
    });
    sh.getRange(1, 1, linhas.length, def.campos.length).setValues(linhas);
    sh.setFrozenRows(1);
  });
  if (db.config) setConfig_(chaveConfig_(email), JSON.stringify(db.config));
}

// ---------- fotos de recibo (Google Drive do dono) ----------
function salvarFoto_(b, email) {
  if (!b.dados || b.dados.length > 8000000) return { ok: false, erro: 'foto grande demais' };
  const it = DriveApp.getFoldersByName(PASTA_RECIBOS);
  const pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_RECIBOS);
  const nome = (email.split('@')[0] + '_' + String(b.nome || 'recibo.jpg')).replace(/[^\w.\-]/g, '_');
  const blob = Utilities.newBlob(Utilities.base64Decode(b.dados), b.mime || 'image/jpeg', nome);
  const arq = pasta.createFile(blob);
  // Quem não é o dono só consegue abrir a foto pelo link.
  if (email !== dono_()) { try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {} }
  return { ok: true, id: arq.getId(), url: arq.getUrl() };
}

// ---------- configuração (aba Motolog_Config: chave | valor) ----------
// O dono continua usando as chaves antigas "config" e "rev".
function chaveConfig_(email) { return email === dono_() ? 'config' : 'config:' + email; }
function chaveRev_(email) { return email === dono_() ? 'rev' : 'rev:' + email; }
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
function rev_(email) { return Number(getConfig_(chaveRev_(email))) || 0; }

function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- resumo semanal por e-mail ----------
// Rode esta função uma vez no editor (Executar) para criar o gatilho de toda segunda às 8h.
function instalarResumoSemanal() {
  desinstalarResumoSemanal();
  ScriptApp.newTrigger('resumoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Logger.log('Gatilho criado: toda segunda às 8h, para cada pessoa que ligou o resumo nos Ajustes do app.');
}
function desinstalarResumoSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'resumoSemanal') ScriptApp.deleteTrigger(t); });
}
// Para testar na hora: envia o seu resumo agora (mesmo com o aviso desligado).
function testarResumoSemanal() { enviarResumo_(dono_(), true); }

function resumoSemanal() {
  const emails = usuarios_().filter(function (u) { return u.status === 'ativo'; }).map(function (u) { return u.email; });
  if (emails.indexOf(dono_()) < 0) emails.push(dono_());
  emails.forEach(function (email) { try { enviarResumo_(email, false); } catch (err) { Logger.log(email + ': ' + err); } });
}

function enviarResumo_(email, forcar) {
  const db = ler_(email), cfg = db.config || {};
  if (forcar !== true && !cfg.emailSemanal) return;
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
