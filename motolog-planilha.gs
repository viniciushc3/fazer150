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
 */
const TOKEN = 'TROQUE-ESTA-SENHA';

// Deixe vazio para usar a planilha onde o script foi colado.
// (Só preencha com um ID se o script estiver num projeto separado.)
const PLANILHA_ID = '';
function planilha_() { return PLANILHA_ID ? SpreadsheetApp.openById(PLANILHA_ID) : SpreadsheetApp.getActive(); }

const ABAS = {
  abastecimentos: { nome: 'Motolog_Abastecimentos', campos: ['id', 'data', 'km', 'litros', 'valor', 'comb', 'posto', 'cheio', 'foto'] },
  manutencoes:    { nome: 'Motolog_Manutencoes',    campos: ['id', 'data', 'km', 'servico', 'outro', 'valor', 'oficina', 'obs', 'foto'] },
  gastos:         { nome: 'Motolog_Gastos',         campos: ['id', 'data', 'valor', 'cat', 'desc', 'foto'] },
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
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
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
    // id e data como texto puro, para o Sheets não converter.
    sh.getRange(1, 1, Math.max(linhas.length, 1), 2).setNumberFormat('@');
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
