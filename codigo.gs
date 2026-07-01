// ============================================================
// JIRA-MEDIAS-CONTAS — Apps Script (Coleta Jira → Firestore)
// Fork de: Portfólio Pequenas Contas | Versão: 2.0 | Data: 23/06/2026
// Backend: Firebase Firestore (REST API via UrlFetchApp)
// ============================================================
//
// SETUP (executar UMA vez):
//
// 1. Menu Extensões → Apps Script → cole este código e salve
// 2. Copie também o conteúdo de appsscript.json para o arquivo manifest
//    (ícone de engrenagem → Mostrar arquivo de manifesto "appsscript.json")
// 3. Configurações → Propriedades do script → adicione:
//    JIRA_BASE_URL   → https://atendimento.betha.com.br
//    JIRA_USERNAME  → usuário de integração Jira
//    JIRA_PASSWORD  → senha do usuário de integração
// 4. Execute testarConexaoJira() para validar o acesso ao Jira
// 5. Execute testarFirestore() para validar o acesso ao Firestore
// 6. Execute onTimeTrigger() manualmente para popular o Firestore
// 7. Execute setupTrigger() para ativar atualização automática
//    (a cada 30 min, entre 08:00 e 18:00)
//
// RESET (limpar todos os dados e recomeçar):
//    Execute resetFirestore() — apaga todas as coleções gerenciadas
// ============================================================

// ────────────────────────────────────────────────────────────
// FIRESTORE — configuração
// ────────────────────────────────────────────────────────────
const FIRESTORE_PROJECT_ID = 'jira-medias-contas';
const FIRESTORE_BASE = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID + '/databases/(default)/documents';

// Coleções gerenciadas por este script
const COL_CHAMADOS        = 'jira_chamados';
const COL_CHAMADOS_ISSUES = 'jira_chamados_issues';
const COL_META            = '_meta';

// ────────────────────────────────────────────────────────────
// JIRA — constantes
// ────────────────────────────────────────────────────────────
const FIELD_MUNICIPIO       = 'customfield_10331';
const FIELD_VERTICAL        = 'customfield_10300';
const FIELD_SLO_ATENDIMENTO = 'customfield_24813';

const _MUN_LIST = '"Água Doce", "Águas de Chapecó", Anitápolis, Armazém, Atalanta, "Balneário Camboriú", "Balneário Gaivota", "Balneário Rincão", "Bom Jardim da Serra", "Braço do Norte", "Braço do Trombudo", Caçador, "Campo Belo do Sul", Capinzal, "Celso Ramos", Chapecó, "Cocal do Sul", Descanso, "Dionísio Cerqueira", Forquilhinha, "Grão Pará", "Herval d\'Oeste", Ibiam, Imaruí, Itajaí, "Lauro Müller", "Major Gercino", "Major Vieira", Maracajá, Maravilha, Mondaí, "Monte Carlo", "Monte Castelo", "Nova Veneza", Orleans, Painel, Palmitos, "Passo de Torres", "Passos Maia", "Paulo Lopes", "Pedras Grandes", Petrolândia, Piratuba, "Ponte Alta", "Pouso Redondo", "Rancho Queimado", Sangão, "Santa Cecília", "São José do Cerrito", "São Lourenço do Oeste", "São Ludgero", "São Miguel do Oeste", Schroeder, Siderópolis, Sombrio, "Trombudo Central", Turvo, Urubici, Vargem, "Vargem Bonita", "Vidal Ramos"';

const JQL = 'category = "Projetos ativos de atendimento - Filial" AND resolution = Unresolved AND issuetype not in (Melhoria, "Melhoria (sub-tarefa)") AND Município in (' + _MUN_LIST + ')';

// ============================================================
// TRIGGER PRINCIPAL
// ============================================================
function onTimeTrigger() {
  const inicio = new Date();
  Logger.log('▶ Iniciando coleta: ' + inicio.toLocaleString('pt-BR'));
  try {
    const issues = fetchJiraIssues(JQL);
    Logger.log('  Issues coletadas: ' + issues.length);

    const rows   = aggregateByMunicipioVertical(issues);
    Logger.log('  Linhas agregadas: ' + rows.length + ' (municipio × vertical)');

    writeJiraChamadosToFirestore(rows, issues);

    Logger.log('✅ Concluído em ' + Math.round((new Date() - inicio) / 1000) + 's');
  } catch (e) {
    Logger.log('❌ ERRO: ' + e.message);
    throw e;
  }
}

// ============================================================
// COLETA JIRA
// ============================================================
function fetchJiraIssues(jql) {
  const props    = PropertiesService.getScriptProperties();
  const BASE_URL = props.getProperty('JIRA_BASE_URL') || '';
  const email    = props.getProperty('JIRA_USERNAME')    || '';
  const token    = props.getProperty('JIRA_PASSWORD') || '';
  if (!BASE_URL || !email || !token) throw new Error('Script Properties ausentes: configure JIRA_BASE_URL, JIRA_USERNAME e JIRA_PASSWORD.');

  const auth     = Utilities.base64Encode(email + ':' + token);
  const fields   = ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL, 'status', FIELD_SLO_ATENDIMENTO, 'assignee', 'issuetype'];
  const PAGE_SIZE = 100;
  let all = [], startAt = 0, total = Infinity;

  while (startAt < total) {
    const body = JSON.stringify({ jql, fields, maxResults: PAGE_SIZE, startAt });
    const resp = UrlFetchApp.fetch(BASE_URL + '/rest/api/2/search', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + auth },
      payload: body,
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error('Jira HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));

    const result = JSON.parse(resp.getContentText());
    total = result.total || 0;
    const issues = result.issues || [];
    if (issues.length === 0) break;
    all = all.concat(issues);
    startAt += issues.length;
    if (issues.length < PAGE_SIZE) break;
  }
  return all;
}

// ============================================================
// AGREGAÇÃO
// ============================================================
function aggregateByMunicipioVertical(issues) {
  const map = {};
  issues.forEach(function(issue) {
    const municipio   = (issue.fields[FIELD_MUNICIPIO] || 'Não informado').toString().trim();
    const vObj        = issue.fields[FIELD_VERTICAL];
    const vertical    = (vObj && vObj.value) ? vObj.value.trim() : 'Não informado';
    const sloBreached = _isSloBreached(issue.fields[FIELD_SLO_ATENDIMENTO]) ? 1 : 0;
    const key = municipio + '||' + vertical;
    if (!map[key]) map[key] = { municipio, vertical, total_chamados: 0, slo_estourado: 0 };
    map[key].total_chamados += 1;
    map[key].slo_estourado  += sloBreached;
  });
  const ts = new Date().toISOString();
  return Object.values(map).map(function(r) {
    r.atualizado_em = ts;
    return r;
  }).sort(function(a, b) {
    return a.municipio.localeCompare(b.municipio, 'pt-BR') || a.vertical.localeCompare(b.vertical, 'pt-BR');
  });
}

// ============================================================
// WRITE PARA FIRESTORE
// ============================================================
function writeJiraChamadosToFirestore(rows, issues) {
  const ts      = new Date().toISOString();
  const baseUrl = PropertiesService.getScriptProperties().getProperty('JIRA_BASE_URL') || '';

  // 1. Limpar coleções antes de reescrever
  Logger.log('  Limpando coleção "' + COL_CHAMADOS + '"...');
  clearFirestoreCollection(COL_CHAMADOS);

  Logger.log('  Limpando coleção "' + COL_CHAMADOS_ISSUES + '"...');
  clearFirestoreCollection(COL_CHAMADOS_ISSUES);

  // 2. Gravar agregados (municipio × vertical)
  if (rows.length > 0) {
    const chamadosDocs = {};
    rows.forEach(function(r) {
      const docId = _normDocId(r.municipio + '__' + r.vertical);
      chamadosDocs[docId] = r;
    });
    _writeCollectionDocs(COL_CHAMADOS, chamadosDocs);
    Logger.log('  "' + COL_CHAMADOS + '": ' + rows.length + ' docs gravados.');
  }

  // 3. Gravar issues individuais
  if (issues && issues.length > 0) {
    const issuesDocs = {};
    issues.forEach(function(issue) {
      const municipio  = (issue.fields[FIELD_MUNICIPIO] || 'Não informado').toString().trim();
      const vObj       = issue.fields[FIELD_VERTICAL];
      const vertical   = (vObj && vObj.value) ? vObj.value.trim() : 'Não informado';
      const status     = (issue.fields.status && issue.fields.status.name) ? issue.fields.status.name : '';
      const sloHoras   = _formatSloHoras(_getSloHoras(issue.fields[FIELD_SLO_ATENDIMENTO]));
      const responsavel= (issue.fields.assignee && issue.fields.assignee.displayName) ? issue.fields.assignee.displayName : '';
      const issuetype  = (issue.fields.issuetype && issue.fields.issuetype.name) ? issue.fields.issuetype.name : '';
      issuesDocs[issue.key] = {
        key:          issue.key,
        url:          baseUrl + '/browse/' + issue.key,
        summary:      issue.fields.summary || '',
        status:       status,
        municipio:    municipio,
        vertical:     vertical,
        slo_horas:    sloHoras,
        responsavel:  responsavel,
        issuetype:    issuetype,
        atualizado_em: ts
      };
    });
    _writeCollectionDocs(COL_CHAMADOS_ISSUES, issuesDocs);
    Logger.log('  "' + COL_CHAMADOS_ISSUES + '": ' + issues.length + ' docs gravados.');
  }

  // 4. Gravar metadado
  const metaDocs = {};
  metaDocs['last_update'] = {
    jira_chamados_at:     ts,
    jira_chamados_count:  rows.length,
    jira_issues_count:    issues ? issues.length : 0
  };
  _writeCollectionDocs(COL_META, metaDocs);
  Logger.log('  Metadado "_meta/last_update" atualizado.');
}

// ============================================================
// FIRESTORE — utilitários REST
// ============================================================

// Serializa qualquer valor JS para o formato de campo do Firestore REST API
function _toFirestoreValue(val) {
  if (val === null || val === undefined)   return { nullValue: null };
  if (typeof val === 'boolean')            return { booleanValue: val };
  if (typeof val === 'number') {
    return Number.isInteger(val) ? { integerValue: String(val) } : { doubleValue: val };
  }
  if (Array.isArray(val)) {
    return { arrayValue: { values: val.map(_toFirestoreValue) } };
  }
  if (typeof val === 'object') {
    var mf = {};
    Object.keys(val).forEach(function(k) { mf[k] = _toFirestoreValue(val[k]); });
    return { mapValue: { fields: mf } };
  }
  return { stringValue: String(val) };
}

// Converte um objeto JS plano em Fields do Firestore
function _toFirestoreFields(obj) {
  var fields = {};
  Object.keys(obj).forEach(function(key) { fields[key] = _toFirestoreValue(obj[key]); });
  return fields;
}

// Normaliza string para uso como doc ID (sem acentos, só alfanum e _)
function _normDocId(s) {
  return (s || '').toString().trim()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

// Executa um batch write (até 500 operações)
function _firestoreBatchWrite(writes) {
  var token   = ScriptApp.getOAuthToken();
  var url     = FIRESTORE_BASE + ':commit';
  var payload = JSON.stringify({ writes: writes });
  var resp    = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: payload,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Firestore batch write HTTP ' + code + ': ' + resp.getContentText().slice(0, 500));
  return JSON.parse(resp.getContentText());
}

// Grava ou atualiza documentos em uma coleção (upsert em lotes de 500)
function _writeCollectionDocs(collectionName, docsMap) {
  var writes = Object.keys(docsMap).map(function(docId) {
    return {
      update: {
        name:   'projects/' + FIRESTORE_PROJECT_ID + '/databases/(default)/documents/' + collectionName + '/' + docId,
        fields: _toFirestoreFields(docsMap[docId])
      }
    };
  });
  var BATCH = 500;
  for (var i = 0; i < writes.length; i += BATCH) {
    _firestoreBatchWrite(writes.slice(i, i + BATCH));
  }
}

// Lista todos os IDs de documentos de uma coleção (suporta paginação)
function _listCollectionDocIds(collectionName) {
  var token = ScriptApp.getOAuthToken();
  var ids   = [];
  var nextPageToken = null;
  do {
    var url  = FIRESTORE_BASE + '/' + collectionName + '?pageSize=300' + (nextPageToken ? '&pageToken=' + nextPageToken : '');
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) break;
    var result = JSON.parse(resp.getContentText());
    if (result.documents) {
      result.documents.forEach(function(doc) {
        ids.push(doc.name.split('/').pop());
      });
    }
    nextPageToken = result.nextPageToken || null;
  } while (nextPageToken);
  return ids;
}

// Apaga todos os documentos de uma coleção
function clearFirestoreCollection(collectionName) {
  var ids = _listCollectionDocIds(collectionName);
  if (ids.length === 0) { Logger.log('    "' + collectionName + '" já está vazia.'); return 0; }
  var deletes = ids.map(function(id) {
    return { delete: 'projects/' + FIRESTORE_PROJECT_ID + '/databases/(default)/documents/' + collectionName + '/' + id };
  });
  var BATCH = 500;
  for (var i = 0; i < deletes.length; i += BATCH) {
    _firestoreBatchWrite(deletes.slice(i, i + BATCH));
  }
  Logger.log('    "' + collectionName + '": ' + ids.length + ' docs deletados.');
  return ids.length;
}

// ============================================================
// RESET — apaga todas as coleções gerenciadas
// ============================================================
function resetFirestore() {
  Logger.log('🗑️ Iniciando reset do Firestore...');
  clearFirestoreCollection(COL_CHAMADOS);
  clearFirestoreCollection(COL_CHAMADOS_ISSUES);
  clearFirestoreCollection(COL_META);
  Logger.log('✅ Reset concluído. Execute onTimeTrigger() para repopular.');
}

// ============================================================
// SLO — helpers
// ============================================================
function _isSloBreached(sloField) {
  if (!sloField) return false;
  if (sloField.ongoingCycle) {
    var oc = sloField.ongoingCycle;
    if (oc.breached === true) return true;
    if (oc.remainingTime && typeof oc.remainingTime.millis === 'number' && oc.remainingTime.millis < 0) return true;
  }
  if (Array.isArray(sloField.completedCycles) && sloField.completedCycles.length > 0) {
    var last = sloField.completedCycles[sloField.completedCycles.length - 1];
    if (last && last.breached === true) return true;
  }
  return false;
}

function _getSloHoras(sloField) {
  if (!sloField) return null;
  if (sloField.ongoingCycle && sloField.ongoingCycle.remainingTime) {
    return sloField.ongoingCycle.remainingTime.millis / 3600000;
  }
  if (Array.isArray(sloField.completedCycles) && sloField.completedCycles.length > 0) {
    var last = sloField.completedCycles[sloField.completedCycles.length - 1];
    if (last && last.remainingTime) return last.remainingTime.millis / 3600000;
  }
  return null;
}

function _formatSloHoras(horas) {
  if (horas === null || horas === undefined) return '';
  var abs = Math.abs(horas);
  var d   = Math.floor(abs / 24);
  var h   = Math.floor(abs % 24);
  var prefix = horas < 0 ? '-' : '+';
  return prefix + (d > 0 ? d + 'd' : '') + h + 'h';
}

// ============================================================
// TRIGGER E DIAGNÓSTICO
// ============================================================

// Instala 5 triggers diários fixos: 08h, 10h, 13h, 16h, 23h (executar UMA vez)
function setupTrigger() {
  // Remove todos os triggers existentes de onTimeTrigger
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'onTimeTrigger'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // Cria 5 triggers nos horários desejados
  [8, 10, 13, 16, 23].forEach(function(hora) {
    ScriptApp.newTrigger('onTimeTrigger')
      .timeBased()
      .everyDays(1)
      .atHour(hora)
      .create();
  });

  Logger.log('✅ 5 triggers configurados: 08h, 10h, 13h, 16h, 23h.');
}

// Testa a conexão com o Jira
function testarConexaoJira() {
  const props    = PropertiesService.getScriptProperties();
  const BASE_URL = props.getProperty('JIRA_BASE_URL') || '';
  const email    = props.getProperty('JIRA_USERNAME')    || '';
  const token    = props.getProperty('JIRA_PASSWORD') || '';
  if (!BASE_URL || !email || !token) { Logger.log('❌ Script Properties ausentes.'); return; }

  const auth = Utilities.base64Encode(email + ':' + token);
  const body = JSON.stringify({ jql: JQL, maxResults: 1, fields: ['summary', FIELD_MUNICIPIO, FIELD_VERTICAL] });
  const resp = UrlFetchApp.fetch(BASE_URL + '/rest/api/2/search', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Basic ' + auth },
    payload: body, muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code !== 200) { Logger.log('❌ Jira HTTP ' + code + ': ' + resp.getContentText().slice(0, 300)); return; }
  const result = JSON.parse(resp.getContentText());
  Logger.log('✅ Conexão Jira OK! Total de issues no filtro: ' + result.total);
  if (result.issues && result.issues[0]) {
    const i = result.issues[0];
    Logger.log('   Exemplo → ' + i.key + ' | Município: ' + (i.fields[FIELD_MUNICIPIO] || '(vazio)') + ' | Vertical: ' + ((i.fields[FIELD_VERTICAL] || {}).value || '(vazio)'));
  }
}

// Testa a conexão com o Firestore
function testarFirestore() {
  try {
    // Tenta ler a coleção _meta (pode estar vazia)
    const ids = _listCollectionDocIds(COL_MET
// ────────────────────────────────────────────────────────────
// WEB APP — doPost: recebe dados da extensão eSfinge Updater
// POST JSON: { action: 'writeStatusEnvio'|'writeRatificacoes', docs: [...] }
// Cada doc deve ter um campo 'id' que vira o nome do documento Firestore.
// Todos os outros campos são serializados automaticamente (string, int,
// boolean, null, array) via _toFirestoreValue().
// ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!Array.isArray(body.docs)) {
      return _jsonResp({ ok: false, error: 'docs deve ser array' });
    }
    var collection;
    if      (body.action === 'writeStatusEnvio')   collection = 'status_envio';
    else if (body.action === 'writeRatificacoes')  collection = 'ratificacoes';
    else return _jsonResp({ ok: false, error: 'Ação inválida: ' + body.action });

    var docsMap = {};
    body.docs.forEach(function(d) {
      var doc = {};
      Object.keys(d).forEach(function(k) { if (k !== 'id') doc[k] = d[k]; });
      docsMap[d.id] = doc;
    });
    _writeCollectionDocs(collection, docsMap);
    return _jsonResp({ ok: true, count: body.docs.length });
  } catch (err) {
    return _jsonResp({ ok: false, error: err.message });
  }
}

function _jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
