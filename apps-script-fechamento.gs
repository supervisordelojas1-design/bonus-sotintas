/**********************************************************************
 * SÓ TINTAS — Recepção do Simulador de Bônus + API do Dashboard
 * --------------------------------------------------------------------
 * Este script faz DUAS coisas na mesma URL /exec:
 *   1) Recebe os envios do simulador (botão "Enviar para o gerente")
 *      e grava uma linha na aba "Envios".
 *   2) Quando chamado com ?action=list&callback=fn, devolve (em JSONP)
 *      o ÚLTIMO envio de cada vendedor no mês — é o que alimenta o
 *      dashboard-fechamento.html.
 *
 * COMO USAR:
 *   1. Abra a planilha do Google onde quer guardar os dados (Drive).
 *   2. Menu Extensões > Apps Script.
 *   3. Apague o conteúdo e cole ESTE arquivo inteiro.
 *   4. Implantar > Nova implantação > Tipo: App da Web.
 *        - Executar como: Eu
 *        - Quem tem acesso: Qualquer pessoa
 *   5. Copie a URL /exec e use a MESMA nos dois lugares:
 *        - no simulador (constante ENVIO_URL)
 *        - no dashboard (constante API_URL)
 *   Obs.: ao alterar o script, use "Gerenciar implantações" e edite a
 *   implantação existente (mantém a mesma URL) OU crie nova (muda a URL).
 **********************************************************************/

var SHEET_NAME = 'Envios';

// Ordem das colunas gravadas. Se mudar, apague a 1ª linha da aba p/ recriar.
var HEADERS = [
  'timestamp','nome','login','filial','supervisor','mes',
  'nivelLoja','nivelIdx','alcLoja','alcLojaNum','alcVend','alcVendNum',
  'meta','faturado','projetado','projetadoNum',
  'ganhoHoje','potencial','faltaBuscar','atingidos','pendentes','telefone','emoji',
  'ident','itens','ticket','descCpf','fatCnpj','descCnpj','posit','novosCnpj','pctPosit'
];

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.action === 'list') {
    return listar_(p.callback);
  }
  if (p.action === 'stats') {
    return stats_(p.callback);
  }
  // Sem action = é um envio do simulador -> grava
  if (p.nome || p.login || p.filial) {
    gravar_(p);
    return resposta_({ ok: true, saved: true }, p.callback);
  }
  return resposta_({ ok: true, info: 'API do simulador Só Tintas ativa.' }, p.callback);
}

// Alguns navegadores mandam POST; tratamos igual.
function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  gravar_(p);
  return resposta_({ ok: true, saved: true }, p.callback);
}

// Colunas que devem ser SEMPRE texto (senão o Sheets converte em data/número
// e quebra o mês, a filial com zero à esquerda, login, etc.).
var COLS_TEXTO = ['login','filial','mes','alcLojaNum','alcVendNum','telefone','emoji','nivelIdx'];

function planilha_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  var novo = false;
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); novo = true; }
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); novo = true; }
  if (novo) { forcarColunasTexto_(sh); }
  return sh;
}

// Marca as colunas sensíveis como formato TEXTO ("@") para o Sheets não converter.
function forcarColunasTexto_(sh) {
  try {
    var max = sh.getMaxRows();
    COLS_TEXTO.forEach(function (h) {
      var c = HEADERS.indexOf(h);
      if (c >= 0) sh.getRange(1, c + 1, max, 1).setNumberFormat('@');
    });
  } catch (e) {}
}

// Reaplica o formato texto na aba atual (rode 1x pelo editor após colar, se a aba já existia).
function blindarPlanilha() {
  var sh = planilha_();
  forcarColunasTexto_(sh);
  return 'Colunas de texto blindadas: ' + COLS_TEXTO.join(', ');
}

function gravar_(p) {
  var sh = planilha_();
  var row = HEADERS.map(function (h) {
    if (h === 'timestamp') return new Date();
    return (p[h] !== undefined && p[h] !== null) ? p[h] : '';
  });
  sh.appendRow(row);
}

function idxHeaders_(){ var idx={}; HEADERS.forEach(function(h,i){ idx[h]=i; }); return idx; }
function normMes_(v){
  if(v===null||v===undefined||v==='') return '';
  if(Object.prototype.toString.call(v)==='[object Date]') return v.getFullYear()+'-'+('0'+(v.getMonth()+1)).slice(-2);
  var s=String(v); var m=s.match(/^(\d{4})-(\d{2})/); if(m) return m[1]+'-'+m[2];
  var d=new Date(s); if(!isNaN(d.getTime())) return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2);
  return s;
}
function normFilial_(v){ if(v===null||v===undefined) return ''; var n=Number(v); return (v!==''&&!isNaN(n))?String(n):String(v); }
// Dia no fuso do Brasil (America/Sao_Paulo): garante corte 00:00:00-23:59:59 local
function diaBR_(d){ try{ return Utilities.formatDate(d,'America/Sao_Paulo','yyyy-MM-dd'); }catch(e){ return d.toISOString().slice(0,10); } }

function listar_(callback) {
  var sh = planilha_();
  var values = sh.getDataRange().getValues();
  values.shift(); // pula a linha de cabecalho (mapeamento e pela ordem fixa de HEADERS)
  var idx = idxHeaders_();
  var map = {};
  values.forEach(function (r) {
    var login = (r[idx.login] || r[idx.nome] || '').toString().trim().toLowerCase();
    var mes = normMes_(r[idx.mes]);
    var key = login + '|' + mes;
    var rawTs = r[idx.timestamp];
    var dt = rawTs ? new Date(rawTs) : null;
    var valido = dt && !isNaN(dt.getTime());
    var ts = valido ? dt.getTime() : 0;
    if (!map[key] || ts >= map[key]._ts) {
      var o = {};
      HEADERS.forEach(function (h, i) { o[h] = r[i]; });
      o.mes = mes;
      o.filial = normFilial_(r[idx.filial]);
      o.emoji = (r[idx.emoji] || '');
      o.telefone = (r[idx.telefone] || '');
      o._ts = ts;
      o.timestamp = valido ? dt.toISOString() : (rawTs ? String(rawTs) : '');
      map[key] = o;
    }
  });
  var rows = Object.keys(map).map(function (k) { return map[k]; });
  return resposta_({ ok: true, rows: rows, geradoEm: new Date().toISOString() }, callback);
}

// Engajamento: agrega por login+mes (nº de envios, dias distintos, ultimo acesso).
function stats_(callback) {
  var sh = planilha_();
  var v = sh.getDataRange().getValues();
  v.shift();
  var idx = idxHeaders_();
  var map = {};
  v.forEach(function (r) {
    var login = (r[idx.login] || r[idx.nome] || '').toString().trim().toLowerCase();
    if (!login) return;
    var mes = normMes_(r[idx.mes]);
    var key = login + '|' + mes;
    var ts = r[idx.timestamp] ? new Date(r[idx.timestamp]) : null;
    var valido = ts && !isNaN(ts.getTime());
    if (!map[key]) map[key] = { login: login, nome: r[idx.nome] || login,
      filial: normFilial_(r[idx.filial]), supervisor: r[idx.supervisor] || '',
      emoji: (r[idx.emoji] || ''), mes: mes,
      envios: 0, diasSet: {}, ultimoTs: 0, ultimo: '', alcVendNum: '' };
    var o = map[key];
    o.envios++;
    if (valido) {
      o.diasSet[diaBR_(ts)] = 1;
      if (ts.getTime() >= o.ultimoTs) {
        o.ultimoTs = ts.getTime();
        o.ultimo = ts.toISOString();
        o.nome = r[idx.nome] || o.nome;
        o.emoji = (r[idx.emoji] || o.emoji);
        o.alcVendNum = r[idx.alcVendNum];
      }
    }
  });
  var rows = Object.keys(map).map(function (k) {
    var o = map[k]; o.dias = Object.keys(o.diasSet).length; delete o.diasSet; delete o.ultimoTs; return o;
  });
  return resposta_({ ok: true, rows: rows, geradoEm: new Date().toISOString() }, callback);
}

function resposta_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
