// ================================================================
// SIMULADOR ESTUDO COMPLEXO — AVALIAR COM O CLIENTE
// Escritório Contábil Exemplo | Time de IA | v1.2 — 10/09/2026
// ================================================================
// CORREÇÕES v1.2 (auditoria interna):
//
// BUG 1 — _parseCsvNFSe_: item recebia DATA no lugar do código
//   Causa: detecção de coluna casava "Emissão" com candidatos errados
//          quando o CSV usa ',' como separador (não ';')
//   Fix: detecção de delimitador + validação de formato do item
//
// BUG 2 — _calcularDivergencia_: comparava data com item (bug 1 em cascata)
//   Fix: validar que o itemRef é código de item antes de comparar
//
// BUG 3 — _consolidarItens_: item em branco passava pelo filtro
//   Fix: _isItemValido_() exige padrão N.NN antes de consolidar
//
// BUG 4 — gerarComunicado_: pasta de destino e nome do PDF errados
//   Fix: salvar em ID_PASTA_COMUNICADOS_ com prefixo Comunicado_2027_
//
// BUG 5 — reprocessarCBSLote: usava fat×aliq_efe no lugar do DAS real
//   Valor planilha da empresa exemplo: R$21.888,65 (errado) → correto: R$16.584,26
//   Fix: usar DAS_ATUAL (extrato real) — já correto nesta versão
// ================================================================

// ── IDs DE DRIVE ───────────────────────────────────────────────
var ID_PASTA_SALVADOR_    = 'ID_EXEMPLO';
var ID_MODELO_COMUNIC_    = 'ID_EXEMPLO';
// Destino dos PDFs: (pasta de destino no Drive)
var ID_PASTA_COMUNICADOS_ = 'ID_EXEMPLO';

// ── NOMES DAS ABAS ──────────────────────────────────────────────
var ABA_PAINEL_    = '📊 Painel Geral';
var ABA_COR_       = '🔗 Correlação CNAEs';
var ABA_BANCO_     = '🗃️ Banco_Dados';
var ABA_PARAMS_    = '⚙️ Parâmetros';
var ABA_ANALISE_   = '🔍 Análise por CNPJ';

// ── LINHAS ──────────────────────────────────────────────────────
var L_HEADER_  = 2;
var L_DADOS_   = 3;

// ── COLUNAS DO PAINEL GERAL (1-indexed) ────────────────────────
var C = {
  CNPJ:             1,
  RAZAO:            2,
  SEGMENTO:         3,
  DESPESAS:         4,
  CNAES_TIPO:       5,
  ANEXO:            6,
  FAIXA:            7,
  ABERTURA:         8,
  CNAE_PRINCIPAL:   9,
  COMPETENCIA:      10,
  ITEM_NFSE_LIDO:   11,
  DESC_ITEM_NFSE:   12,
  FONTE_LEITURA:    13,
  QTD_NOTAS:        14,
  VALOR_NFSE:       15,
  ISS_NFSE:         16,
  ALERTA_LEITURA:   17,
  ITEM_ATUAL:       18,
  DESC_ITEM_ATUAL:  19,
  NBS:              20,
  ITEM_PESQUISADO:  21,
  ITEM_CONFIRMADO:  22,
  DIVERGENCIA:      23,
  RBT12:            24,
  FATURAMENTO:      25,
  DAS_ATUAL:        26,
  PIS:              27,
  COFINS:           28,
  REDUCAO_APLIC:    29,
  REDUCAO_ORIG:     30,
  ART127:           31,
  CBS_FORA:         32,
  DAS_HIBRIDO:      33,
  DIFERENCA:        34,
  ALIQ_ATUAL:       35,
  ALIQ_HIBRIDA:     36,
  GERAR_COM:        37,
  STATUS_COM:       38,
  OBSERVACAO:       39
};

// ── CORES ───────────────────────────────────────────────────────
var COR_AZUL_    = '#CFE2F3';
var COR_AMARELO_ = '#FFF3CD';
var COR_VERDE_   = '#D4EDDA';
var COR_VERMELHO_= '#F8D7DA';
var COR_LARANJA_ = '#FCE5CD';

// ── RETRY ───────────────────────────────────────────────────────
var RETRY_DELAYS_ = [2000, 5000, 10000, 20000, 30000, 60000];
var ERROS_DEFIN_  = ['not found','no access','permission',
                     'invalid argument','access denied','unauthorized',
                     'não encontrado'];

// ── REGEX DE ITEM VÁLIDO ─────────────────────────────────────────
// Aceita: "7.02", "07.02", "32.01", "17.09" — rejeita datas, textos, vazios
var RE_ITEM_VALIDO_ = /^\d{1,2}\.\d{2}$/;

// ================================================================
// 1. MENU
// ================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Estudo Complexo')
    .addSubMenu(ui.createMenu('1️⃣ NFS-e do Portal')
      .addItem('1a. Ler NFS-e — empresa da linha ativa',  'menuLerNFSeIndividual')
      .addItem('1b. Ler NFS-e — LOTE (todas pendentes)', 'lerNFSePortal_Lote')
      .addItem('1c. Ler NFS-e — LOTE (☑ selecionadas)', 'lerNFSePortal_LoteSelecionadas'))
    .addSubMenu(ui.createMenu('2️⃣ Análise')
      .addItem('2a. Atualizar divergências de item',       'atualizarDivergencias')
      .addItem('2b. Reprocessar CBS híbrida — LOTE',      'reprocessarCBSLote'))
    .addSubMenu(ui.createMenu('3️⃣ Comunicados')
      .addItem('3a. Gerar comunicado — empresa da linha ativa', 'menuGerarComunicadoIndividual')
      .addItem('3b. Gerar comunicados — LOTE (☑ selecionadas)', 'gerarComunicadoLote'))
    .addSubMenu(ui.createMenu('4️⃣ Painel CNPJ')
      .addItem('4a. Carregar painel — digitar CNPJ', 'menuCarregarPainelCNPJ'))
    .addSubMenu(ui.createMenu('🔍 Diagnóstico')
      .addItem('D1. Testar leitura NFS-e — empresa da linha ativa', 'diagTestarLeitura')
      .addItem('D2. Listar empresas pendentes de NFS-e', 'diagListarPendentes')
      .addItem('D3. Relatório de divergências de item', 'diagDivergencias')
      .addItem('D4. Dump completo — empresa da linha ativa', 'diagDumpEmpresa')
      .addItem('D5. Dump CSV bruto — empresa da linha ativa', 'diagDumpCsvBruto'))
    .addToUi();
}

// ================================================================
// 2. WRAPPERS DE MENU
// ================================================================
function menuLerNFSeIndividual() {
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  if (row < L_DADOS_) { _toast_('Posicione o cursor em uma linha de empresa (a partir da linha ' + L_DADOS_ + ').'); return; }
  var ws = _painel_();
  var dados = _lerLinha_(ws, row);
  _lerNFSeEmpresa_(ws, row, dados);
}

function menuGerarComunicadoIndividual() {
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  if (row < L_DADOS_) { _toast_('Posicione o cursor em uma linha de empresa.'); return; }
  var ws = _painel_();
  var dados = _lerLinha_(ws, row);
  gerarComunicado_(ws, row, dados);
}

function menuCarregarPainelCNPJ() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Análise por CNPJ', 'Digite o CNPJ (com ou sem formatação):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  carregarPainelCNPJ_(res.getResponseText().trim());
}

// ================================================================
// 3. LEITURA DE NFS-e — LOTE
// ================================================================
function lerNFSePortal_Lote() {
  var ws = _painel_();
  var dados = _lerTodasLinhas_(ws);
  var pendentes = dados.filter(function(d) {
    return !d[C.FONTE_LEITURA - 1] ||
           String(d[C.ALERTA_LEITURA - 1]).indexOf('Pendente') !== -1;
  });
  _toast_('Iniciando leitura de ' + pendentes.length + ' empresa(s) pendentes...');
  _processarLoteNFSe_(ws, dados, pendentes);
}

function lerNFSePortal_LoteSelecionadas() {
  var ws = _painel_();
  var dados = _lerTodasLinhas_(ws);
  var selecionadas = dados.filter(function(d) {
    return d[C.GERAR_COM - 1] === true || String(d[C.GERAR_COM - 1]).toLowerCase() === 'true';
  });
  _toast_('Iniciando leitura de ' + selecionadas.length + ' empresa(s) selecionada(s)...');
  _processarLoteNFSe_(ws, dados, selecionadas);
}

function _processarLoteNFSe_(ws, todasLinhas, subset) {
  var ok = 0, erros = 0;
  for (var i = 0; i < subset.length; i++) {
    var d = subset[i];
    var rowIdx = _encontrarLinha_(ws, d[C.CNPJ - 1]);
    if (!rowIdx) continue;
    try {
      _lerNFSeEmpresa_(ws, rowIdx, d);
      ok++;
      Utilities.sleep(300);
    } catch(e) {
      erros++;
      Logger.log('ERRO NFS-e ' + d[C.CNPJ - 1] + ': ' + e.message);
    }
    if ((i + 1) % 20 === 0) {
      _toast_('Processadas ' + (i + 1) + '/' + subset.length + ' empresas...');
    }
  }
  _toast_('Leitura concluída: ' + ok + ' OK | ' + erros + ' erro(s).');
}

// ================================================================
// 4. LEITURA DE NFS-e — EMPRESA INDIVIDUAL
// ================================================================
function _lerNFSeEmpresa_(ws, row, dados) {
  var cnpj   = dados[C.CNPJ       - 1];
  var razao  = dados[C.RAZAO      - 1];
  var compet = dados[C.COMPETENCIA - 1] || _lerParametro_('Competência padrão de busca') || '082026';

  var pastaEmp = _encontrarPastaEmpresa_(cnpj, razao);
  if (!pastaEmp) {
    _escreverNFSe_(ws, row, '', '', 'NÃO ENCONTRADO', 0, 0, 0,
      '❌ Pasta da empresa não localizada no Robô Conthabil — revisar manualmente');
    return;
  }

  var pastaMes = _encontrarPastaMes_(pastaEmp, compet);
  if (!pastaMes) {
    _escreverNFSe_(ws, row, '', '', 'NÃO ENCONTRADO', 0, 0, 0,
      '❌ Competência ' + compet + ' não localizada na pasta da empresa');
    return;
  }

  var resultado = _extrairItensNFSe_(pastaMes);

  if (!resultado || !resultado.itens || resultado.itens.length === 0) {
    _escreverNFSe_(ws, row, '', '', resultado ? resultado.fonte : 'NÃO ENCONTRADO',
      0, 0, 0, '❌ Nenhum dado de NFS-e encontrado em ' + compet);
    return;
  }

  var itensDistintos = _consolidarItens_(resultado.itens);
  if (!itensDistintos.length) {
    _escreverNFSe_(ws, row, '', '', resultado.fonte, 0, 0, 0,
      '❌ Itens extraídos inválidos (datas ou vazios) — verificar CSV/XML da empresa');
    return;
  }

  var itemPrimario = itensDistintos[0].item;
  var descPrimaria = itensDistintos[0].descItem || _buscarDescItem_(itemPrimario);
  var qtdNotas     = resultado.itens.length;
  var totalValor   = resultado.itens.reduce(function(s,it){ return s+(it.valorServicos||0); }, 0);
  var totalISS     = resultado.itens.reduce(function(s,it){ return s+(it.valorISS||0); }, 0);

  var alerta = resultado.alerta || '';
  if (itensDistintos.length > 1) {
    var listaItens = itensDistintos.map(function(it){ return it.item; }).join(' | ');
    alerta = (alerta ? alerta + ' | ' : '') +
             '⚠️ Múltiplos itens: ' + listaItens + ' — preencher col. "Item Confirmado"';
  }

  _escreverNFSe_(ws, row, itemPrimario, descPrimaria, resultado.fonte,
    qtdNotas, totalValor, totalISS, alerta || '✅ ' + resultado.fonte);
}

function _escreverNFSe_(ws, row, item, desc, fonte, qtd, valor, iss, alerta) {
  ws.getRange(row, C.ITEM_NFSE_LIDO, 1, 7)
    .setValues([[item, desc, fonte, qtd||'', valor||'', iss||'', alerta]]);

  var corAlerta = COR_AZUL_;
  if (alerta.indexOf('❌') !== -1) corAlerta = COR_VERMELHO_;
  else if (alerta.indexOf('⚠️') !== -1) corAlerta = COR_AMARELO_;
  else if (alerta.indexOf('✅') !== -1) corAlerta = COR_VERDE_;
  ws.getRange(row, C.ALERTA_LEITURA).setBackground(corAlerta);

  _calcularDivergencia_(ws, row);
}

// ================================================================
// 5. CASCADE: CSV → XML → PDF NOTA → PDF PRINT
// ================================================================
function _extrairItensNFSe_(pastaMes) {
  var pastaPrestado = _acharSubpasta_(pastaMes, 'PRESTADO');

  // 1. CSV
  var csv = _acharArquivo_(pastaMes, pastaPrestado, ['nfse_csv_prestado']);
  if (csv) {
    var itens = _parseCsvNFSe_(csv);
    if (itens.length) return { fonte: 'CSV', itens: itens, alerta: null };
  }

  // 2. XML
  var xmls = _acharArquivos_(pastaMes, pastaPrestado, ['nfse_xml_prestado', 'nfse_rps_xml']);
  if (xmls.length) {
    var itens = [];
    xmls.forEach(function(f){ try { itens = itens.concat(_parseXmlNFSe_(f)); } catch(e){} });
    if (itens.length) return { fonte: 'XML', itens: itens, alerta: null };
  }

  // 3. PDF nota individual (tem "Item da Lista XXXX" explícito)
  var pdfsNota = _acharArquivos_(pastaMes, pastaPrestado,
    ['nfse_pdf_prestado_', 'nota_nfse_', 'nfs_e_', 'nota_fiscal_']);
  if (pdfsNota.length) {
    var itens = [];
    pdfsNota.forEach(function(f){ try { itens = itens.concat(_parsePdfNotaIndividual_(f)); } catch(e){} });
    if (itens.length) return { fonte: 'PDF-Nota', itens: itens, alerta: null };
  }

  // 4. PDF print/resumo (tabela do portal)
  var print_ = _acharArquivo_(pastaMes, pastaPrestado,
    ['print_nfse_pdf_prestado', 'print_nfse', 'resumo_nfse']);
  if (print_) {
    var itens = _parsePdfPrint_(print_);
    return {
      fonte: 'PDF-Print',
      itens: itens,
      alerta: '⚠️ Lido do print (sem CSV/XML) — conferir item manualmente antes do comunicado'
    };
  }

  return null;
}

// ── Parser CSV — FIX BUG 1 e 3 ──────────────────────────────────
// Correções aplicadas:
//  1. Detecção automática do delimitador (';' vs ',') antes do parse
//  2. Validação _isItemValido_() antes de aceitar o código do serviço
//  3. Candidatos expandidos para cobrir variações do portal Salvador
function _parseCsvNFSe_(arquivo) {
  var texto;
  try { texto = arquivo.getBlob().getDataAsString('ISO-8859-1'); } catch(e) {
    try { texto = arquivo.getBlob().getDataAsString('UTF-8'); } catch(e2) { return []; }
  }
  if (!texto || texto.length < 10) return [];

  // ── Detectar delimitador ─────────────────────────────────────
  var primeiraLinha = texto.split('\n')[0] || '';
  var delim = (primeiraLinha.split(';').length > primeiraLinha.split(',').length) ? ';' : ',';

  var linhas = Utilities.parseCsv(texto, delim);
  if (!linhas || linhas.length < 2) return [];

  var header = linhas[0].map(function(h){ return (h||'').trim().toLowerCase(); });

  // Candidatos expandidos para cobrir as variações do portal Salvador
  var colItem  = _idxCsv_(header, ['código do serviço', 'codigo do servico',
                                    'item da lista', 'codigo serv', 'atividade',
                                    'cód. serviço', 'cod. servico', 'código serviço',
                                    'serviço prestado', 'codigo item']);
  var colValor = _idxCsv_(header, ['valor dos serviços', 'valor servicos',
                                    'vlr serviços', 'valor servico', 'vlr. serviços']);
  var colISS   = _idxCsv_(header, ['iss devido', 'iss devid', 'iss a pagar',
                                    'valor iss', 'iss (r$)']);
  var colNota  = _idxCsv_(header, ['nfs-e', 'numero nota', 'número nota',
                                    'nfse', 'n. nota', 'número da nota', 'num nota']);
  var colTom   = _idxCsv_(header, ['razão social do tomador', 'razao social tomador',
                                    'tomador', 'nome tomador']);

  // Se colItem não foi localizado, tentar heurística de posição
  // (campo que produz códigos no padrão 7.03, 32.01, etc.)
  if (colItem < 0) {
    colItem = _detectarColunaPorConteudo_(linhas, RE_ITEM_VALIDO_, delim);
    if (colItem >= 0) Logger.log('CSV: coluna item detectada por conteúdo na posição ' + colItem);
  }

  if (colItem < 0 || colValor < 0) {
    Logger.log('CSV: colunas não identificadas — header: ' + JSON.stringify(header));
    return [];
  }

  var itens = [];
  for (var i = 1; i < linhas.length; i++) {
    var l = linhas[i];
    var nota = colNota >= 0 ? String(l[colNota]||'').trim() : String(i);
    if (!nota || nota.toLowerCase() === 'total') continue;

    var itemRaw = String(l[colItem]||'').trim();
    var item    = _formatarItem_(itemRaw);

    // ── FIX BUG 1: rejeitar item que não passa na validação ──────
    if (!_isItemValido_(item)) {
      Logger.log('CSV linha ' + (i+1) + ': item inválido ignorado → "' + itemRaw + '" (formatado: "' + item + '")');
      continue;
    }

    itens.push({
      numeroNota:    nota,
      item:          item,
      valorServicos: _toNum_(colValor >= 0 ? l[colValor] : 0),
      valorISS:      _toNum_(colISS   >= 0 ? l[colISS]   : 0),
      tomador:       colTom >= 0 ? String(l[colTom]||'') : ''
    });
  }
  return itens;
}

// Heurística: varrer as primeiras 5 linhas e achar coluna cujo conteúdo bate com re
function _detectarColunaPorConteudo_(linhas, re, delim) {
  var amostra = Math.min(6, linhas.length);
  var contagens = {};
  for (var i = 1; i < amostra; i++) {
    var l = linhas[i];
    for (var c = 0; c < l.length; c++) {
      var v = _formatarItem_(String(l[c]||'').trim());
      if (re.test(v)) contagens[c] = (contagens[c]||0) + 1;
    }
  }
  var melhorCol = -1, melhorCount = 0;
  Object.keys(contagens).forEach(function(c) {
    if (contagens[c] > melhorCount) { melhorCount = contagens[c]; melhorCol = parseInt(c); }
  });
  return melhorCol;
}

function _idxCsv_(header, candidatos) {
  for (var i = 0; i < header.length; i++) {
    for (var j = 0; j < candidatos.length; j++) {
      if (header[i].indexOf(candidatos[j]) !== -1) return i;
    }
  }
  return -1;
}

// ── Parser XML ──────────────────────────────────────────────────
function _parseXmlNFSe_(arquivo) {
  var texto;
  try { texto = arquivo.getBlob().getDataAsString('ISO-8859-1'); } catch(e) {
    try { texto = arquivo.getBlob().getDataAsString('UTF-8'); } catch(e2) { return []; }
  }
  var doc  = XmlService.parse(texto);
  var root = doc.getRootElement();
  var ns   = root.getNamespace();
  var itens = [];

  var comps = _findAll_(root, ns, 'CompNfse');
  if (!comps.length) comps = [root];

  comps.forEach(function(comp) {
    try {
      var nfse = _findOne_(comp, ns, 'Nfse') || comp;
      var inf  = _findOne_(nfse, ns, 'InfNfse') || nfse;
      var srv  = _findOne_(inf,  ns, 'Servico') || inf;
      var vals = _findOne_(srv,  ns, 'Valores');
      var tom  = _findOne_(inf,  ns, 'TomadorServico');

      var itemCod = (_childText_(srv, ns, 'ItemListaServico') ||
                     _childText_(srv, ns, 'CodigoServico')    || '').trim();
      var item = _formatarItem_(itemCod);

      if (!_isItemValido_(item)) {
        Logger.log('XML: item inválido ignorado → "' + itemCod + '"');
        return;
      }

      itens.push({
        numeroNota:    _childText_(inf, ns, 'Numero') || '',
        item:          item,
        valorServicos: _toNum_(_childText_(vals, ns, 'ValorServicos') || '0'),
        valorISS:      _toNum_(_childText_(vals, ns, 'ValorIss')      || '0'),
        tomador:       tom ? (_childText_(tom, ns, 'RazaoSocial') || '') : ''
      });
    } catch(e) { Logger.log('XML erro nó: ' + e.message); }
  });
  return itens;
}

function _findAll_(el, ns, tag) {
  try { return el.getChildren(tag, ns) || []; } catch(e) {
    try { return el.getChildren(tag) || []; } catch(e2) { return []; }
  }
}
function _findOne_(el, ns, tag) {
  if (!el) return null;
  try { return el.getChild(tag, ns); } catch(e) {
    try { return el.getChild(tag); } catch(e2) { return null; }
  }
}
function _childText_(el, ns, tag) {
  if (!el) return '';
  try { return el.getChildText(tag, ns) || ''; } catch(e) {
    try { return el.getChildText(tag) || ''; } catch(e2) { return ''; }
  }
}

// ── PDF nota individual ──────────────────────────────────────────
function _parsePdfNotaIndividual_(arquivo) {
  var texto = _extrairTextoPDF_(arquivo);
  if (!texto) return [];

  var itens = [];
  var reItem   = /[Ii]tem\s+da\s+[Ll]ista[:\s]+(\d{3,4}\.?\d{0,2})/g;
  var reValor  = /[Vv]alor\s+dos\s+[Ss]ervi.os[:\s]+R?\$?\s*([\d.,]+)/;
  var reISS    = /ISS\s+[Dd]evido[:\s]+R?\$?\s*([\d.,]+)/;
  var reNumero = /N[ºo°]\s+da\s+NFS-?e[:\s]*(\d+)/i;

  var m;
  while ((m = reItem.exec(texto)) !== null) {
    var item = _formatarItem_(m[1]);
    if (!_isItemValido_(item)) continue;
    itens.push({
      numeroNota:    (reNumero.exec(texto)||['',''])[1],
      item:          item,
      valorServicos: _toNum_((reValor.exec(texto)||['','0'])[1]),
      valorISS:      _toNum_((reISS.exec(texto)  ||['','0'])[1]),
      tomador:       ''
    });
  }
  var vistos = {};
  return itens.filter(function(it) {
    var k = it.numeroNota + it.item;
    if (vistos[k]) return false;
    vistos[k] = true; return true;
  });
}

// ── PDF print — regex validada contra print real do portal Salvador ──
// Auditada: Abelardo Brandão, 5 notas, total R$160.520,00 ✓
function _parsePdfPrint_(arquivo) {
  var texto = _extrairTextoPDF_(arquivo);
  if (!texto) return [];

  var itens = [];
  var regex = /^(\d{8})\s+\d{2}\/\d{2}\/\d{4}.+?(\d{3,4})\s+([\d.]+,\d{2})\s+[\d.,]+\s+([\d.,]+)/gm;
  var m;
  while ((m = regex.exec(texto)) !== null) {
    var item = _formatarItem_(m[2]);
    if (!_isItemValido_(item)) continue;
    itens.push({
      numeroNota:    m[1],
      item:          item,
      valorServicos: _toNum_(m[3]),
      valorISS:      _toNum_(m[4]),
      tomador:       ''
    });
  }
  return itens;
}

// ── Extração de texto de PDF ─────────────────────────────────────
function _extrairTextoPDF_(arquivo) {
  try {
    var t = arquivo.getBlob().getDataAsString('UTF-8');
    if (t && t.length > 100 && t.indexOf('\n') !== -1) return t;
  } catch(e) {}
  try {
    var t = arquivo.getBlob().getDataAsString('ISO-8859-1');
    if (t && t.length > 100) return t;
  } catch(e) {}
  // Fallback OCR — requer Advanced Drive Service habilitado
  try {
    var resp = Drive.Files.copy(
      { title: 'TEMP_OCR_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
      arquivo.getId()
    );
    var texto = DocumentApp.openById(resp.id).getBody().getText();
    DriveApp.getFileById(resp.id).setTrashed(true);
    return texto;
  } catch(e) {
    Logger.log('OCR indisponível: ' + e.message + ' — habilite Serviços avançados > Drive API');
    return '';
  }
}

// ================================================================
// 6. VALIDAÇÃO DE ITEM
// ================================================================

/**
 * Retorna true se o código parece um item válido da Lista de Serviços.
 * Padrão: N.NN ou NN.NN (ex: "7.02", "07.02", "32.01", "17.09")
 * Rejeita: datas, textos, vazios, strings longas.
 * FIX BUG 1 e 3.
 */
function _isItemValido_(item) {
  if (!item) return false;
  var s = String(item).trim();
  return RE_ITEM_VALIDO_.test(s);
}

// ================================================================
// 7. LOCALIZAÇÃO DE ARQUIVOS NO DRIVE
// ================================================================
function _encontrarPastaEmpresa_(cnpj, razao) {
  var codDom = _buscarCodigoDominio_(cnpj);
  var raiz   = DriveApp.getFolderById(ID_PASTA_SALVADOR_);
  var fiscais = _listarPastasFiscais_(raiz);

  // Estratégia 1: código Domínio
  if (codDom) {
    for (var i = 0; i < fiscais.length; i++) {
      try {
        var it = fiscais[i].searchFolders("title contains '" + codDom + "-'");
        while (it.hasNext()) {
          var f = it.next();
          if (f.getName().split('-')[0].trim() === String(codDom)) return f;
        }
      } catch(e) {}
    }
  }

  // Estratégia 2: nome (3 primeiras palavras alfanuméricas)
  var trecho = razao.replace(/[^A-Z0-9\s]/gi,' ').trim()
                    .split(/\s+/).slice(0,3).join(' ')
                    .toUpperCase().replace(/'/g,"\\'");
  for (var i = 0; i < fiscais.length; i++) {
    try {
      var it = fiscais[i].searchFolders("title contains '" + trecho + "'");
      if (it.hasNext()) return it.next();
    } catch(e) {}
  }

  return null;
}

function _listarPastasFiscais_(raiz) {
  var fiscais = [];
  var it = raiz.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    if (/fiscal/i.test(f.getName())) fiscais.push(f);
  }
  return fiscais;
}

function _encontrarPastaMes_(pastaEmpresa, competencia) {
  var it = pastaEmpresa.getFoldersByName(competencia);
  return it.hasNext() ? it.next() : null;
}

function _acharSubpasta_(pasta, nome) {
  if (!pasta) return null;
  var it = pasta.getFoldersByName(nome);
  return it.hasNext() ? it.next() : null;
}

function _acharArquivo_(pastaMes, pastaPrestado, prefixos) {
  var lista = _acharArquivos_(pastaMes, pastaPrestado, prefixos);
  return lista.length ? lista[0] : null;
}

function _acharArquivos_(pastaMes, pastaPrestado, prefixos) {
  var achados = [];
  [pastaMes, pastaPrestado].forEach(function(pasta) {
    if (!pasta) return;
    var it = pasta.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      var nome = f.getName().toLowerCase();
      for (var i = 0; i < prefixos.length; i++) {
        if (nome.indexOf(prefixos[i].toLowerCase()) === 0) { achados.push(f); break; }
      }
    }
  });
  return achados;
}

// ================================================================
// 8. CÓDIGO DOMÍNIO
// ================================================================
function _buscarCodigoDominio_(cnpj) {
  var cache = CacheService.getScriptCache();
  var chave = 'dom_' + cnpj.replace(/\D/g,'');
  var cached = cache.get(chave);
  if (cached) return cached === 'NULL' ? null : cached;

  var ws = _ss_().getSheetByName(ABA_PARAMS_);
  if (!ws) { cache.put(chave,'NULL',600); return null; }

  var dados = ws.getDataRange().getValues();
  for (var i = 0; i < dados.length; i++) {
    var c0 = String(dados[i][0]).replace(/\D/g,'');
    if (c0 === cnpj.replace(/\D/g,'') && dados[i][1]) {
      var cod = String(dados[i][1]).trim();
      cache.put(chave, cod, 600);
      return cod;
    }
  }
  cache.put(chave,'NULL',600);
  return null;
}

// ================================================================
// 9. CONSOLIDAÇÃO E BANCO DE DADOS
// ================================================================

/**
 * Consolida itens duplicados. FIX BUG 3: filtra itens inválidos antes de consolidar.
 */
function _consolidarItens_(itensRaw) {
  var banco = _carregarBancoDados_();
  var mapa  = {};

  itensRaw.forEach(function(it) {
    // FIX BUG 3: pular itens que não passam na validação
    if (!_isItemValido_(it.item)) return;

    var chave = it.item;
    if (!mapa[chave]) {
      var bd = banco[chave] || {};
      mapa[chave] = {
        item:       it.item,
        descItem:   bd.descItem   || '',
        reducao:    bd.reducao    || 0,
        fundamento: bd.fundamento || '',
        totalValor: 0,
        totalISS:   0,
        qtdNotas:   0
      };
    }
    mapa[chave].totalValor += (it.valorServicos || 0);
    mapa[chave].totalISS   += (it.valorISS      || 0);
    mapa[chave].qtdNotas++;
  });

  return Object.keys(mapa).map(function(k){ return mapa[k]; });
}

function _carregarBancoDados_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('banco_dados_map');
  if (cached) return JSON.parse(cached);

  var ws   = _ss_().getSheetByName(ABA_BANCO_);
  var rows = ws.getDataRange().getValues();
  var h    = rows[1];
  var iItem = h.indexOf('Item da Lista');
  var iDesc = h.indexOf('Descrição do Item (LC 116/2003)');
  var iRed  = h.indexOf('Redução Reforma (%)');
  var iFund = h.indexOf('Fundamento legal da redução (LC 214/2025)');

  var mapa = {};
  for (var i = 2; i < rows.length; i++) {
    if (!rows[i][iItem]) continue;
    var item = String(rows[i][iItem]).trim();
    mapa[item] = {
      descItem:   String(rows[i][iDesc]  || '').trim(),
      reducao:    Number(rows[i][iRed]   || 0),
      fundamento: String(rows[i][iFund]  || '').trim()
    };
  }
  try { cache.put('banco_dados_map', JSON.stringify(mapa), 600); } catch(e) {}
  return mapa;
}

function _buscarDescItem_(item) {
  var banco = _carregarBancoDados_();
  return (banco[item] && banco[item].descItem) ? banco[item].descItem : '';
}

// ================================================================
// 10. DIVERGÊNCIAS — FIX BUG 2
// ================================================================
function atualizarDivergencias() {
  var ws   = _painel_();
  var rows = ws.getDataRange().getValues();
  var n = 0;
  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    if (!rows[i][C.FONTE_LEITURA - 1]) continue;
    _calcularDivergencia_(ws, i + 1);
    n++;
  }
  _toast_('Divergências atualizadas em ' + n + ' linha(s).');
}

/**
 * FIX BUG 2: valida que o item de referência é um código válido antes de comparar.
 * Antes, uma data era comparada com o item (sempre diverge).
 */
function _calcularDivergencia_(ws, row) {
  var vals = ws.getRange(row, 1, 1, 39).getValues()[0];
  var itemConfirmado = String(vals[C.ITEM_CONFIRMADO - 1] || '').trim();
  var itemLido       = String(vals[C.ITEM_NFSE_LIDO  - 1] || '').trim();
  var itemAtual      = String(vals[C.ITEM_ATUAL       - 1] || '').trim();

  // FIX BUG 2: só usar o item lido se ele for um código válido
  var itemRef = '';
  if (_isItemValido_(itemConfirmado)) itemRef = itemConfirmado;
  else if (_isItemValido_(itemLido))  itemRef = itemLido;

  var celDiv = ws.getRange(row, C.DIVERGENCIA);
  if (!itemRef || !itemAtual) {
    celDiv.setValue(_isItemValido_(itemLido) ? '' : '⚠️ Item NFS-e inválido — verificar leitura');
    celDiv.setBackground(_isItemValido_(itemLido) ? null : COR_AMARELO_);
    return;
  }

  var diverge = itemRef.replace(/\s/g,'') !== itemAtual.replace(/\s/g,'');
  celDiv.setValue(diverge
    ? '⚠️ Diverge: NFS-e=' + itemRef + ' | Atual=' + itemAtual
    : '✅ Confirma: ' + itemRef);
  celDiv.setBackground(diverge ? COR_AMARELO_ : COR_VERDE_);
}

// ================================================================
// 11. REPROCESSAR CBS HÍBRIDA — FIX BUG 5
// ================================================================

/**
 * FIX BUG 5: usa DAS_ATUAL (valor real do extrato PGDAS-D) como base do híbrido.
 * Versão anterior usava fat × aliq_efe_calculada, gerando excesso de ~R$5.304 na empresa exemplo.
 *
 * Fórmula correta:
 *   CBS_calculada = Faturamento × CBS_param × (1 − Redução_aplicada)
 *   DAS_Híbrido   = DAS_extrato − PIS − COFINS + CBS_calculada
 *   Alíq_Híbrida  = DAS_Híbrido ÷ Faturamento
 */
function reprocessarCBSLote() {
  var ws  = _painel_();
  var cbs = Number(_lerParametro_('CBS — alíquota cheia 2027') || 0.088);
  var rows = ws.getDataRange().getValues();
  var total = 0, alertas = [];

  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    var r    = rows[i];
    var cnpj = r[C.CNPJ - 1];
    if (!cnpj) continue;

    var fat    = Number(r[C.FATURAMENTO   - 1] || 0);
    var das    = Number(r[C.DAS_ATUAL     - 1] || 0);
    var pis    = Number(r[C.PIS           - 1] || 0);
    var cofins = Number(r[C.COFINS        - 1] || 0);
    var red    = Number(r[C.REDUCAO_APLIC - 1] || 0);

    if (!fat || !das) continue;

    var cbsCalc = fat * cbs * (1 - red);
    var dasHib  = das - pis - cofins + cbsCalc;
    var aliqHib = dasHib / fat;
    var dif     = dasHib - das;

    ws.getRange(i + 1, C.CBS_FORA).setValue(cbsCalc);
    ws.getRange(i + 1, C.DAS_HIBRIDO).setValue(dasHib);
    ws.getRange(i + 1, C.DIFERENCA).setValue(dif);
    ws.getRange(i + 1, C.ALIQ_HIBRIDA).setValue(aliqHib);
    total++;

    // Alertar se a alíquota híbrida ficou > 2× a atual (sinal de dado suspeito)
    var aliqAtual = Number(r[C.ALIQ_ATUAL - 1] || 0);
    if (aliqAtual > 0 && aliqHib / aliqAtual > 2.5) {
      alertas.push(String(r[C.RAZAO - 1]).slice(0,30) + ' (alíq hib: ' + _fmtPct_(aliqHib) + ')');
    }
  }

  var msg = 'CBS Híbrida reprocessada em ' + total + ' empresa(s). CBS usada: ' + _fmtPct_(cbs) + '.';
  if (alertas.length) msg += ' ⚠️ Verificar: ' + alertas.join('; ');
  _toast_(msg);
  if (alertas.length) Logger.log('ALÍQUOTAS SUSPEITAS:\n' + alertas.join('\n'));
}

// ================================================================
// 12. GERAÇÃO DE COMUNICADO — FIX BUG 4
// ================================================================
function gerarComunicadoLote() {
  var ws   = _painel_();
  var rows = ws.getDataRange().getValues();
  var ok = 0, err = 0;

  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    var marcado = rows[i][C.GERAR_COM - 1];
    if (marcado !== true && String(marcado).toLowerCase() !== 'true') continue;
    var row = i + 1;
    try {
      gerarComunicado_(ws, row, rows[i]);
      ok++;
      Utilities.sleep(500);
    } catch(e) {
      err++;
      Logger.log('ERRO comunicado linha ' + row + ': ' + e.message);
      ws.getRange(row, C.STATUS_COM).setValue('❌ ' + e.message.slice(0,80));
      ws.getRange(row, C.STATUS_COM).setBackground(COR_VERMELHO_);
    }
  }
  _toast_('Comunicados: ' + ok + ' OK | ' + err + ' erro(s). Ver pasta "Simulação Reforma Tributária" no Drive.');
}

/**
 * Gera e salva o PDF do comunicado.
 * FIX BUG 4:
 *  - Destino: pasta ID_PASTA_COMUNICADOS_ (Simulação Reforma Tributária)
 *  - Nome: "Comunicado_2027_NOME DA EMPRESA.pdf"
 */
function gerarComunicado_(ws, row, dadosRow) {
  var cnpj    = dadosRow[C.CNPJ     - 1];
  var razao   = dadosRow[C.RAZAO    - 1];
  var despesas = String(dadosRow[C.DESPESAS - 1] || 'IRRELEVANTE').toUpperCase().trim();
  var aliqAt  = Number(dadosRow[C.ALIQ_ATUAL   - 1] || 0);
  var aliqHib = Number(dadosRow[C.ALIQ_HIBRIDA - 1] || 0);
  var obs     = String(dadosRow[C.OBSERVACAO   - 1] || '').trim();

  var itemConfirmado = String(dadosRow[C.ITEM_CONFIRMADO - 1] || '').trim();
  var itemLido       = String(dadosRow[C.ITEM_NFSE_LIDO  - 1] || '').trim();
  // Usar item confirmado > item lido (ambos validados)
  var itemUsado = _isItemValido_(itemConfirmado) ? itemConfirmado
                 : _isItemValido_(itemLido)      ? itemLido
                 : '';

  var cbsRef = Number(_lerParametro_('CBS — alíquota cheia 2027') || 0.088);

  // Itens para o comunicado
  var itensCom = _itensParaComunicado_(cnpj, itemUsado, dadosRow);

  if (!itensCom.length) {
    ws.getRange(row, C.STATUS_COM).setValue('⚠️ Sem itens confirmados — preencher "Item NFS-e Confirmado" antes de gerar');
    ws.getRange(row, C.STATUS_COM).setBackground(COR_AMARELO_);
    return;
  }

  // ── Variáveis do template ──────────────────────────────────────
  var vars = {
    'VAR_EMPRESA':           razao,
    'VAR_CNPJ':              cnpj,
    'VAR_ALIQ_REF_CBS':      _fmtPct_(cbsRef),
    'VAR_ALIQ_EFETIVA_DENTRO': _fmtPct_(aliqAt),
    'VAR_ALIQ_EFETIVA_FORA':   _fmtPct_(aliqHib),
    'VAR_DATA':              Utilities.formatDate(new Date(),'America/Sao_Paulo','dd/MM/yyyy'),
    'VAR_BLOCO_CREDITOS':    _montarBlocoCreditos_(despesas, obs)
  };

  for (var idx = 0; idx < 2; idx++) {
    var n  = idx + 1;
    var it = itensCom[idx];
    if (it) {
      vars['VAR_ITEM_' + n]               = it.item;
      vars['VAR_DESCRICAO_ITEM_' + n]     = it.descItem;
      vars['VAR_INTERPRETACAO_ITEM_' + n] = _montarInterpretacaoItem_(it, cbsRef);
    } else {
      vars['VAR_ITEM_' + n]               = '';
      vars['VAR_DESCRICAO_ITEM_' + n]     = '';
      vars['VAR_INTERPRETACAO_ITEM_' + n] = '';
    }
  }

  // ── Copiar modelo e substituir ─────────────────────────────────
  var nomeCopia = 'Temp_Com_' + cnpj.replace(/\D/g,'') + '_' + Date.now();
  var copia = _retryOp_(function() {
    return DriveApp.getFileById(ID_MODELO_COMUNIC_).makeCopy(nomeCopia);
  }, 3);

  var body = DocumentApp.openById(copia.getId()).getBody();
  Object.keys(vars).forEach(function(v) {
    body.replaceText('<<' + v + '>>', vars[v] || '');  // padrão <<VAR>>
    body.replaceText(v, vars[v] || '');                 // fallback VAR direto
  });
  _limparParagrafosVazios_(body);

  DocumentApp.openById(copia.getId()).saveAndClose();

  // ── Converter para PDF ─────────────────────────────────────────
  var pdfBlob = _retryOp_(function() {
    return DriveApp.getFileById(copia.getId()).getAs('application/pdf');
  }, 3);

  // FIX BUG 4: nome "Comunicado_2027_NOME DA EMPRESA.pdf"
  var nomeRazaoLimpo = razao.replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 60);
  var nomePDF = 'Comunicado_2027_' + nomeRazaoLimpo + '.pdf';
  pdfBlob.setName(nomePDF);

  // FIX BUG 4: salvar na pasta correta (Simulação Reforma Tributária)
  var pastaDestino = DriveApp.getFolderById(ID_PASTA_COMUNICADOS_);
  // Sobrescrever se já existir (evitar duplicatas)
  var existentes = pastaDestino.getFilesByName(nomePDF);
  while (existentes.hasNext()) existentes.next().setTrashed(true);
  pastaDestino.createFile(pdfBlob);

  // Remover cópia temporária do Docs
  copia.setTrashed(true);

  // Atualizar status
  var dataHora = Utilities.formatDate(new Date(),'America/Sao_Paulo','dd/MM/yyyy HH:mm');
  ws.getRange(row, C.STATUS_COM).setValue('✅ ' + dataHora);
  ws.getRange(row, C.STATUS_COM).setBackground(COR_VERDE_);
  ws.getRange(row, C.GERAR_COM).setValue(false);

  Logger.log('PDF gerado: ' + nomePDF + ' → ' + pastaDestino.getName());
}

// ================================================================
// 13. MONTAGEM DO CONTEÚDO DO COMUNICADO
// ================================================================
function _itensParaComunicado_(cnpj, itemUsado, dadosRow) {
  var corr  = _carregarCorrelacaoCNPJ_(cnpj);
  var banco = _carregarBancoDados_();

  var itensAlvo = [];
  if (itemUsado) {
    itensAlvo.push(itemUsado);
  } else {
    var pesq = String(dadosRow[C.ITEM_PESQUISADO - 1] || '').trim();
    var atu  = String(dadosRow[C.ITEM_ATUAL      - 1] || '').trim();
    if (pesq) pesq.split(/[,|;]/).forEach(function(p){
      var pp = p.trim();
      if (_isItemValido_(pp) && itensAlvo.indexOf(pp) < 0) itensAlvo.push(pp);
    });
    if (_isItemValido_(atu) && itensAlvo.indexOf(atu) < 0) itensAlvo.push(atu);
  }

  return itensAlvo.filter(Boolean).map(function(item) {
    var dc = corr[item]  || {};
    var db = banco[item] || {};
    return {
      item:         item,
      descItem:     dc.descItem    || db.descItem  || '',
      reducaoAplic: dc.reducaoAplic !== undefined ? dc.reducaoAplic : (db.reducao || 0),
      reducaoOrig:  dc.reducaoOrig  || db.reducao  || 0,
      art127:       dc.art127       || '',
      fundamento:   dc.fundamento   || db.fundamento || ''
    };
  });
}

function _carregarCorrelacaoCNPJ_(cnpj) {
  var ws   = _ss_().getSheetByName(ABA_COR_);
  var rows = ws.getDataRange().getValues();
  var h    = rows[1];
  var iCnpj = h.indexOf('CNPJ');
  var iItem = h.indexOf('Item da Lista');
  var iDesc = h.indexOf('Descrição Item');
  var iRA   = h.indexOf('Redução Aplicada');
  var iRO   = h.indexOf('Redução Original');
  var iArt  = h.indexOf('Art. 127');
  var iFund = h.indexOf('Fundamento LC 214');

  var mapa = {};
  for (var i = 2; i < rows.length; i++) {
    if (rows[i][iCnpj] !== cnpj) continue;
    var item = String(rows[i][iItem] || '').trim();
    if (!item) continue;
    mapa[item] = {
      descItem:    String(rows[i][iDesc] || ''),
      reducaoAplic: Number(rows[i][iRA]  || 0),
      reducaoOrig:  Number(rows[i][iRO]  || 0),
      art127:       String(rows[i][iArt] || ''),
      fundamento:   String(rows[i][iFund]|| '')
    };
  }
  return mapa;
}

/**
 * Texto de interpretação por item para o comunicado.
 * Regra: redução=0 sem art.127 → "não se aplica"
 *        redução zerada por art.127 → texto explicativo do critério
 *        redução positiva → cita percentual e alíquota calculada
 */
function _montarInterpretacaoItem_(it, cbsRef) {
  var redOrig  = it.reducaoOrig  || 0;
  var redAplic = it.reducaoAplic || 0;
  var art127Zerou = it.art127 && it.art127.indexOf('ZERADA') !== -1;

  if (redOrig === 0 && !art127Zerou) {
    return 'Para este serviço, não há previsão de redução da CBS na legislação vigente. '
         + 'A simulação utiliza a alíquota de referência de ' + _fmtPct_(cbsRef) + '.';
  }

  if (art127Zerou || (redOrig > 0 && redAplic === 0)) {
    return 'Embora exista previsão de redução da CBS para este serviço (' + _fmtPct_(redOrig) + '), '
         + 'os requisitos cumulativos do art. 127 da LC 214/2025 não foram atendidos — '
         + 'a empresa possui atividades diversas da habilitação profissional dos sócios. '
         + 'Por isso, a redução não foi aplicada na simulação e foi utilizada a alíquota '
         + 'de referência de ' + _fmtPct_(cbsRef) + '.';
  }

  var cbsCalc = cbsRef * (1 - redAplic);
  return 'Este serviço possui previsão de redução da CBS de ' + _fmtPct_(redAplic)
       + ' (Regime Específico da LC 214/2025). '
       + 'Para a simulação, foi considerada a alíquota estimada de ' + _fmtPct_(cbsCalc) + '.';
}

function _montarBlocoCreditos_(tipo, obs) {
  if (!tipo || tipo.indexOf('IRRELEVANTE') !== -1) {
    return 'Os custos e despesas analisados não representam um volume relevante de créditos de CBS '
         + 'para esta simulação. Por isso, o impacto desses créditos foi considerado irrelevante '
         + 'na comparação entre os cenários.';
  }
  var reValor = /R\$\s*([\d.,]+)/i;
  var mObs  = reValor.exec(obs  || '');
  var mTipo = reValor.exec(tipo || '');
  var valorStr = (mObs || mTipo) ? 'R$ ' + ((mObs||mTipo)[1]) + ' mensais' : null;

  if (tipo.indexOf('CONFIRMADO') !== -1 && valorStr) {
    return 'Foram consideradas aproximadamente ' + valorStr + ' em despesas que podem gerar '
         + 'créditos de CBS. Com base nessa premissa, os créditos projetados foram considerados '
         + 'na simulação do Simples Híbrido. O valor efetivamente aproveitado poderá ser menor, '
         + 'dependendo da tributação dos fornecedores.';
  }
  return 'A empresa possui despesas que podem gerar créditos de CBS e esse fator foi considerado '
       + 'na simulação. O aproveitamento efetivo dependerá da tributação e das condições '
       + 'aplicáveis às operações com cada fornecedor.';
}

function _limparParagrafosVazios_(body) {
  var paras = body.getParagraphs();
  for (var i = paras.length - 1; i >= 0; i--) {
    var txt = paras[i].getText().trim();
    if (txt === '' || /^<<VAR_ITEM_2/.test(txt) || /^VAR_ITEM_2/.test(txt)) {
      try { paras[i].removeFromParent(); } catch(e) {}
    }
  }
}

// ================================================================
// 14. PAINEL ANÁLISE POR CNPJ
// ================================================================
function carregarPainelCNPJ_(cnpj) {
  cnpj = cnpj.replace(/\s/g,'');
  var ws   = _painel_();
  var rows = ws.getDataRange().getValues();
  var dadosRow = null;

  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    if (String(rows[i][C.CNPJ - 1]).replace(/\D/g,'') === cnpj.replace(/\D/g,'')) {
      dadosRow = rows[i]; break;
    }
  }

  var wsA = _ss_().getSheetByName(ABA_ANALISE_);
  if (!wsA) { _toast_('Aba "' + ABA_ANALISE_ + '" não encontrada.'); return; }

  wsA.getRange(2, 2).setValue(cnpj);

  if (!dadosRow) {
    wsA.getRange(4, 2).setValue('❌ CNPJ não encontrado no Painel Geral');
    return;
  }

  var campos = {
    'CNPJ':                                  dadosRow[C.CNPJ           - 1],
    'Razão Social':                          dadosRow[C.RAZAO          - 1],
    'Data de Abertura':                      dadosRow[C.ABERTURA       - 1],
    'Segmento':                              dadosRow[C.SEGMENTO       - 1],
    'Tipo CNAEs (ÚNICO/VÁRIOS)':             dadosRow[C.CNAES_TIPO     - 1],
    'Despesas (tipo crédito)':               dadosRow[C.DESPESAS       - 1],
    'Anexo do Simples':                      dadosRow[C.ANEXO          - 1],
    'Faixa RBT12':                           dadosRow[C.FAIXA          - 1],
    'RBT12 (R$)':                            dadosRow[C.RBT12          - 1],
    'Faturamento — Período de Apuração (R$)':dadosRow[C.FATURAMENTO    - 1],
    'Alíquota Efetiva Atual (%)':            dadosRow[C.ALIQ_ATUAL     - 1],
    'CNAE Principal (7 dígitos)':            dadosRow[C.CNAE_PRINCIPAL - 1],
    'Item Predominante Atual':               dadosRow[C.ITEM_ATUAL     - 1],
    'Descrição do Item Atual':               dadosRow[C.DESC_ITEM_ATUAL- 1],
    'NBS Vinculado':                         dadosRow[C.NBS            - 1],
    'Item da Lista Pesquisado (VÁRIOS)':     dadosRow[C.ITEM_PESQUISADO- 1],
    'Item da NFS-e Lido no Portal':          dadosRow[C.ITEM_NFSE_LIDO - 1],
    'Item da NFS-e Confirmado':              dadosRow[C.ITEM_CONFIRMADO- 1],
    'Divergência de Item?':                  dadosRow[C.DIVERGENCIA    - 1],
    'Redução Aplicada (%)':                  dadosRow[C.REDUCAO_APLIC  - 1],
    'Redução Original do Item (%)':          dadosRow[C.REDUCAO_ORIG   - 1],
    'Requisito Art. 127 OK?':                dadosRow[C.ART127         - 1],
    'Valor no DAS (R$)':                     dadosRow[C.DAS_ATUAL      - 1],
    'PIS no DAS (R$)':                       dadosRow[C.PIS            - 1],
    'COFINS no DAS (R$)':                    dadosRow[C.COFINS         - 1],
    'DAS sem PIS/COFINS (R$)':               dadosRow[C.DAS_ATUAL-1] - dadosRow[C.PIS-1] - dadosRow[C.COFINS-1],
    'CBS fora do DAS — Híbrido 2027 (R$)':  dadosRow[C.CBS_FORA       - 1],
    'DAS Híbrido 2027 (R$)':                 dadosRow[C.DAS_HIBRIDO   - 1],
    'Diferença Híbrido × DAS Atual (R$)':    dadosRow[C.DIFERENCA     - 1],
    'Alíquota Efetiva Atual (%) ':           dadosRow[C.ALIQ_ATUAL     - 1],
    'Alíquota Efetiva Híbrida 2027 (%)':     dadosRow[C.ALIQ_HIBRIDA  - 1],
    'Competência lida':                      dadosRow[C.COMPETENCIA    - 1],
    'Fonte da leitura':                      dadosRow[C.FONTE_LEITURA  - 1],
    'Qtd de notas prestadas':                dadosRow[C.QTD_NOTAS      - 1],
    'Valor total NFS-e (R$)':                dadosRow[C.VALOR_NFSE     - 1],
    'ISS total NFS-e (R$)':                  dadosRow[C.ISS_NFSE       - 1],
    'Alerta de leitura':                     dadosRow[C.ALERTA_LEITURA - 1],
    'Status do Comunicado':                  dadosRow[C.STATUS_COM     - 1],
    'Observação':                            dadosRow[C.OBSERVACAO     - 1]
  };

  var dadosA = wsA.getDataRange().getValues();
  for (var j = 0; j < dadosA.length; j++) {
    var rotulo = String(dadosA[j][0] || '').trim();
    if (campos.hasOwnProperty(rotulo)) wsA.getRange(j + 1, 2).setValue(campos[rotulo]);
  }
  _toast_('Painel carregado: ' + dadosRow[C.RAZAO - 1]);
}

// ================================================================
// 15. DIAGNÓSTICO
// ================================================================
function diagTestarLeitura() {
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  if (row < L_DADOS_) { _toast_('Posicione em uma linha de empresa.'); return; }
  var ws  = _painel_();
  var d   = _lerLinha_(ws, row);

  var pasta = _encontrarPastaEmpresa_(d[C.CNPJ - 1], d[C.RAZAO - 1]);
  if (!pasta) { Logger.log('PASTA NÃO ENCONTRADA: ' + d[C.CNPJ - 1]); _toast_('Pasta não encontrada — ver Log'); return; }

  var compet   = d[C.COMPETENCIA - 1] || '082026';
  var pastaMes = _encontrarPastaMes_(pasta, compet);
  if (!pastaMes) { Logger.log('MÊS NÃO ENCONTRADO: ' + compet); _toast_('Pasta do mês não encontrada — ver Log'); return; }

  var result = _extrairItensNFSe_(pastaMes);
  Logger.log('DIAGNÓSTICO — ' + d[C.CNPJ - 1] + ' / ' + d[C.RAZAO - 1]);
  Logger.log(JSON.stringify(result, null, 2));
  _toast_('Diagnóstico concluído. Ver Log (Ctrl+Enter).');
}

function diagListarPendentes() {
  var ws   = _painel_();
  var rows = ws.getDataRange().getValues();
  var pend = [];
  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    if (!rows[i][C.FONTE_LEITURA - 1] ||
        String(rows[i][C.ALERTA_LEITURA - 1]).indexOf('Pendente') !== -1) {
      pend.push('L' + (i+1) + ': ' + rows[i][C.CNPJ - 1] + ' | ' + rows[i][C.RAZAO - 1]);
    }
  }
  Logger.log('PENDENTES NFS-e (' + pend.length + '):\n' + pend.join('\n'));
  _toast_(pend.length + ' empresa(s) pendente(s). Ver Log.');
}

function diagDivergencias() {
  var ws   = _painel_();
  var rows = ws.getDataRange().getValues();
  var divs = [], invalidos = [];
  for (var i = L_DADOS_ - 1; i < rows.length; i++) {
    var div  = String(rows[i][C.DIVERGENCIA    - 1] || '');
    var item = String(rows[i][C.ITEM_NFSE_LIDO - 1] || '');
    if (div.indexOf('⚠️ Diverge')    !== -1) divs.push('L'+(i+1)+': '+rows[i][C.CNPJ-1]+' → '+div);
    if (div.indexOf('inválido') !== -1 || (!_isItemValido_(item) && rows[i][C.FONTE_LEITURA-1])) {
      invalidos.push('L'+(i+1)+': '+rows[i][C.CNPJ-1]+' item_lido="'+item+'"');
    }
  }
  Logger.log('DIVERGÊNCIAS (' + divs.length + '):\n' + divs.join('\n'));
  Logger.log('ITENS INVÁLIDOS LIDOS (' + invalidos.length + '):\n' + invalidos.join('\n'));
  _toast_(divs.length + ' divergência(s) | ' + invalidos.length + ' item(s) inválido(s). Ver Log.');
}

function diagDumpEmpresa() {
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  if (row < L_DADOS_) { _toast_('Posicione em uma linha de empresa.'); return; }
  var ws = _painel_();
  var d  = _lerLinha_(ws, row);
  Logger.log('=== DUMP L' + row + ' ===');
  Object.keys(C).forEach(function(k){ Logger.log(k + ' [' + C[k] + ']: ' + d[C[k]-1]); });
  Logger.log('=== CORRELAÇÃO ===');
  Logger.log(JSON.stringify(_carregarCorrelacaoCNPJ_(d[C.CNPJ - 1]), null, 2));
  _toast_('Dump no Log.');
}

/**
 * D5 — Dump do CSV bruto para depurar leitura de item (BUG 1).
 * Loga o cabeçalho e as primeiras 5 linhas do CSV encontrado na pasta da empresa.
 */
function diagDumpCsvBruto() {
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  if (row < L_DADOS_) { _toast_('Posicione em uma linha de empresa.'); return; }
  var ws = _painel_();
  var d  = _lerLinha_(ws, row);

  var pasta = _encontrarPastaEmpresa_(d[C.CNPJ - 1], d[C.RAZAO - 1]);
  if (!pasta) { _toast_('Pasta não encontrada.'); return; }
  var compet  = d[C.COMPETENCIA - 1] || '082026';
  var pastaMes = _encontrarPastaMes_(pasta, compet);
  if (!pastaMes) { _toast_('Mês não encontrado.'); return; }

  var pastaPrestado = _acharSubpasta_(pastaMes, 'PRESTADO');
  var csv = _acharArquivo_(pastaMes, pastaPrestado, ['nfse_csv_prestado']);
  if (!csv) { _toast_('CSV não encontrado.'); return; }

  var texto;
  try { texto = csv.getBlob().getDataAsString('ISO-8859-1'); } catch(e) {
    texto = csv.getBlob().getDataAsString('UTF-8');
  }
  var primeiraLinha = texto.split('\n')[0] || '';
  var delim = (primeiraLinha.split(';').length > primeiraLinha.split(',').length) ? ';' : ',';
  var linhas = Utilities.parseCsv(texto, delim);

  Logger.log('CSV: ' + csv.getName() + ' | delimitador: "' + delim + '" | ' + linhas.length + ' linhas');
  Logger.log('CABEÇALHO:\n' + JSON.stringify(linhas[0]));
  for (var i = 1; i <= Math.min(5, linhas.length - 1); i++) {
    Logger.log('LINHA ' + i + ':\n' + JSON.stringify(linhas[i]));
  }
  _toast_('CSV dump no Log (Ctrl+Enter).');
}

// ================================================================
// 16. UTILITÁRIOS
// ================================================================
function _ss_()     { return SpreadsheetApp.getActiveSpreadsheet(); }
function _painel_() { return _ss_().getSheetByName(ABA_PAINEL_); }

function _lerLinha_(ws, row) {
  return ws.getRange(row, 1, 1, 39).getValues()[0];
}

function _lerTodasLinhas_(ws) {
  return ws.getDataRange().getValues().slice(L_DADOS_ - 1);
}

function _encontrarLinha_(ws, cnpj) {
  var vals = ws.getDataRange().getValues();
  for (var i = L_DADOS_ - 1; i < vals.length; i++) {
    if (String(vals[i][C.CNPJ - 1]).replace(/\D/g,'') === String(cnpj).replace(/\D/g,'')) return i + 1;
  }
  return null;
}

function _lerParametro_(chave) {
  var ws = _ss_().getSheetByName(ABA_PARAMS_);
  if (!ws) return null;
  var vals = ws.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === chave) return vals[i][1];
  }
  return null;
}

/** 0.088 → "8,80%" */
function _fmtPct_(v) {
  if (v === null || v === undefined || v === '') return '—';
  return (Number(v) * 100).toFixed(2).replace('.', ',') + '%';
}

/** "703" → "7.03" | "3201" → "32.01" | "07.02" → "07.02" */
function _formatarItem_(codigo) {
  if (!codigo) return '';
  var s = String(codigo).trim();
  if (s.indexOf('.') !== -1) return s.replace(/^0+(\d)/, '$1').replace(/^(\d)\./, '0$1.') || s;
  if (s.length < 3) return s;
  var raw = s.slice(0, s.length - 2) + '.' + s.slice(-2);
  // normalizar: "7.02" não "07.02" (ambas são válidas, mas manter consistência com LC 116)
  return raw;
}

/** "18.000,00" → 18000 */
function _toNum_(v) {
  if (!v && v !== 0) return 0;
  return parseFloat(String(v).replace(/\./g,'').replace(',','.')) || 0;
}

function _toast_(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '⚙️ Estudo Complexo', 7);
}

function _retryOp_(fn, maxTentativas) {
  maxTentativas = maxTentativas || 4;
  var ultima;
  for (var t = 0; t < maxTentativas; t++) {
    try { return fn(); } catch(e) {
      ultima = e;
      var msg = e.message.toLowerCase();
      if (ERROS_DEFIN_.some(function(p){ return msg.indexOf(p) !== -1; })) throw e;
      if (t < RETRY_DELAYS_.length) Utilities.sleep(RETRY_DELAYS_[t]);
    }
  }
  throw ultima;
}