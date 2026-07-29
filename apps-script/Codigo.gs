/**
 * API do One Page — Google Apps Script vinculado à planilha DRE Banco.
 * Publicar como Web App (Implantar > Nova implantação > App da Web).
 * Mesmo modelo usado no Forecast 2026 e no Painel de Tarefas.
 *
 * Abas esperadas na planilha:
 *   "2026"      -> Mês Ano | Loja | Canal | DRE | -2026     | Orçado -2026
 *   "2025"      -> Mês Ano | Loja | Canal | DRE | -2025     | Orçado -2025
 *   "Atacado"   -> Período | DRE | Real | Orçado
 *                  (base Lynkz Gerencial Consolidado, já totalizada — NÃO é mensal
 *                  como as abas 2026/2025). Formato ACUMULATIVO: a cada fechamento
 *                  (6M, 7M, 8M...) cola-se um novo bloco de 4 linhas (uma por métrica
 *                  de METRICAS) embaixo das anteriores, marcado com o Período daquele
 *                  fechamento (ex: "6M/2026", "6M/2025", "7M/2026"...). Nada é
 *                  sobrescrito — o histórico completo fica sempre na planilha.
 *   "USUARIOS"  -> nome | email | senha | papel | ativo
 */

const SPREADSHEET_ID = '1qYlMyLFWbtqkGw7ingnFFAilDDYxDwzun0TGs8lBMr0';
const SHEET_2026 = '2026';
const SHEET_2025 = '2025';
const SHEET_ATACADO = 'Atacado';
const SHEET_USERS = 'USUARIOS';

// Chave de acesso — precisa ser idêntica à do index.html (front-end).
const CHAVE_ACESSO = '1130e0d6d3f33367e5d75519fd5827ba887bc46524426a33';

// Métricas de topo usadas no One Page (nomes exatamente como aparecem na
// coluna DRE, sem os espaços de indentação).
const METRICAS = ['Receita Bruta', 'Receita Líquida', 'Lucro Bruto Ajustado', 'EBITDA'];

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.chave !== CHAVE_ACESSO) {
      return _json({ status: 'erro', message: 'Acesso negado.' });
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const dreBanco = []
      .concat(_lerDreBanco(ss, SHEET_2026, 2026))
      .concat(_lerDreBanco(ss, SHEET_2025, 2025));

    const atacado = _lerAtacado(ss, SHEET_ATACADO);
    const usuarios = _lerUsuarios(ss, SHEET_USERS);

    return _json({ status: 'ok', dreBanco: dreBanco, atacado: atacado, usuarios: usuarios });
  } catch (err) {
    return _json({ status: 'erro', message: err.message });
  }
}

// Lê a aba anual da DRE Banco e devolve só os totais das linhas de topo,
// já somados por Canal + Mês (em vez das milhares de linhas de detalhe por
// loja/conta contábil) — mantém o JSON pequeno e rápido de carregar.
function _lerDreBanco(ss, sheetName, ano) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const totals = {}; // chave: mes|canal|metrica -> soma

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dataMes = row[0];
    const canal = row[2];
    const dre = row[3];
    const valor = row[4];
    if (!dataMes || !dre) continue;
    if (String(dre).charAt(0) === ' ') continue; // linha de detalhe, ignora
    const label = String(dre).trim();
    if (METRICAS.indexOf(label) === -1) continue;

    const mes = (dataMes instanceof Date) ? (dataMes.getMonth() + 1) : Number(String(dataMes).slice(5, 7));
    const chave = mes + '|' + canal + '|' + label;
    totals[chave] = (totals[chave] || 0) + (Number(valor) || 0);
  }

  const out = [];
  Object.keys(totals).forEach(function (chave) {
    const partes = chave.split('|');
    out.push({
      ano: ano,
      mes: Number(partes[0]),
      canal: partes[1],
      metrica: partes[2],
      valor: totals[chave],
    });
  });
  return out;
}

// Lê a aba Atacado (Lynkz Gerencial Consolidado), formato acumulativo:
// Período | DRE | Real | Orçado. Devolve uma lista simples de registros;
// o front-end filtra pelo texto do Período (ex: "6M/2026") selecionado.
function _lerAtacado(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!data.length) return [];

  const header = data[0].map(function (h) {
    return String(h || '').trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
  });
  const iPeriodo = header.indexOf('periodo') >= 0 ? header.indexOf('periodo') : header.indexOf('período');
  const iDre = header.indexOf('dre');
  const iReal = header.indexOf('real');
  const iOrcado = header.indexOf('orcado') >= 0 ? header.indexOf('orcado') : header.indexOf('orçado');

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[iPeriodo] || !row[iDre]) continue;
    out.push({
      periodo: String(row[iPeriodo]).trim(),
      dre: String(row[iDre]).trim(),
      real: Number(row[iReal]) || 0,
      orcado: iOrcado >= 0 ? (Number(row[iOrcado]) || 0) : 0,
    });
  }
  return out;
}

// Mesma lógica de leitura flexível de usuários do Forecast 2026.
function _lerUsuarios(ss, sheetName) {
  const usuarios = [];
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return usuarios;
  const data = sh.getDataRange().getValues();
  if (!data.length) return usuarios;

  const header = data[0].map(function (x) {
    return String(x).trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]/g, '');
  });
  const col = function (keywords) {
    for (let i = 0; i < header.length; i++) {
      for (let k = 0; k < keywords.length; k++) {
        if (header[i].indexOf(keywords[k]) >= 0) return i;
      }
    }
    return -1;
  };
  const iNome = col(['nome']);
  const iEmail = col(['email', 'mail']);
  const iSenha = col(['senha']);
  const iPapel = col(['papel', 'perfil', 'role']);
  const iAtivo = col(['ativo', 'status']);

  data.slice(1).forEach(function (row) {
    const email = iEmail >= 0 ? String(row[iEmail] || '').trim() : '';
    if (!email) return;
    usuarios.push({
      nome: iNome >= 0 ? String(row[iNome] || '').trim() : email,
      email: email,
      senha: iSenha >= 0 ? String(row[iSenha] || '').trim() : '',
      papel: iPapel >= 0 ? String(row[iPapel] || 'leitor').trim().toLowerCase() : 'leitor',
      ativo: iAtivo >= 0 ? String(row[iAtivo] || 'sim').trim().toLowerCase() : 'sim',
    });
  });
  return usuarios;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
