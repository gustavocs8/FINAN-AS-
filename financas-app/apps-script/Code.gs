/**
 * Finanças do Gustavo — banco de dados do app.
 *
 * Este script guarda os dados do app numa planilha própria, criada por ele
 * mesmo na primeira vez que o app se conecta. Ninguém precisa abrir essa
 * planilha: tudo é feito pelo app.
 *
 * Publicar: Implantar › Nova implantação › App da Web
 *   Executar como: Eu · Quem pode acessar: Qualquer pessoa
 */

// Chave de acesso do app. Quem tiver a chave consegue ler e lançar.
const CHAVE = 'LaFgQYI7JAwX4p28';

const NOME_PLANILHA = 'Finanças do Gustavo (dados do app)';
const COLS = ['id', 'data', 'descricao', 'tipo', 'categoria', 'subcategoria', 'conta', 'forma', 'natureza',
              'valor', 'status', 'lancado_por', 'grupo', 'parcela', 'criado_em', 'atualizado_em'];
const CAMPOS = ['id', 'data', 'desc', 'tipo', 'cat', 'sub', 'conta', 'forma', 'nat',
                'valor', 'status', 'autor', 'grupo', 'parcela', 'criado', 'atualizado'];

// Configuração inicial, com as listas que já eram usadas na planilha de controle.
const CONFIG_INICIAL = {
  categorias: ['Moradia', 'Alimentação', 'Transporte', 'Combustível', 'Saúde', 'Maria Helena', 'Educação',
               'Lazer', 'Assinaturas', 'Impostos', 'Empréstimos', 'Investimentos', 'Outros'],
  categoriasReceita: ['Salário', 'Outras receitas'],
  subcategorias: ['Aluguel/Financiamento', 'Energia', 'Água', 'Internet', 'Supermercado', 'Restaurante',
                  'Combustível', 'Manutenção', 'Farmácia', 'Consultas', 'Escola/Creche', 'Roupas', 'Viagem', 'Outros'],
  contas: ['Conta principal', 'Dinheiro'],
  cartoes: [
    { nome: 'Ourocard Visa Infinite 9666', limite: 42447, fechamento: 15, vencimento: 25 },
    { nome: 'Nubank', limite: 16400, fechamento: 1, vencimento: 9 },
    { nome: 'Caixa 7845', limite: 4600, fechamento: 7, vencimento: 17 }
  ],
  formas: ['Pix', 'Débito', 'Dinheiro', 'Boleto', 'Transferência', 'Cartão de crédito'],
  status: ['Pago', 'Previsto', 'Recebido', 'Atrasado', 'Cancelado'],
  naturezas: ['Variável', 'Fixa'],
  orcamento: {}
};

/* ---------- entrada ---------- */

function doGet(e) {
  return responder_(() => {
    checarChave_(e.parameter.chave);
    if (e.parameter.acao === 'ping') { banco_(); return { ok: true }; }
    if (e.parameter.acao === 'dados') return dados_(String(e.parameter.ano || new Date().getFullYear()));
    throw erro_('acao_invalida', 'Ação desconhecida.');
  });
}

function doPost(e) {
  return responder_(() => {
    const req = JSON.parse((e.postData && e.postData.contents) || '{}');
    checarChave_(req.chave);
    const trava = LockService.getScriptLock();
    trava.waitLock(25000);
    try {
      if (req.acao === 'salvar') return salvar_(req.lancs || []);
      if (req.acao === 'apagar') return apagar_(req.ids || []);
      if (req.acao === 'config') return salvarConfig_(req.config);
      throw erro_('acao_invalida', 'Ação desconhecida.');
    } finally {
      SpreadsheetApp.flush();
      trava.releaseLock();
    }
  });
}

/* ---------- planilha de dados ---------- */

function banco_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BANCO_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* apagada: cria outra */ }
  }
  const trava = LockService.getScriptLock();
  const jaTinha = trava.hasLock();
  if (!jaTinha) trava.waitLock(25000);
  try {
    const id2 = props.getProperty('BANCO_ID');
    if (id2 && id2 !== id) return SpreadsheetApp.openById(id2);
    const ss = SpreadsheetApp.create(NOME_PLANILHA);
    const sh = ss.getSheets()[0].setName('Lancamentos');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');   // id e data como texto
    sh.getRange('M:N').setNumberFormat('@');   // grupo e parcela como texto
    sh.getRange('J:J').setNumberFormat('#,##0.00');
    const cfg = ss.insertSheet('Config');
    cfg.getRange('A1:B1').setValues([['chave', 'valor']]).setFontWeight('bold');
    cfg.getRange('A2:B2').setValues([['config', JSON.stringify(CONFIG_INICIAL)]]);
    cfg.getRange('A4').setValue('Esta planilha é o banco de dados do app Finanças do Gustavo. Não altere a estrutura.');
    props.setProperty('BANCO_ID', ss.getId());
    return ss;
  } finally {
    if (!jaTinha) trava.releaseLock();
  }
}

function lerConfig_(ss) {
  try { return Object.assign({}, CONFIG_INICIAL, JSON.parse(ss.getSheetByName('Config').getRange('B2').getValue())); }
  catch (e) { return CONFIG_INICIAL; }
}

function lerLinhas_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COLS.length).getValues().map((r, i) => {
    const o = { _linha: i + 2 };
    CAMPOS.forEach((c, j) => { o[c] = r[j]; });
    o.id = String(o.id); o.data = String(o.data); o.valor = Number(o.valor) || 0;
    o.grupo = String(o.grupo || ''); o.parcela = String(o.parcela || '');
    return o;
  }).filter(o => o.id);
}

const publico_ = o => ({ id: o.id, data: o.data, desc: String(o.desc), tipo: o.tipo, cat: o.cat, sub: o.sub,
  conta: o.conta, forma: o.forma, nat: o.nat, valor: o.valor, status: o.status, autor: o.autor,
  grupo: o.grupo, parcela: o.parcela });

/* ---------- ações ---------- */

function dados_(ano) {
  const ss = banco_();
  const linhas = lerLinhas_(ss.getSheetByName('Lancamentos'));
  return {
    ok: true, agora: new Date().toISOString(), ano,
    config: lerConfig_(ss),
    lanc: linhas.filter(o => o.data.slice(0, 4) === ano).map(publico_),
    parcelados: linhas.filter(o => o.grupo).map(publico_)
  };
}

function salvar_(lancs) {
  const ss = banco_();
  const sh = ss.getSheetByName('Lancamentos');
  const porId = {};
  lerLinhas_(sh).forEach(o => { porId[o.id] = o; });
  const agora = new Date().toISOString();
  const novos = [];
  lancs.forEach(l => {
    validar_(l);
    const atual = porId[l.id];
    const linha = CAMPOS.map(c => {
      if (c === 'criado') return atual ? atual.criado : agora;
      if (c === 'atualizado') return agora;
      if (c === 'valor') return Number(l.valor);
      return l[c] == null ? '' : String(l[c]);
    });
    if (atual) sh.getRange(atual._linha, 1, 1, COLS.length).setValues([linha]);
    else { novos.push(linha); porId[l.id] = { _linha: -1 }; }
  });
  if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, COLS.length).setValues(novos);
  return { ok: true, salvos: lancs.length };
}

function apagar_(ids) {
  const sh = banco_().getSheetByName('Lancamentos');
  const alvo = {};
  ids.forEach(id => { alvo[String(id)] = true; });
  lerLinhas_(sh).filter(o => alvo[o.id]).map(o => o._linha)
    .sort((a, b) => b - a).forEach(l => sh.deleteRow(l));
  return { ok: true };
}

function salvarConfig_(cfg) {
  if (!cfg || !Array.isArray(cfg.categorias) || !Array.isArray(cfg.contas) || !Array.isArray(cfg.cartoes)) {
    throw erro_('dados_invalidos', 'Configuração incompleta.');
  }
  banco_().getSheetByName('Config').getRange('B2').setValue(JSON.stringify(cfg));
  return { ok: true };
}

function validar_(l) {
  if (!l || !l.id) throw erro_('dados_invalidos', 'Lançamento sem identificador.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(l.data)) throw erro_('dados_invalidos', 'Data inválida.');
  if (!l.desc) throw erro_('dados_invalidos', 'Descrição vazia.');
  if (!(Number(l.valor) > 0)) throw erro_('dados_invalidos', 'Valor inválido.');
  if (l.tipo !== 'Despesa' && l.tipo !== 'Receita') throw erro_('dados_invalidos', 'Tipo inválido.');
}

/* ---------- utilitários ---------- */

function checarChave_(c) { if (c !== CHAVE) throw erro_('chave_invalida', 'Chave de acesso incorreta.'); }
function erro_(codigo, mensagem) { const e = new Error(mensagem); e.codigo = codigo; return e; }
function responder_(fn) {
  let out;
  try { out = fn(); }
  catch (e) { out = { ok: false, codigo: e.codigo || 'erro', mensagem: e.message }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Rode esta função uma vez pelo editor para autorizar o script antes de publicar.
function autorizar() {
  const ss = banco_();
  Logger.log('Banco de dados pronto: ' + ss.getUrl());
}
