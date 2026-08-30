/**
 * CLC - RECEBIMENTO DE PRÉ-FILIAÇÕES
 * Este projeto deve ser criado dentro da MESMA planilha do Google Planilhas.
 * Aba de destino: "Cadastro"
 *
 * Depois de publicar como Web App, copie a URL /exec e cole em
 * FILIACAO_API_URL no arquivo index.html.
 */

const ABA_DESTINO = 'Cadastro';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');

    const required = ['nome','sexo','cpf','nascimento','telefone','cidade','estado','rua','bairro'];
    for (const key of required) {
      if (!String(data[key] || '').trim()) {
        return json_({ ok:false, message:'Campo obrigatório ausente: ' + key });
      }
    }

    if (String(data.cpf).replace(/\D/g,'').length !== 11) {
      return json_({ ok:false, message:'CPF inválido.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ABA_DESTINO);
    if (!sheet) throw new Error('A aba "' + ABA_DESTINO + '" não foi encontrada.');

    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1,1,1,lastColumn).getValues()[0]
      .map(v => normalize_(v));

    // Localiza as colunas mesmo que o nome usado no cabeçalho varie.
    const aliases = {
      nome: ['nome','nome completo','filiado','lacador','laçador'],
      sexo: ['sexo','genero','gênero'],
      cpf: ['cpf'],
      nascimento: ['data de nascimento','nascimento','dt nascimento'],
      telefone: ['telefone','celular','whatsapp','fone'],
      cidade: ['cidade','municipio','município'],
      estado: ['estado','uf'],
      rua: ['rua','endereco','endereço'],
      bairro: ['bairro'],
      status: ['status','situação','situacao'],
      origem: ['origem'],
      termos: ['termos','aceite dos termos','aceite'],
      solicitacao: ['data solicitacao','data solicitação','solicitacao','solicitação']
    };

    // Se algum campo novo não existir, cria a coluna sem apagar o banco atual.
    const fieldNames = Object.keys(aliases);
    fieldNames.forEach(field => {
      if (findColumn_(headers, aliases[field]) === -1) {
        const col = sheet.getLastColumn() + 1;
        sheet.getRange(1,col).setValue(prettyHeader_(field));
        headers.push(normalize_(prettyHeader_(field)));
      }
    });

    const rowNum = sheet.getLastRow() + 1;
    const row = new Array(headers.length).fill('');

    setField_(row, headers, aliases.nome, data.nome);
    setField_(row, headers, aliases.sexo, data.sexo);
    setField_(row, headers, aliases.cpf, data.cpf);
    setField_(row, headers, aliases.nascimento, data.nascimento);
    setField_(row, headers, aliases.telefone, data.telefone);
    setField_(row, headers, aliases.cidade, data.cidade);
    setField_(row, headers, aliases.estado, data.estado);
    setField_(row, headers, aliases.rua, data.rua);
    setField_(row, headers, aliases.bairro, data.bairro);
    setField_(row, headers, aliases.status, 'Inativo');
    setField_(row, headers, aliases.origem, 'Site CLC');
    setField_(row, headers, aliases.termos, 'Li e concordo');
    setField_(row, headers, aliases.solicitacao, new Date());

    sheet.getRange(rowNum,1,1,row.length).setValues([row]);

    // Nova solicitação em vermelho, como solicitado.
    const newRow = sheet.getRange(rowNum,1,1,row.length);
    newRow.setBackground('#F4CCCC');
    newRow.setFontColor('#9C0006');

    // Formata data de nascimento e data da solicitação quando existirem.
    const birthCol = findColumn_(headers, aliases.nascimento);
    if (birthCol !== -1) sheet.getRange(rowNum,birthCol+1).setNumberFormat('dd/mm/yyyy');
    const reqCol = findColumn_(headers, aliases.solicitacao);
    if (reqCol !== -1) sheet.getRange(rowNum,reqCol+1).setNumberFormat('dd/mm/yyyy HH:mm');

    return json_({ ok:true, message:'Solicitação registrada.' });
  } catch (err) {
    return json_({ ok:false, message:String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ABA_DESTINO);
    if (!sheet) throw new Error('A aba "' + ABA_DESTINO + '" não foi encontrada.');

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return json_({ ok:true, records:[], empty:true });

    const headers = values[0].map(normalize_);

    // ESTA PLANILHA TEM EXATAMENTE ESTES CABEÇALHOS:
    // A = ID Associado | B = Nome Completo
    // Primeiro procura a correspondência EXATA para nunca confundir
    // "ID Associado" com "Nome Completo".
    function exactColumn_(name) {
      return headers.indexOf(normalize_(name));
    }
    function aliasColumn_(names) {
      for (const name of names) {
        const idx = exactColumn_(name);
        if (idx !== -1) return idx;
      }
      return -1;
    }

    const cols = {
      id: exactColumn_('ID Associado'),
      nome: exactColumn_('Nome Completo'),
      sexo: aliasColumn_(['Sexo']),
      nascimento: aliasColumn_(['Data de Nascimento','Data de Nascer','Nascimento']),
      idade: aliasColumn_(['Anos/Idade','Anos Idade','Idade']),
      cidade: aliasColumn_(['Cidade']),
      estado: aliasColumn_(['Estado','UF']),
      filiacao: aliasColumn_(['Data de Filiação','Data de Filiacao','Filiação','Filiacao']),
      categoria: aliasColumn_(['Categoria','Modalidade']),
      status: aliasColumn_(['Status','Situação','Situacao']),
      observacoes: aliasColumn_(['Observações','Observacoes','Observação','Observacao']),
      foto: aliasColumn_(['Foto URL','Link da Foto','Foto','Imagem','URL da Foto']),
      armadas: aliasColumn_(['Armadas','N de Armadas','Número de Armadas','Numero de Armadas','Pontos','Pontuação','Pontuacao'])
    };

    if (cols.nome === -1) {
      throw new Error('Não encontrei a coluna "Nome Completo" na aba Cadastro. Cabeçalhos encontrados: ' + headers.join(' | '));
    }

    function value_(row, col) {
      return col === -1 || col == null ? '' : String(row[col] || '').trim();
    }

    const records = values.slice(1)
      .map(row => ({
        id: value_(row, cols.id),
        nome: value_(row, cols.nome),
        sexo: value_(row, cols.sexo),
        nascimento: value_(row, cols.nascimento),
        idade: value_(row, cols.idade),
        cidade: value_(row, cols.cidade),
        estado: value_(row, cols.estado),
        filiacao: value_(row, cols.filiacao),
        categoria: value_(row, cols.categoria),
        status: value_(row, cols.status) || 'Ativo',
        observacoes: value_(row, cols.observacoes),
        foto: value_(row, cols.foto),
        armadas: value_(row, cols.armadas)
      }))
      .filter(item => item.nome);

    return json_({ ok:true, records:records, updatedAt:new Date().toISOString() });
  } catch (err) {
    return json_({ ok:false, message:String(err && err.message ? err.message : err) });
  }
}

function normalize_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .trim()
    .replace(/\s+/g,' ');
}

function findColumn_(headers, aliases) {
  const list = aliases.map(normalize_);
  return headers.findIndex(h => list.includes(h) || list.some(a => h.indexOf(a) !== -1));
}

function setField_(row, headers, aliases, value) {
  const col = findColumn_(headers, aliases);
  if (col !== -1) row[col] = value;
}

function prettyHeader_(field) {
  const map = {
    nome:'Nome Completo', sexo:'Sexo', cpf:'CPF', nascimento:'Data de Nascimento',
    telefone:'Telefone', cidade:'Cidade', estado:'Estado', rua:'Rua / Endereço',
    bairro:'Bairro', status:'Status', origem:'Origem', termos:'Aceite dos Termos',
    solicitacao:'Data Solicitação'
  };
  return map[field] || field;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
