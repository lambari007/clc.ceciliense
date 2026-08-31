/**
 * CLC - RECEBIMENTO DE PRÉ-FILIAÇÕES E ÁREA RESTRITA
 * Projeto criado dentro da mesma planilha do Google Planilhas.
 *
 * Aba principal: Cadastro
 * Aba de usuários: USUÁRIOS
 */

const ABA_DESTINO = 'Cadastro';
const ABA_USUARIOS = 'USUÁRIOS';

function doPost(e) {
  try {
    const conteudo = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(conteudo || '{}');

    if (data.action === 'login') return login_(data);
    if (data.action === 'editarFiliado') return editarFiliado_(data);

    const required = ['nome','sexo','cpf','nascimento','telefone','cidade','estado','rua','bairro'];
    for (const key of required) {
      if (!String(data[key] || '').trim()) {
        return json_({ ok:false, message:'Campo obrigatório ausente: ' + key });
      }
    }

    if (String(data.cpf).replace(/\D/g, '').length !== 11) {
      return json_({ ok:false, message:'CPF inválido.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ABA_DESTINO);
    if (!sheet) throw new Error('A aba "' + ABA_DESTINO + '" não foi encontrada.');

    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1,1,1,lastColumn).getValues()[0].map(v => normalize_(v));

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

    Object.keys(aliases).forEach(field => {
      if (findColumn_(headers, aliases[field]) === -1) {
        const col = sheet.getLastColumn() + 1;
        const titulo = prettyHeader_(field);
        sheet.getRange(1,col).setValue(titulo);
        headers.push(normalize_(titulo));
      }
    });

    const rowNum = sheet.getLastRow() + 1;
    const row = new Array(headers.length).fill('');

    setField_(row, headers, aliases.nome, upperText_(data.nome));
    setField_(row, headers, aliases.sexo, upperText_(data.sexo));
    setField_(row, headers, aliases.cpf, data.cpf);
    setField_(row, headers, aliases.nascimento, data.nascimento);
    setField_(row, headers, aliases.telefone, data.telefone);
    setField_(row, headers, aliases.cidade, upperText_(data.cidade));
    setField_(row, headers, aliases.estado, upperText_(data.estado));
    setField_(row, headers, aliases.rua, upperText_(data.rua));
    setField_(row, headers, aliases.bairro, upperText_(data.bairro));
    setField_(row, headers, aliases.status, 'INATIVO');
    setField_(row, headers, aliases.origem, 'SITE CLC');
    setField_(row, headers, aliases.termos, 'LI E CONCORDO');
    setField_(row, headers, aliases.solicitacao, new Date());

    sheet.getRange(rowNum,1,1,row.length).setValues([row]);

    const newRow = sheet.getRange(rowNum,1,1,row.length);
    newRow.setBackground('#F4CCCC');
    newRow.setFontColor('#9C0006');

    const birthCol = findColumn_(headers, aliases.nascimento);
    if (birthCol !== -1) sheet.getRange(rowNum,birthCol+1).setNumberFormat('dd/mm/yyyy');

    const reqCol = findColumn_(headers, aliases.solicitacao);
    if (reqCol !== -1) sheet.getRange(rowNum,reqCol+1).setNumberFormat('dd/mm/yyyy HH:mm');

    SpreadsheetApp.flush();
    return json_({ ok:true, message:'Solicitação registrada.' });

  } catch (err) {
    console.error(err);
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
      throw new Error('Não encontrei a coluna "Nome Completo" na aba Cadastro.');
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
        status: value_(row, cols.status) || 'ATIVO',
        observacoes: value_(row, cols.observacoes),
        foto: value_(row, cols.foto),
        armadas: value_(row, cols.armadas)
      }))
      .filter(item => item.nome);

    return json_({ ok:true, records:records, updatedAt:new Date().toISOString() });

  } catch (err) {
    console.error(err);
    return json_({ ok:false, message:String(err && err.message ? err.message : err) });
  }
}

function login_(data) {
  const usuario = String(data.usuario || '').trim();
  const senha = String(data.senha || '');

  if (!usuario || !senha) {
    return json_({ ok:false, message:'Informe usuário e senha.' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_USUARIOS);

  if (!sheet) {
    return json_({ ok:false, message:'A aba USUÁRIOS não foi encontrada.' });
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return json_({ ok:false, message:'Nenhum usuário cadastrado.' });
  }

  const headers = values[0].map(normalize_);
  const colUsuario = headers.indexOf(normalize_('Usuário'));
  const colSenha = headers.indexOf(normalize_('Senha'));
  const colNome = headers.indexOf(normalize_('Nome'));
  const colNivel = headers.indexOf(normalize_('Nível'));
  const colStatus = headers.indexOf(normalize_('Status'));

  if (colUsuario === -1 || colSenha === -1) {
    return json_({ ok:false, message:'Confira as colunas Usuário e Senha.' });
  }

  const encontrado = values.slice(1).find(row => {
    const user = String(row[colUsuario] || '').trim();
    const pass = String(row[colSenha] || '');
    const status = colStatus === -1 ? 'ATIVO' : String(row[colStatus] || '').trim().toUpperCase();

    return user === usuario && pass === senha && status === 'ATIVO';
  });

  if (!encontrado) {
    return json_({ ok:false, message:'Usuário ou senha inválidos.' });
  }

  return json_({
    ok:true,
    message:'Login realizado com sucesso.',
    user:{
      usuario: encontrado[colUsuario],
      nome: colNome === -1 ? encontrado[colUsuario] : encontrado[colNome],
      nivel: colNivel === -1 ? 'ADMIN' : encontrado[colNivel]
    }
  });
}

function editarFiliado_(data) {
  try {
    const id = String(data.id || '').trim();

    if (!id) {
      return json_({ ok:false, message:'ID do filiado não informado.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ABA_DESTINO);

    if (!sheet) {
      throw new Error('A aba "' + ABA_DESTINO + '" não foi encontrada.');
    }

    // Usa os valores exibidos para comparar o ID exatamente como o site o recebe.
    const values = sheet.getDataRange().getDisplayValues();

    if (values.length < 2) {
      throw new Error('Nenhum filiado encontrado.');
    }

    const headers = values[0].map(normalize_);

    const colunas = {
      id: headers.indexOf(normalize_('ID Associado')),
      nome: headers.indexOf(normalize_('Nome Completo')),
      sexo: headers.indexOf(normalize_('Sexo')),
      nascimento: headers.indexOf(normalize_('Data de Nascimento')),
      cidade: headers.indexOf(normalize_('Cidade')),
      estado: headers.indexOf(normalize_('Estado')),
      categoria: headers.indexOf(normalize_('Categoria')),
      status: headers.indexOf(normalize_('Status')),
      observacoes: headers.indexOf(normalize_('Observações'))
    };

    if (colunas.id === -1) {
      throw new Error('Não encontrei a coluna ID Associado.');
    }

    const linha = values.findIndex((row, index) => {
      if (index === 0) return false;
      const idPlanilha = String(row[colunas.id] || '').trim();
      return idPlanilha === id;
    });

    if (linha === -1) {
      throw new Error('Filiado não encontrado. ID recebido: ' + id);
    }

    const rowNumber = linha + 1;

    if (colunas.nome !== -1) {
      sheet.getRange(rowNumber, colunas.nome + 1).setValue(upperText_(data.nome));
    }

    if (colunas.sexo !== -1) {
      sheet.getRange(rowNumber, colunas.sexo + 1).setValue(upperText_(data.sexo));
    }

    if (colunas.nascimento !== -1 && data.nascimento !== undefined) {
      sheet.getRange(rowNumber, colunas.nascimento + 1).setValue(data.nascimento);
    }

    if (colunas.cidade !== -1) {
      sheet.getRange(rowNumber, colunas.cidade + 1).setValue(upperText_(data.cidade));
    }

    if (colunas.estado !== -1) {
      sheet.getRange(rowNumber, colunas.estado + 1).setValue(upperText_(data.estado));
    }

    if (colunas.categoria !== -1) {
      sheet.getRange(rowNumber, colunas.categoria + 1).setValue(upperText_(data.categoria));
    }

    if (colunas.status !== -1) {
      sheet.getRange(rowNumber, colunas.status + 1).setValue(upperText_(data.status));
    }

    if (colunas.observacoes !== -1) {
      sheet.getRange(rowNumber, colunas.observacoes + 1).setValue(upperText_(data.observacoes));
    }

    SpreadsheetApp.flush();

    return json_({
      ok:true,
      message:'Filiado atualizado com sucesso.',
      id:id
    });

  } catch (err) {
    console.error(err);
    return json_({
      ok:false,
      message:String(err && err.message ? err.message : err)
    });
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

function upperText_(value) {
  return String(value == null ? '' : value).trim().toUpperCase();
}

function prettyHeader_(field) {
  const map = {
    nome:'Nome Completo',
    sexo:'Sexo',
    cpf:'CPF',
    nascimento:'Data de Nascimento',
    telefone:'Telefone',
    cidade:'Cidade',
    estado:'Estado',
    rua:'Rua / Endereço',
    bairro:'Bairro',
    status:'Status',
    origem:'Origem',
    termos:'Aceite dos Termos',
    solicitacao:'Data Solicitação'
  };
  return map[field] || field;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
