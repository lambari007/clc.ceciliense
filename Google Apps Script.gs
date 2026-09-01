/** CLC - FILIAÇÕES E ÁREA RESTRITA */
const ABA_DESTINO='Cadastro';
const ABA_USUARIOS='USUÁRIOS';
const PASTA_FOTOS_ID='10Mok68v3HQfvnqib-PaZMOq4Y5H_ZGbB';

function doPost(e){
  try{
    const data=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(data.action==='login') return login_(data);
    if(data.action==='editarFiliado') return editarFiliado_(data);
    return novoFiliado_(data);
  }catch(err){console.error(err);return json_({ok:false,message:String(err.message||err)})}
}

function novoFiliado_(data){
  const required=['nome','sexo','cpf','nascimento','telefone','cidade','estado','rua','bairro'];
  for(const k of required) if(!String(data[k]||'').trim()) return json_({ok:false,message:'Campo obrigatório ausente: '+k});
  const cpf=String(data.cpf).replace(/\D/g,'');
  const tel=String(data.telefone).replace(/\D/g,'');
  if(!isValidCPF_(cpf)) return json_({ok:false,message:'CPF inválido.'});
  if(tel.length<10||tel.length>11) return json_({ok:false,message:'Telefone inválido.'});

  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName(ABA_DESTINO);
  if(!sheet) throw new Error('A aba "'+ABA_DESTINO+'" não foi encontrada.');
  let headers=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getDisplayValues()[0].map(normalize_);
  const aliases={
    id:['id associado','id','codigo','código','numero associado','número associado'],
    nome:['nome','nome completo','filiado','lacador','laçador'],sexo:['sexo','genero','gênero'],cpf:['cpf'],nascimento:['data de nascimento','nascimento','dt nascimento'],telefone:['telefone','celular','whatsapp','fone'],cidade:['cidade','municipio','município'],estado:['estado','uf'],rua:['rua','endereco','endereço'],bairro:['bairro'],categoria:['categoria','modalidade'],status:['status','situação','situacao'],observacoes:['observacoes','observações','observacao','observação'],foto:['foto url','link da foto','foto','imagem','url da foto'],origem:['origem'],termos:['termos','aceite dos termos','aceite'],solicitacao:['data solicitacao','data solicitação','solicitacao','solicitação'],pagamento:['pagamento','meio de pagamento','forma de pagamento']
  };
  Object.keys(aliases).forEach(f=>{if(findColumn_(headers,aliases[f])===-1){const title=prettyHeader_(f);sheet.getRange(1,sheet.getLastColumn()+1).setValue(title);headers.push(normalize_(title));}});
  const idCol=findColumn_(headers,aliases.id);
  if(idCol===-1) throw new Error('Não encontrei a coluna de ID. A coluna de ID deve existir na planilha Cadastro.');
  const novoId=String(data.id||generateNextId_(sheet,idCol+1)).trim();
  if(!novoId) throw new Error('Não foi possível gerar o ID do filiado.');
  let fotoUrl='';
  if(data.fotoUpload&&data.fotoUpload.data) fotoUrl=uploadFoto_(data.fotoUpload, data.nome, cpf);
  const row=new Array(headers.length).fill('');
  const values={id:novoId,nome:upperText_(data.nome),sexo:upperText_(data.sexo),cpf:data.cpf,nascimento:data.nascimento,telefone:data.telefone,cidade:upperText_(data.cidade),estado:upperText_(data.estado),rua:upperText_(data.rua),bairro:upperText_(data.bairro),categoria:upperText_(data.categoria||''),status:upperText_(data.status||'ATIVO'),observacoes:data.observacoes||'',foto:fotoUrl,origem:'ÁREA RESTRITA',termos:'CADASTRO ADMINISTRATIVO',solicitacao:new Date(),pagamento:upperText_(data.pagamento||'')};
  Object.keys(values).forEach(k=>setField_(row,headers,aliases[k],values[k]));
  const rowNum=sheet.getLastRow()+1;sheet.getRange(rowNum,1,1,row.length).setValues([row]);
  const nc=sheet.getRange(rowNum,1,1,row.length);nc.setBackground('#F4CCCC');nc.setFontColor('#9C0006');
  const birth=findColumn_(headers,aliases.nascimento);if(birth!==-1)sheet.getRange(rowNum,birth+1).setNumberFormat('dd/mm/yyyy');
  SpreadsheetApp.flush();return json_({ok:true,message:'Filiado salvo com sucesso.',id:novoId,foto:fotoUrl});
}

function doGet(){
  try{
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);if(!sheet)throw new Error('A aba "'+ABA_DESTINO+'" não foi encontrada.');
    const values=sheet.getDataRange().getDisplayValues();if(values.length<2)return json_({ok:true,records:[],empty:true});
    const headers=values[0].map(normalize_), alias=(a)=>findColumn_(headers,a), cols={id:alias(['id associado','id','codigo','código','numero associado','número associado']),nome:alias(['nome completo','nome']),sexo:alias(['sexo']),cpf:alias(['cpf']),nascimento:alias(['data de nascimento','data de nascer','nascimento']),idade:alias(['anos/idade','anos idade','idade']),telefone:alias(['telefone','celular','whatsapp','fone']),cidade:alias(['cidade']),estado:alias(['estado','uf']),filiacao:alias(['data de filiação','data de filiacao','filiação','filiacao']),categoria:alias(['categoria','modalidade']),status:alias(['status','situação','situacao']),observacoes:alias(['observações','observacoes','observação','observacao']),foto:alias(['foto url','link da foto','foto','imagem','url da foto']),armadas:alias(['armadas','n de armadas','número de armadas','numero de armadas','pontos','pontuação','pontuacao'])};
    const val=(r,c)=>c===-1?'':String(r[c]||'').trim();
    const records=values.slice(1).map(r=>({id:val(r,cols.id),nome:val(r,cols.nome),sexo:val(r,cols.sexo),cpf:val(r,cols.cpf),nascimento:val(r,cols.nascimento),idade:val(r,cols.idade),telefone:val(r,cols.telefone),cidade:val(r,cols.cidade),estado:val(r,cols.estado),filiacao:val(r,cols.filiacao),categoria:val(r,cols.categoria),status:val(r,cols.status)||'ATIVO',observacoes:val(r,cols.observacoes),foto:val(r,cols.foto),armadas:val(r,cols.armadas)})).filter(x=>x.nome);
    return json_({ok:true,records,updatedAt:new Date().toISOString()});
  }catch(err){console.error(err);return json_({ok:false,message:String(err.message||err)})}
}

function editarFiliado_(data){
  const lock=LockService.getScriptLock();let locked=false;
  try{lock.waitLock(30000);locked=true;const id=String(data.id||'').trim();const newId=String(data.newId||id).trim();if(!id)return json_({ok:false,message:'ID do filiado não informado.'});
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);if(!sheet)throw new Error('A aba Cadastro não foi encontrada.');
    const lastRow=sheet.getLastRow(),lastColumn=sheet.getLastColumn(),headers=sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0].map(normalize_);
    const aliases={id:['id associado','id','codigo','código','numero associado','número associado'],nome:['nome completo','nome'],sexo:['sexo'],cpf:['cpf'],nascimento:['data de nascimento','data nascimento','nascimento','data de nascer'],telefone:['telefone','celular','whatsapp','fone'],cidade:['cidade','municipio','município'],estado:['estado','uf'],categoria:['categoria','modalidade'],status:['status','situacao','situação'],observacoes:['observacoes','observações','observacao','observação']};
    const cols={};Object.keys(aliases).forEach(k=>cols[k]=findColumn_(headers,aliases[k]));if(cols.id===-1)throw new Error('Não encontrei a coluna de ID.');
    const ids=sheet.getRange(2,cols.id+1,lastRow-1,1).getDisplayValues().flat(),i=ids.findIndex(v=>String(v).trim()===id);if(i===-1)throw new Error('Filiado não encontrado.');const row=i+2;
    if(newId!==id){
      const duplicate=ids.some((v,idx)=>idx!==i&&String(v).trim()===newId);
      if(duplicate) return json_({ok:false,message:'Este novo ID já está em uso por outro filiado.'});
      sheet.getRange(row,cols.id+1).setValue(newId);
    }
    if(data.cpf!==undefined&&!isValidCPF_(String(data.cpf).replace(/\D/g,'')))return json_({ok:false,message:'CPF inválido.'});
    const set=(k,v)=>{if(cols[k]!==-1)sheet.getRange(row,cols[k]+1).setValue(v)};
    set('nome',upperText_(data.nome));set('sexo',upperText_(data.sexo));set('cpf',String(data.cpf||'').trim());set('telefone',String(data.telefone||'').trim());set('cidade',upperText_(data.cidade));set('estado',upperText_(data.estado));set('categoria',upperText_(data.categoria));set('status',upperText_(data.status));set('observacoes',upperText_(data.observacoes));
    if(cols.nascimento!==-1&&data.nascimento!==undefined){const p=String(data.nascimento||'').split('-'),c=sheet.getRange(row,cols.nascimento+1);if(p.length===3){c.setValue(new Date(+p[0],+p[1]-1,+p[2]));c.setNumberFormat('dd/mm/yyyy')}else if(!data.nascimento)c.clearContent();}
    SpreadsheetApp.flush();return json_({ok:true,message:'Filiado atualizado com sucesso.',id:newId,updatedAt:new Date().toISOString()});
  }catch(err){console.error(err);return json_({ok:false,message:String(err.message||err)})}finally{if(locked)lock.releaseLock()}
}

function uploadFoto_(upload,nome,cpf){
  const folder=DriveApp.getFolderById(PASTA_FOTOS_ID);const bytes=Utilities.base64Decode(String(upload.data));const mime=String(upload.type||'image/jpeg');const ext=mime.includes('png')?'png':mime.includes('webp')?'webp':'jpg';const safe=upperText_(nome).replace(/[^A-Z0-9]+/g,'_').replace(/^_|_$/g,'');const blob=Utilities.newBlob(bytes,mime,(safe||cpf||'FOTO')+'_'+Date.now()+'.'+ext);const file=folder.createFile(blob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return 'https://drive.google.com/uc?export=view&id='+file.getId();
}

function login_(data){const u=String(data.usuario||'').trim(),p=String(data.senha||'');if(!u||!p)return json_({ok:false,message:'Informe usuário e senha.'});const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_USUARIOS);if(!sh)return json_({ok:false,message:'A aba USUÁRIOS não foi encontrada.'});const v=sh.getDataRange().getDisplayValues(),h=v[0].map(normalize_),cu=h.indexOf(normalize_('Usuário')),cp=h.indexOf(normalize_('Senha')),cn=h.indexOf(normalize_('Nome')),cs=h.indexOf(normalize_('Status'));const found=v.slice(1).find(r=>String(r[cu]||'').trim()===u&&String(r[cp]||'')===p&&(cs===-1||String(r[cs]||'').trim().toUpperCase()==='ATIVO'));return found?json_({ok:true,user:{usuario:found[cu],nome:cn===-1?found[cu]:found[cn]}}):json_({ok:false,message:'Usuário ou senha inválidos.'});}
function isValidCPF_(v){v=String(v).replace(/\D/g,'');if(!/^[0-9]{11}$/.test(v)||/^(\d)\1+$/.test(v))return false;let s=0;for(let i=0;i<9;i++)s+=+v[i]*(10-i);let d=11-s%11;if(d>=10)d=0;if(d!==+v[9])return false;s=0;for(let i=0;i<10;i++)s+=+v[i]*(11-i);d=11-s%11;if(d>=10)d=0;return d===+v[10]}
function normalize_(v){return String(v||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().replace(/\s+/g,' ')}
function findColumn_(headers,aliases){const list=aliases.map(normalize_).filter(Boolean);const exact=headers.findIndex(h=>list.includes(h));if(exact!==-1)return exact;const safe=list.filter(a=>a.length>=3);return headers.findIndex(h=>safe.some(a=>h.indexOf(a)!==-1))}
function setField_(row,headers,aliases,value){const c=findColumn_(headers,aliases);if(c!==-1)row[c]=value}
function upperText_(v){return String(v==null?'':v).trim().toUpperCase()}
function prettyHeader_(f){return ({nome:'Nome Completo',sexo:'Sexo',cpf:'CPF',nascimento:'Data de Nascimento',telefone:'Telefone',cidade:'Cidade',estado:'Estado',rua:'Rua / Endereço',bairro:'Bairro',categoria:'Categoria',status:'Status',observacoes:'Observações',foto:'Foto URL',origem:'Origem',termos:'Aceite dos Termos',solicitacao:'Data Solicitação',pagamento:'Meio de Pagamento'})[f]||f}
function generateNextId_(sheet,idColumn){
  const last=sheet.getLastRow();
  if(last<2) return '1';
  const values=sheet.getRange(2,idColumn,last-1,1).getDisplayValues().flat();
  let max=0;
  values.forEach(v=>{
    const s=String(v||'').trim();
    const m=s.match(/\d+/g);
    if(m){
      const n=parseInt(m.join(''),10);
      if(Number.isFinite(n)&&n>max) max=n;
    }
  });
  return String(max+1);
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
function autorizarDrive() {
  const pasta = DriveApp.getFolderById(PASTA_FOTOS_ID);
  Logger.log(pasta.getName());
}
function testarUploadDrive() {
  const pasta = DriveApp.getFolderById(PASTA_FOTOS_ID);
  const arquivo = pasta.createFile(
    'teste_drive.txt',
    'Teste de gravação no Google Drive.'
  );
  Logger.log(arquivo.getUrl());
}
