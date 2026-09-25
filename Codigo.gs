/*******************************************************************
 *  VEGAS · TERMOS DIGITAIS — API (Google Apps Script)   v2
 *
 *  Apps Script = APENAS backend/banco. O frontend é estático
 *  (GitHub Pages) e conversa por fetch JSON.
 *
 *  ----------------------------------------------------------------
 *  ORDEM DE INSTALAÇÃO (rode no editor, pelo botão Executar)
 *  ----------------------------------------------------------------
 *   1) instalar()          -> cria abas, pastas e o usuário admin
 *   2) MODELO_DOC_ID       -> cole o ID do Google Docs modelo abaixo
 *   3) configurarTudo()    -> grava APP_BASE e TEMPLATE_DOC_ID
 *   4) corrigirFormatos()  -> só se a planilha já tiver dados antigos
 *   5) corrigirLinks()     -> só se já existirem links quebrados
 *   6) diagnosticar()      -> confere tudo e mostra no log
 *
 *  Implantar: Implantar > Nova implantação > App da Web
 *    - Executar como: Eu mesmo
 *    - Quem pode acessar: Qualquer pessoa
 *  IMPORTANTE: ao alterar este código, republique em
 *  Implantar > Gerenciar implantações > lápis > Versão: Nova versão.
 *  Sem isso a URL /exec continua servindo a versão antiga.
 *******************************************************************/

// ==================================================================
//  1. CONFIGURAÇÃO MANUAL — os dois valores que você precisa revisar
// ==================================================================

/* URL do site estático onde estão o index.html e o assinar.html.
   Repositório: https://github.com/VegasVig/cracha
   GitHub Pages (é ESTE que vai aqui, com barra no final):        */
var SITE_BASE = "https://vegasvig.github.io/cracha/";

/* ID do Google Docs com o modelo do termo.
   Na URL do documento, é o trecho entre /d/ e /edit:
   https://docs.google.com/document/d/AQUI_O_ID/edit
   Pode colar a URL inteira também — o código extrai o ID.          */
var MODELO_DOC_ID = "";

// ====== EMPRESAS DO GRUPO (seletor do RH) ======
var EMPRESAS = [
  {id:"ALARMES",   nome:"VEGAS ALARMES MONITORADOS LTDA",                               cnpj:"46.937.114/0001-63", cidade:"Volta Redonda"},
  {id:"MONIT_VR",  nome:"VEGAS MONITORAMENTO ELETRONICO E SERVIÇOS TERCEIRIZADOS LTDA", cnpj:"50.196.573/0001-00", cidade:"Volta Redonda"},
  {id:"MONIT_NIT", nome:"VEGAS MONITORAMENTO ELETRONICO E SERVIÇOS (NITERÓI)",          cnpj:"50.196.573/0002-91", cidade:"Niterói"},
  {id:"SEGTEC",    nome:"VEGAS SEGURANÇA E TÉCNOLOGIA PREVENTIVA LTDA",                 cnpj:"33.549.671/0001-65", cidade:"Volta Redonda"},
  {id:"VGS",       nome:"SERVIÇOS CORPORATIVOS VGS LTDA",                               cnpj:"46.955.475/0001-32", cidade:"Volta Redonda"},
  {id:"VIGILANCIA",nome:"VEGAS VIGILÂNCIA E SEGURANÇA LTDA",                            cnpj:"29.890.721/0001-04", cidade:"Volta Redonda"}
];
var CIDADES = ["Volta Redonda","Niterói","Piraí"];

// ====== PARÂMETROS ======
var ITEM_PADRAO   = "1 (um) crachá funcional provisório";
var ITENS         = ["1 (um) crachá funcional provisório","1 (um) crachá funcional definitivo","2 (dois) crachás funcionais provisórios"];
var VALOR_PADRAO  = "R$ 20,00 (vinte reais)";
var VALIDADE_DIAS = 15;
var EXIGIR_PIN    = false;
var MAX_TENTATIVAS= 5;
var LOGIN_MAX     = 8;      // tentativas de login por e-mail a cada 15 min

// ====== USUÁRIO INICIAL (criado/reposto pelo instalar) ======
var ADMIN_LOGIN  = "admin";
var ADMIN_NOME   = "Administrador";
var ADMIN_SENHA  = "Vegas4747";   // troque pelo painel (aba Conta) depois do 1º acesso
var ANEXAR_EVIDENCIAS = true;

var ABAS = {
  usuarios:    ["id","nome","email","senhaHash","salt","perfil","ativo","criadoEm"],
  funcionarios:["id","nome","cpf","email","telefone","matricula","cargo","base","criadoEm","criadoPor"],
  entregas:    ["id","empresaId","empresa","cnpj","funcionarioId","nome","cpf","documentoTipo","item",
                "valorReposicao","cidade","dataEntrega","status","token","pin","tentativas","validadeAte",
                "linkAssinatura","criadoEm","criadoPor","acessadoEm","assinadoEm","assinaturaId",
                "pdfFileId","hashPdf","canceladoEm","canceladoPor","atualizadoEm"],
  assinaturas: ["id","entregaId","nome","cpf","dataHora","ip","userAgent","assinaturaFileId","hashPdf","evidencias"],
  auditoria:   ["id","dataHora","usuario","acao","entregaId","detalhe"]
};

/* Colunas que o Sheets adora converter em Date ou em número,
   estragando comparações de texto e zeros à esquerda.
   São normalizadas na leitura e podem ser reparadas na planilha. */
var COL_DATAHORA = {criadoEm:1, atualizadoEm:1, acessadoEm:1, assinadoEm:1, canceladoEm:1, dataHora:1};
var COL_DATA     = {dataEntrega:1, validadeAte:1};
var COL_TEXTO    = {cpf:1, telefone:1, matricula:1, pin:1, token:1};

// ===================== CONFIG =====================
function prop_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ""; }
function setProp_(k,v){ PropertiesService.getScriptProperties().setProperty(k, String(v)); }
function CFG(){
  return {
    SHEET_ID:        prop_("SHEET_ID"),
    TEMPLATE_DOC_ID: prop_("TEMPLATE_DOC_ID"),
    PASTA_PDF_ID:    prop_("PASTA_PDF_ID"),
    PASTA_ASSIN_ID:  prop_("PASTA_ASSIN_ID"),
    PASTA_TMP_ID:    prop_("PASTA_TMP_ID"),
    APP_BASE:        prop_("APP_BASE")   // ex.: https://vegasvig.github.io/cracha/
  };
}

/* Grava as duas propriedades de uma vez, usando as constantes do topo.
   Use SEMPRE esta função: as funções que recebem argumento não podem
   ser executadas pelo botão Executar do editor (o argumento chega
   vazio e a propriedade é gravada como "undefined").               */
function configurarTudo(){
  if(!MODELO_DOC_ID){
    throw new Error("Cole o ID do Google Docs modelo na constante MODELO_DOC_ID, no topo deste arquivo, e rode de novo.");
  }
  var a = configurarSite(SITE_BASE);
  var b = configurarModelo(MODELO_DOC_ID);
  limparCacheModelo();
  Logger.log(a); Logger.log(b);
  return diagnosticar();
}

function configurarSite(urlBase){
  urlBase = String(urlBase||"").trim();
  if(urlBase.indexOf("http")!==0)
    throw new Error("URL do site inválida (\""+urlBase+"\"). Preencha SITE_BASE no topo do arquivo e rode configurarTudo().");
  if(urlBase.indexOf("github.com")>=0)
    throw new Error("Esse é o endereço do repositório. Use o endereço do GitHub Pages, por exemplo: https://vegasvig.github.io/cracha/");
  if(!/\/$/.test(urlBase)) urlBase += "/";
  setProp_("APP_BASE", urlBase);
  return "APP_BASE salva: "+urlBase;
}

function configurarModelo(docId){
  docId = String(docId||"").trim();
  var m = docId.match(/[-\w]{25,}/);     // aceita o ID puro ou a URL inteira
  if(!m)
    throw new Error("ID do Google Docs inválido (\""+docId+"\"). Preencha MODELO_DOC_ID no topo do arquivo e rode configurarTudo().");
  var id = m[0];
  try{ DocumentApp.openById(id).getName(); }
  catch(e){ throw new Error("Não consegui abrir esse Google Docs. Confira o ID e se a conta deste script tem acesso ao documento."); }
  setProp_("TEMPLATE_DOC_ID", id);
  return "TEMPLATE_DOC_ID salvo: "+id;
}

/* Base do site validada — usada antes de montar qualquer link. */
function baseValida_(){
  var b = CFG().APP_BASE;
  if(!b || b.indexOf("http")!==0 || b.indexOf("undefined")>=0)
    throw new Error("O endereço do site de assinatura não está configurado. No editor do Apps Script, rode configurarTudo().");
  return b;
}
function linkDe_(token){ return baseValida_() + "assinar.html?t=" + token; }

/* RODE UMA VEZ.
   Cria tudo DENTRO da própria planilha onde este script está (a "cracha").
   Se o script não estiver vinculado a nenhuma planilha, cria uma nova.   */
function instalar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) ss = SpreadsheetApp.create("VEGAS · TERMOS DIGITAIS — BANCO");
  setProp_("SHEET_ID", ss.getId());
  Object.keys(ABAS).forEach(function(nome){
    var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    // texto puro em toda a área de dados: impede o Sheets de virar datas/números
    sh.getRange(1, 1, sh.getMaxRows(), ABAS[nome].length).setNumberFormat("@");
    if(sh.getLastRow()===0){
      sh.appendRow(ABAS[nome]);
      sh.getRange(1,1,1,ABAS[nome].length).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
  });
  ["Sheet1","Planilha1","Página1"].forEach(function(n){
    var s=ss.getSheetByName(n); if(s && !ABAS[n]){ try{ss.deleteSheet(s);}catch(e){} }
  });
  var raiz = pasta_("VEGAS_TERMOS", null);
  setProp_("PASTA_PDF_ID",   pasta_("PDF_ASSINADOS", raiz).getId());
  setProp_("PASTA_ASSIN_ID", pasta_("ASSINATURAS_IMG", raiz).getId());
  setProp_("PASTA_TMP_ID",   pasta_("_TMP", raiz).getId());
  definirSenha(ADMIN_LOGIN, ADMIN_SENHA, ADMIN_NOME);
  return "OK. Banco nesta planilha: "+ss.getName()+
    "\nAbas criadas: "+Object.keys(ABAS).join(", ")+
    "\nLogin: "+ADMIN_LOGIN+" / "+ADMIN_SENHA+" (troque pelo painel, aba Conta)."+
    "\nPróximo: preencha MODELO_DOC_ID no topo e rode configurarTudo().";
}
function pasta_(nome, pai){
  var it = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if(it.hasNext()) return it.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}

// ===================== ROTEADOR JSON =====================
function doGet(e){  return rotear_(e); }
function doPost(e){ return rotear_(e); }

function rotear_(e){
  var out;
  try{
    var p = (e && e.parameter) ? e.parameter : {};
    var body = {};
    if(e && e.postData && e.postData.contents){
      try{ body = JSON.parse(e.postData.contents); }catch(err){}
    }
    var d = Object.keys(p).length ? Object.assign({}, p, body) : body;
    var acao = d.acao || "ping";
    out = {ok:true, dados: despachar_(acao, d)};
  }catch(err){
    out = {ok:false, erro: String(err && err.message ? err.message : err).replace("Error: ","")};
  }
  return ContentService.createTextOutput(JSON.stringify(out))
          .setMimeType(ContentService.MimeType.JSON);
}

function despachar_(acao, d){
  switch(acao){
    // --- público (funcionário) ---
    case "ping":        return {msg:"VEGAS Termos API online"};
    case "pubInfo":     return pubInfo(d.token);
    case "pubValidar":  return pubValidar(d.token, d.cpf5, d.pin);
    case "pubAssinar":  return pubAssinar(d.acc, d.assinatura, d.meta);
    case "pubPdf":      return pubPdf(d.acc);
    // --- RH (exige sessão) ---
    case "login":       return apiLogin(d.email, d.senha);
    case "logout":      return apiLogout(d.tk);
    case "trocarSenha": return apiTrocarSenha(d.tk, d.atual, d.nova);
    case "bootstrap":   return apiBootstrap(d.tk);
    case "criarEntrega":return apiCriarEntrega(d.tk, d.dados||{});
    case "listar":      return apiListar(d.tk, d.filtro||{});
    case "dashboard":   return apiDashboard(d.tk);
    case "detalhe":     return apiDetalhe(d.tk, d.id);
    case "cancelar":    return apiCancelar(d.tk, d.id, d.motivo);
    case "excluir":     return apiExcluir(d.tk, d.id, d.motivo);
    case "novoLink":    return apiNovoLink(d.tk, d.id);
    case "pdf":         return apiPdf(d.tk, d.id);
    case "status":      return apiStatus(d.tk);
    default: throw new Error("ação desconhecida: "+acao);
  }
}

// ===================== BANCO (Sheets) =====================
function sh_(aba){
  var id = CFG().SHEET_ID;
  if(!id) throw new Error("Banco não instalado. Rode instalar() no editor do Apps Script.");
  var sh = SpreadsheetApp.openById(id).getSheetByName(aba);
  if(!sh) throw new Error("Aba \""+aba+"\" não existe na planilha. Rode instalar().");
  return sh;
}
function fmtDH_(d){ return Utilities.formatDate(d,"America/Sao_Paulo","yyyy-MM-dd HH:mm:ss"); }
function fmtD_(d){  return Utilities.formatDate(d,"America/Sao_Paulo","yyyy-MM-dd"); }

/* Devolve todo valor lido da planilha no formato que o código espera:
   datas como texto "yyyy-MM-dd[ HH:mm:ss]" e CPF com 11 dígitos.
   Isso conserta, na leitura, registros antigos já convertidos pelo Sheets. */
function normCel_(head, v){
  if(v instanceof Date) return COL_DATA[head] ? fmtD_(v) : fmtDH_(v);
  if(head==="cpf"){
    var s = String(v==null?"":v).replace(/\D/g,"");
    if(!s) return "";
    while(s.length<11) s = "0"+s;      // Sheets come o zero da frente
    return s;
  }
  if(COL_TEXTO[head]) return String(v==null?"":v);
  return v;
}

function lerTabela_(aba){
  var sh = sh_(aba), last = sh.getLastRow();
  if(last<2) return [];
  var head = ABAS[aba];
  return sh.getRange(2,1,last-1,head.length).getValues().map(function(r,i){
    var o = {_linha:i+2};
    head.forEach(function(h,c){ o[h] = normCel_(h, r[c]); });
    return o;
  }).filter(function(o){ return o.id!==""; });
}
function buscar_(aba, pred){ var r = lerTabela_(aba).filter(pred); return r.length?r[0]:null; }
function porId_(aba, id){ return buscar_(aba, function(o){ return String(o.id)===String(id); }); }
function inserir_(aba, obj){
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    obj.id = obj.id || novoId_(aba);
    sh_(aba).appendRow(ABAS[aba].map(function(h){ return obj[h]!==undefined ? obj[h] : ""; }));
    return obj;
  } finally { lock.releaseLock(); }
}
function atualizar_(aba, id, patch){
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var reg = porId_(aba, id); if(!reg) throw new Error("Registro não encontrado");
    var head = ABAS[aba], sh = sh_(aba);
    head.forEach(function(h,c){ if(patch[h]!==undefined) sh.getRange(reg._linha, c+1).setValue(patch[h]); });
    if(head.indexOf("atualizadoEm")>=0) sh.getRange(reg._linha, head.indexOf("atualizadoEm")+1).setValue(agora_());
    return porId_(aba,id);
  } finally { lock.releaseLock(); }
}
function novoId_(aba){
  var pfx = {entregas:"ENT",funcionarios:"FUN",assinaturas:"ASS",auditoria:"AUD",usuarios:"USR"}[aba]||"REG";
  return pfx + "-" + new Date().getFullYear() + "-" + rand_(6);
}
function agora_(){ return fmtDH_(new Date()); }
function hoje_(){  return fmtD_(new Date()); }
function rand_(n){
  var abc="ABCDEFGHJKLMNPQRSTUVWXYZ23456789", s="", b=Utilities.getUuid().replace(/-/g,"");
  for(var i=0;i<n;i++) s += abc.charAt(Math.floor((parseInt(b.substr(i*2,2),16)/256)*abc.length));
  return s;
}
function token_(){ return rand_(12)+rand_(12)+rand_(8); }
function sha256_(txt){ return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8)); }
function sha256Bytes_(bytes){ return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)); }
function hex_(bytes){ return bytes.map(function(b){ var v=(b<0?b+256:b).toString(16); return v.length===1?"0"+v:v; }).join(""); }
function auditar_(usuario, acao, entregaId, detalhe){
  try{ inserir_("auditoria",{id:novoId_("auditoria"),dataHora:agora_(),usuario:usuario||"-",
        acao:acao, entregaId:entregaId||"", detalhe:detalhe||""}); }catch(e){}
}

// ===================== MANUTENÇÃO / REPARO =====================

/* Formata as colunas como texto puro e reescreve os valores já
   convertidos pelo Sheets (datas viradas Date, CPF sem o zero inicial).
   Seguro rodar quantas vezes quiser.                                 */
function corrigirFormatos(){
  var ss = SpreadsheetApp.openById(CFG().SHEET_ID), log = [];
  Object.keys(ABAS).forEach(function(nome){
    var sh = ss.getSheetByName(nome);
    if(!sh){ log.push(nome+": aba não existe"); return; }
    var head = ABAS[nome], nCol = head.length;
    sh.getRange(1, 1, sh.getMaxRows(), nCol).setNumberFormat("@");
    var last = sh.getLastRow();
    if(last < 2){ log.push(nome+": sem registros"); return; }
    var rg = sh.getRange(2, 1, last-1, nCol), vals = rg.getValues(), mud = 0;
    for(var r=0; r<vals.length; r++){
      for(var c=0; c<nCol; c++){
        var novo = normCel_(head[c], vals[r][c]);
        if(novo !== vals[r][c]){ vals[r][c] = novo; mud++; }
      }
    }
    if(mud) rg.setValues(vals);
    log.push(nome+": "+(last-1)+" registro(s), "+mud+" célula(s) corrigida(s)");
  });
  return log.join("\n");
}

/* Reescreve o link de assinatura das entregas ainda abertas usando a
   APP_BASE atual. Use depois de configurar (ou trocar) o site.        */
function corrigirLinks(){
  var base = baseValida_(), n = 0, total = 0;
  lerTabela_("entregas").forEach(function(e){
    if(e.status==="ASSINADO" || e.status==="CANCELADO") return;
    if(!e.token) return;
    total++;
    var novo = base + "assinar.html?t=" + e.token;
    if(String(e.linkAssinatura) !== novo){ atualizar_("entregas", e.id, {linkAssinatura:novo}); n++; }
  });
  return n+" de "+total+" link(s) aberto(s) atualizado(s) para "+base;
}

// ===================== AUTENTICAÇÃO DO RH =====================
function hashSenha_(senha, salt){
  var h = salt + "|" + senha;
  for(var i=0;i<2000;i++) h = sha256_(h);
  return h;
}
/* Cria o usuário ou repõe a senha dele, e limpa o bloqueio de tentativas.
   Para resetar acesso sem passar argumento, edite RESET_* e rode resetarAdmin(). */
function definirSenha(login, senha, nome){
  login = String(login||"").toLowerCase().trim();
  if(!login || !senha) throw new Error("Informe login e senha. Pelo editor, use resetarAdmin().");
  var salt = rand_(16);
  var u = buscar_("usuarios", function(x){ return String(x.email).toLowerCase()===login; });
  if(u){
    atualizar_("usuarios", u.id, {salt:salt, senhaHash:hashSenha_(senha,salt), ativo:"SIM"});
  }else{
    inserir_("usuarios",{id:novoId_("usuarios"), nome:nome||login, email:login,
      senhaHash:hashSenha_(senha,salt), salt:salt, perfil:"admin", ativo:"SIM", criadoEm:agora_()});
  }
  CacheService.getScriptCache().remove("rl_"+sha256_(login).substr(0,20));
  return "Acesso definido para: "+login;
}
/* Repõe a senha do admin sem precisar passar argumento pelo editor. */
function resetarAdmin(){ return definirSenha(ADMIN_LOGIN, ADMIN_SENHA, ADMIN_NOME); }

function criarUsuario(email, nome, senha, perfil){
  var salt = rand_(16);
  return inserir_("usuarios",{id:novoId_("usuarios"),nome:nome,email:String(email).toLowerCase().trim(),
    senhaHash:hashSenha_(senha,salt), salt:salt, perfil:perfil||"rh", ativo:"SIM", criadoEm:agora_()});
}
function apiLogin(email, senha){
  email = String(email||"").toLowerCase().trim();
  var cache = CacheService.getScriptCache(), rk = "rl_"+sha256_(email).substr(0,20);
  var n = Number(cache.get(rk)||0);
  if(n >= LOGIN_MAX) throw new Error("Muitas tentativas. Aguarde 15 minutos.");

  var u = buscar_("usuarios", function(x){ return String(x.email).toLowerCase()===email && x.ativo==="SIM"; });
  if(!u || hashSenha_(senha, u.salt)!==u.senhaHash){
    cache.put(rk, String(n+1), 900);
    auditar_(email,"LOGIN_FALHOU","","tentativa "+(n+1));
    Utilities.sleep(700);
    throw new Error("Usuário ou senha inválidos.");
  }
  cache.remove(rk);
  var tk = token_();
  cache.put("sess_"+tk, JSON.stringify({id:u.id,nome:u.nome,email:u.email,perfil:u.perfil}), 21600);
  auditar_(u.email,"LOGIN","","");
  return {token:tk, nome:u.nome, perfil:u.perfil};
}
function sess_(tk){
  var j = CacheService.getScriptCache().get("sess_"+String(tk||""));
  if(!j) throw new Error("SESSAO_EXPIRADA");
  return JSON.parse(j);
}
function apiLogout(tk){ CacheService.getScriptCache().remove("sess_"+tk); return true; }
function apiTrocarSenha(tk, atual, nova){
  var s = sess_(tk), u = porId_("usuarios", s.id);
  if(hashSenha_(atual,u.salt)!==u.senhaHash) throw new Error("Senha atual incorreta.");
  if(String(nova||"").length<8) throw new Error("A nova senha precisa de pelo menos 8 caracteres.");
  var salt = rand_(16);
  atualizar_("usuarios", u.id, {salt:salt, senhaHash:hashSenha_(nova,salt)});
  auditar_(s.email,"TROCA_SENHA","","");
  return true;
}

// ===================== CPF =====================
function soDig_(v){ return String(v||"").replace(/\D/g,""); }
function cpfValido_(cpf){
  cpf = soDig_(cpf);
  if(cpf.length!==11 || /^(\d)\1{10}$/.test(cpf)) return false;
  for(var t=9;t<11;t++){
    var s=0;
    for(var i=0;i<t;i++) s += parseInt(cpf.charAt(i),10)*((t+1)-i);
    var dv = (s*10)%11; if(dv===10) dv=0;
    if(dv!==parseInt(cpf.charAt(t),10)) return false;
  }
  return true;
}
function cpfFmt_(cpf){ return soDig_(cpf).replace(/(\d{3})(\d{3})(\d{3})(\d{2})/,"$1.$2.$3-$4"); }
function cpfMask_(cpf){ cpf=soDig_(cpf); return cpf ? "***.***."+cpf.substr(6,3)+"-"+cpf.substr(9,2) : ""; }

// ===================== API DO PAINEL =====================
function apiBootstrap(tk){
  sess_(tk);
  return {empresas:EMPRESAS, cidades:CIDADES, itens:ITENS, valorPadrao:VALOR_PADRAO,
          validadeDias:VALIDADE_DIAS, exigirPin:EXIGIR_PIN};
}
/* Situação da configuração, visível pelo painel (útil para suporte). */
function apiStatus(tk){
  sess_(tk);
  var cfg = CFG(), modeloOk = false, modeloErro = "";
  try{ textoModelo_(); modeloOk = true; }catch(e){ modeloErro = e.message; }
  var baseOk = true, baseErro = "";
  try{ baseValida_(); }catch(e){ baseOk = false; baseErro = e.message; }
  return {appBase:cfg.APP_BASE, baseOk:baseOk, baseErro:baseErro,
          templateDocId:cfg.TEMPLATE_DOC_ID, modeloOk:modeloOk, modeloErro:modeloErro};
}
function apiCriarEntrega(tk, dados){
  var s = sess_(tk);
  baseValida_();                       // falha cedo, com mensagem clara
  var emp = EMPRESAS.filter(function(e){ return e.id===dados.empresaId; })[0];
  if(!emp) throw new Error("Selecione a empresa.");
  var nome = String(dados.nome||"").trim().toUpperCase().replace(/\s+/g," ");
  if(nome.split(" ").length<2) throw new Error("Informe o nome completo do funcionário.");
  if(!cpfValido_(dados.cpf)) throw new Error("CPF inválido.");
  var cpf = soDig_(dados.cpf);

  var func = buscar_("funcionarios", function(f){ return soDig_(f.cpf)===cpf; });
  if(!func){
    func = inserir_("funcionarios",{id:novoId_("funcionarios"),nome:nome,cpf:cpf,
      email:String(dados.email||"").trim(), telefone:soDig_(dados.telefone),
      matricula:String(dados.matricula||"").trim(), cargo:String(dados.cargo||"").trim(),
      base:dados.cidade||emp.cidade, criadoEm:agora_(), criadoPor:s.email});
  }

  var id = novoId_("entregas"), tkAssin = token_();
  var pin = EXIGIR_PIN ? String(Math.floor(100000+Math.random()*900000)) : "";
  var validade = new Date(); validade.setDate(validade.getDate()+VALIDADE_DIAS);
  var link = linkDe_(tkAssin);

  inserir_("entregas",{
    id:id, empresaId:emp.id, empresa:emp.nome, cnpj:emp.cnpj,
    funcionarioId:func.id, nome:nome, cpf:cpf,
    documentoTipo:"TERMO_CRACHA", item:dados.item||ITEM_PADRAO,
    valorReposicao:dados.valorReposicao||VALOR_PADRAO, cidade:dados.cidade||emp.cidade,
    dataEntrega:String(dados.dataEntrega||hoje_()).substr(0,10),
    status:"PENDENTE", token:tkAssin, pin:pin, tentativas:0,
    validadeAte:fmtD_(validade),
    linkAssinatura:link, criadoEm:agora_(), criadoPor:s.email, atualizadoEm:agora_()
  });
  auditar_(s.email,"ENTREGA_CRIADA",id, emp.nome+" · "+nome);
  return {id:id, link:link, pin:pin};
}
function apiListar(tk, filtro){
  sess_(tk);
  var q = String(filtro.q||"").toLowerCase().trim(), qd = soDig_(filtro.q);
  var baix = pdfsBaixados_();
  return lerTabela_("entregas").filter(naoExcluido_).map(expirar_).filter(function(e){
    if(filtro.status && filtro.status!=="TODOS" && e.status!==filtro.status) return false;
    if(filtro.empresaId && filtro.empresaId!=="TODAS" && e.empresaId!==filtro.empresaId) return false;
    if(filtro.de && String(e.criadoEm).substr(0,10) < filtro.de) return false;
    if(filtro.ate && String(e.criadoEm).substr(0,10) > filtro.ate) return false;
    if(q){
      var ok = String(e.nome).toLowerCase().indexOf(q)>=0
            || String(e.id).toLowerCase().indexOf(q)>=0
            || (qd.length>=3 && soDig_(e.cpf).indexOf(qd)>=0);
      if(!ok) return false;
    }
    return true;
  }).sort(function(a,b){ return String(b.criadoEm).localeCompare(String(a.criadoEm)); })
    .slice(0, filtro.limite||300).map(function(e){
      return {id:e.id, nome:e.nome, cpfMask:cpfMask_(e.cpf), empresa:e.empresa, empresaId:e.empresaId,
        cidade:e.cidade, dataEntrega:e.dataEntrega, status:e.status, criadoEm:e.criadoEm,
        assinadoEm:e.assinadoEm, validadeAte:e.validadeAte, link:e.linkAssinatura,
        temPdf: !!e.pdfFileId, criadoPor:e.criadoPor,
        pdfBaixadoEm: baix[e.id] ? baix[e.id].em : "", pdfBaixadoPor: baix[e.id] ? baix[e.id].por : ""};
    });
}
function apiDashboard(tk){
  sess_(tk);
  var arr = lerTabela_("entregas").filter(naoExcluido_).map(expirar_);
  var cont = function(st){ return arr.filter(function(e){ return e.status===st; }).length; };
  var mesAtual = Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM");
  return {
    total:arr.length, pendentes:cont("PENDENTE")+cont("ACESSADO"), assinados:cont("ASSINADO"),
    cancelados:cont("CANCELADO"), expirados:cont("EXPIRADO"),
    mes: arr.filter(function(e){ return String(e.criadoEm).substr(0,7)===mesAtual; }).length,
    porEmpresa: EMPRESAS.map(function(em){
      return {empresa:em.nome, total:arr.filter(function(e){ return e.empresaId===em.id; }).length};
    }).filter(function(x){ return x.total>0; }),
    recentes: arr.sort(function(a,b){ return String(b.criadoEm).localeCompare(String(a.criadoEm)); })
                 .slice(0,8).map(function(e){ return {id:e.id,nome:e.nome,status:e.status,criadoEm:e.criadoEm}; })
  };
}
function apiDetalhe(tk, id){
  var s = sess_(tk);
  var e = expirar_(porId_("entregas", id));
  if(!e || e.status==="EXCLUIDO") throw new Error("Entrega não encontrada.");
  var baix = pdfsBaixados_()[id];
  var a = buscar_("assinaturas", function(x){ return x.entregaId===id; });
  var evs = lerTabela_("auditoria").filter(function(x){ return x.entregaId===id; })
             .sort(function(x,y){ return String(x.dataHora).localeCompare(String(y.dataHora)); });
  auditar_(s.email,"ENTREGA_CONSULTADA",id,"");

  /* O texto do termo vem do Google Docs modelo. Se o modelo estiver
     indisponível, o registro ainda abre — antes a tela inteira quebrava. */
  var termo, termoErro = "";
  try{ termo = montarTermo_(e); }
  catch(err){ termoErro = err.message; termo = ["[Texto do termo indisponível: "+err.message+"]"]; }

  return {
    entrega:{id:e.id, nome:e.nome, cpfMask:cpfMask_(e.cpf), empresa:e.empresa, cnpj:e.cnpj,
      cidade:e.cidade, item:e.item, valorReposicao:e.valorReposicao, dataEntrega:e.dataEntrega,
      status:e.status, link:e.linkAssinatura, pin:e.pin, validadeAte:e.validadeAte,
      criadoEm:e.criadoEm, criadoPor:e.criadoPor, acessadoEm:e.acessadoEm, assinadoEm:e.assinadoEm,
      hashPdf:e.hashPdf, temPdf:!!e.pdfFileId,
      pdfBaixadoEm: baix ? baix.em : "", pdfBaixadoPor: baix ? baix.por : ""},
    assinatura: a ? {id:a.id, dataHora:a.dataHora, ip:a.ip, userAgent:a.userAgent, hashPdf:a.hashPdf} : null,
    eventos: evs.map(function(x){ return {dataHora:x.dataHora, usuario:x.usuario, acao:x.acao, detalhe:x.detalhe}; }),
    termo: termo, termoErro: termoErro
  };
}
function apiCancelar(tk, id, motivo){
  var s = sess_(tk);
  var e = porId_("entregas", id); if(!e) throw new Error("Entrega não encontrada.");
  if(e.status==="ASSINADO") throw new Error("Termo já assinado — não pode ser cancelado.");
  if(e.status==="EXCLUIDO") throw new Error("Este envio foi excluído.");
  atualizar_("entregas", id, {status:"CANCELADO", canceladoEm:agora_(), canceladoPor:s.email, token:""});
  auditar_(s.email,"ENTREGA_CANCELADA",id, motivo||"");
  return true;
}
/* Exclui um envio feito por engano.
   - Só vale para termos NÃO assinados (termo assinado é prova e fica guardado).
   - O link para de funcionar na hora e o registro some das listas e do dashboard.
   - A linha continua na planilha com status EXCLUIDO + auditoria (quem e quando),
     para não perder rastreabilidade.                                         */
function apiExcluir(tk, id, motivo){
  var s = sess_(tk);
  var e = porId_("entregas", id); if(!e) throw new Error("Entrega não encontrada.");
  if(e.status==="ASSINADO") throw new Error("Termo já assinado — não pode ser excluído.");
  if(e.status==="EXCLUIDO") return true;
  atualizar_("entregas", id, {status:"EXCLUIDO", canceladoEm:agora_(), canceladoPor:s.email, token:""});
  auditar_(s.email,"ENTREGA_EXCLUIDA",id, (motivo||"enviado por engano")+" · "+e.nome);
  return true;
}
function naoExcluido_(e){ return e.status!=="EXCLUIDO"; }

/* Último download do PDF feito pelo painel, por entrega: {id:{em,por}} */
function pdfsBaixados_(){
  var m = {};
  lerTabela_("auditoria").forEach(function(a){
    if(a.acao!=="PDF_BAIXADO") return;
    var atual = m[a.entregaId];
    if(!atual || String(a.dataHora) > String(atual.em)) m[a.entregaId] = {em:String(a.dataHora), por:a.usuario};
  });
  return m;
}

function apiNovoLink(tk, id){
  var s = sess_(tk);
  baseValida_();
  var e = expirar_(porId_("entregas", id)); if(!e) throw new Error("Entrega não encontrada.");
  if(e.status==="ASSINADO") throw new Error("Termo já assinado.");
  if(e.status==="EXCLUIDO") throw new Error("Este envio foi excluído.");
  var t = token_(), validade = new Date(); validade.setDate(validade.getDate()+VALIDADE_DIAS);
  var link = linkDe_(t);
  atualizar_("entregas", id, {token:t, status:"PENDENTE", tentativas:0,
    validadeAte:fmtD_(validade), linkAssinatura:link});
  auditar_(s.email,"LINK_REGERADO",id,"");
  return {link:link};
}
function apiPdf(tk, id){
  var s = sess_(tk);
  var e = porId_("entregas", id); if(!e || !e.pdfFileId) throw new Error("PDF não disponível.");
  auditar_(s.email,"PDF_BAIXADO",id,"");
  var blob = DriveApp.getFileById(e.pdfFileId).getBlob();
  return {nome:blob.getName(), base64:Utilities.base64Encode(blob.getBytes())};
}

// ===================== FLUXO PÚBLICO =====================
function porToken_(t){
  t = String(t||"").trim(); if(t.length<20) return null;
  return buscar_("entregas", function(e){ return String(e.token)===t; });
}
function expirar_(e){
  if(!e) return e;
  if((e.status==="PENDENTE"||e.status==="ACESSADO") && e.validadeAte &&
     hoje_() > String(e.validadeAte).substr(0,10)){
    atualizar_("entregas", e.id, {status:"EXPIRADO"});
    e.status = "EXPIRADO";
  }
  return e;
}
function pubInfo(token){
  var e = expirar_(porToken_(token));
  if(!e) return {valido:false, motivo:"LINK_INVALIDO"};
  if(e.status==="CANCELADO") return {valido:false, motivo:"CANCELADO"};
  if(e.status==="EXPIRADO")  return {valido:false, motivo:"EXPIRADO"};
  if(e.status==="ASSINADO")  return {valido:true, status:"ASSINADO", assinadoEm:e.assinadoEm};
  return {valido:true, status:e.status, empresa:e.empresa, documento:"Termo de Entrega de Crachá",
          primeiroNome:String(e.nome).split(" ")[0], exigirPin:!!e.pin};
}
function pubValidar(token, cpf5, pin){
  var e = expirar_(porToken_(token));
  if(!e) throw new Error("Link inválido.");
  if(e.status==="CANCELADO") throw new Error("Este link foi cancelado pelo RH.");
  if(e.status==="EXPIRADO")  throw new Error("Este link expirou. Solicite um novo ao RH.");
  if(Number(e.tentativas||0) >= MAX_TENTATIVAS) throw new Error("Bloqueado por tentativas incorretas. Procure o RH.");

  var ok = soDig_(cpf5)===soDig_(e.cpf).substr(6,5) && (!e.pin || String(pin||"").trim()===String(e.pin));
  if(!ok){
    atualizar_("entregas", e.id, {tentativas:Number(e.tentativas||0)+1});
    auditar_("funcionario","VALIDACAO_FALHOU",e.id,"tentativa "+(Number(e.tentativas||0)+1));
    Utilities.sleep(800);
    throw new Error("Dados não conferem. Confira os 5 últimos dígitos do seu CPF.");
  }
  var termo;
  try{ termo = montarTermo_(e); }
  catch(err){
    auditar_("funcionario","MODELO_INDISPONIVEL",e.id,err.message);
    throw new Error("O texto do termo não pôde ser carregado. Avise o RH e tente mais tarde.");
  }
  var acc = token_();
  CacheService.getScriptCache().put("acc_"+acc, e.id, 1800);
  if(e.status==="PENDENTE"){
    atualizar_("entregas", e.id, {status:"ACESSADO", acessadoEm:agora_(), tentativas:0});
    auditar_("funcionario","LINK_ACESSADO",e.id,"");
  }
  return {acc:acc, status:e.status==="ASSINADO"?"ASSINADO":"ACESSADO",
          termo:termo, assinadoEm:e.assinadoEm||""};
}
function acc_(acc){
  var id = CacheService.getScriptCache().get("acc_"+String(acc||""));
  if(!id) throw new Error("Sessão expirada. Abra o link novamente.");
  return porId_("entregas", id);
}
function pubAssinar(acc, assinaturaDataUrl, meta){
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var e = acc_(acc);
    if(e.status==="ASSINADO") throw new Error("Este termo já foi assinado.");
    if(e.status==="CANCELADO"||e.status==="EXPIRADO") throw new Error("Link indisponível.");
    if(!assinaturaDataUrl || String(assinaturaDataUrl).indexOf("data:image/")!==0) throw new Error("Assinatura não capturada.");
    meta = meta || {};
    if(meta.aceite!==true) throw new Error("É necessário aceitar o termo.");

    var bytes = Utilities.base64Decode(String(assinaturaDataUrl).split(",")[1]);
    var assinBlob = Utilities.newBlob(bytes, "image/png", e.id+"_assinatura.png");
    var arqAssin = DriveApp.getFolderById(CFG().PASTA_ASSIN_ID).createFile(assinBlob);

    var quando = agora_(), assinId = novoId_("assinaturas");
    var res = gerarPdf_(e, assinBlob, {assinaturaId:assinId, dataHora:quando,
      ip:String(meta.ip||"não informado").substr(0,45), userAgent:String(meta.userAgent||"").substr(0,180)});

    inserir_("assinaturas",{id:assinId, entregaId:e.id, nome:e.nome, cpf:soDig_(e.cpf), dataHora:quando,
      ip:String(meta.ip||"não informado"), userAgent:String(meta.userAgent||"").substr(0,300),
      assinaturaFileId:arqAssin.getId(), hashPdf:res.hash,
      evidencias:JSON.stringify({origemIp:"informado pelo navegador (não verificável no Apps Script)",
        tela:meta.tela||"", fuso:meta.fuso||"", origem:meta.origem||"", aceiteEm:quando,
        tokenUsado:String(e.token).substr(0,6)+"…"})});

    atualizar_("entregas", e.id, {status:"ASSINADO", assinadoEm:quando, assinaturaId:assinId,
      pdfFileId:res.fileId, hashPdf:res.hash, token:""});
    auditar_("funcionario","TERMO_ASSINADO",e.id,"assinatura "+assinId);

    CacheService.getScriptCache().put("accpdf_"+acc, e.id, 3600);
    return {assinadoEm:quando, assinaturaId:assinId, hash:res.hash};
  } finally { lock.releaseLock(); }
}
function pubPdf(acc){
  var c = CacheService.getScriptCache();
  var id = c.get("accpdf_"+String(acc||"")) || c.get("acc_"+String(acc||""));
  if(!id) throw new Error("Sessão expirada.");
  var e = porId_("entregas", id);
  if(!e || !e.pdfFileId) throw new Error("Documento ainda não disponível.");
  auditar_("funcionario","PDF_BAIXADO_COLABORADOR",e.id,"");
  var blob = DriveApp.getFileById(e.pdfFileId).getBlob();
  return {nome:blob.getName(), base64:Utilities.base64Encode(blob.getBytes())};
}

// ===================== TERMO =====================
function camposDe_(e, dataRef){
  var d = dataRef ? new Date(String(dataRef).replace(" ","T")) : new Date();
  if(isNaN(d.getTime())) d = new Date();
  var meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  return {
    NOME:e.nome, CPF:cpfFmt_(e.cpf), EMPRESA:e.empresa, CNPJ:e.cnpj,
    ITEM:e.item||ITEM_PADRAO, VALOR_REPOSICAO:e.valorReposicao||VALOR_PADRAO,
    CIDADE:e.cidade, DIA:Utilities.formatDate(d,"America/Sao_Paulo","dd"),
    MES:meses[Number(Utilities.formatDate(d,"America/Sao_Paulo","M"))-1],
    ANO:Utilities.formatDate(d,"America/Sao_Paulo","yyyy")
  };
}
function montarTermo_(e){
  var c = camposDe_(e, e.status==="ASSINADO" && e.assinadoEm ? e.assinadoEm : null);
  return textoModelo_().map(function(l){
    var t = l;
    Object.keys(c).forEach(function(k){ t = t.split("{{"+k+"}}").join(c[k]); });
    return t.split("{{ASSINATURA_COLABORADOR}}").join("").split("{{ASSINATURA_RESPONSAVEL}}").join("");
  }).filter(function(t,i,arr){ return !(t.trim()==="" && (i===0 || arr[i-1].trim()==="")); });
}
function textoModelo_(){
  var cache = CacheService.getScriptCache(), k="modelo_txt", c=cache.get(k);
  if(c) return JSON.parse(c);
  var id = CFG().TEMPLATE_DOC_ID;
  if(!id || id==="undefined")
    throw new Error("O modelo do termo não está configurado. No editor do Apps Script, preencha MODELO_DOC_ID e rode configurarTudo().");
  var doc;
  try{ doc = DocumentApp.openById(id); }
  catch(err){ throw new Error("Não foi possível abrir o Google Docs modelo (ID "+id+"). Confira o ID e as permissões, e rode configurarTudo()."); }
  var linhas = doc.getBody().getText().split("\n");
  cache.put(k, JSON.stringify(linhas), 21600);
  return linhas;
}
function limparCacheModelo(){ CacheService.getScriptCache().remove("modelo_txt"); return "ok"; }

/* Confere se o Docs modelo tem todos os marcadores esperados. */
function conferirModelo(){
  var txt = textoModelo_().join("\n");
  var esperados = ["NOME","CPF","EMPRESA","CNPJ","ITEM","VALOR_REPOSICAO","CIDADE","DIA","MES","ANO","ASSINATURA_COLABORADOR"];
  var falta = esperados.filter(function(k){ return txt.indexOf("{{"+k+"}}")<0; });
  var msg = falta.length ? "Faltam no modelo: {{"+falta.join("}}, {{")+"}}"
                         : "Modelo OK — todos os marcadores presentes.";
  Logger.log(msg);
  return msg;
}

// ===================== PDF =====================
function gerarPdf_(e, assinBlob, ev){
  var cfg = CFG();
  if(!cfg.TEMPLATE_DOC_ID || cfg.TEMPLATE_DOC_ID==="undefined")
    throw new Error("O modelo do termo não está configurado. Avise o RH.");
  var copia = DriveApp.getFileById(cfg.TEMPLATE_DOC_ID).makeCopy("TMP_"+e.id, DriveApp.getFolderById(cfg.PASTA_TMP_ID));
  try{
    var doc = DocumentApp.openById(copia.getId()), body = doc.getBody();
    var c = camposDe_(e, ev.dataHora);
    Object.keys(c).forEach(function(k){ body.replaceText("\\{\\{"+k+"\\}\\}", String(c[k]==null?"":c[k])); });
    body.replaceText("\\{\\{ASSINATURA_RESPONSAVEL\\}\\}", "");

    var achou = body.findText("\\{\\{ASSINATURA_COLABORADOR\\}\\}");
    if(achou){
      var el = achou.getElement().asText();
      el.deleteText(achou.getStartOffset(), achou.getEndOffsetInclusive());
      var par = el.getParent().asParagraph();
      var img = par.insertInlineImage(0, assinBlob);
      var prop = img.getHeight()/img.getWidth();
      img.setWidth(170); img.setHeight(Math.round(170*prop));
      par.setLineSpacing(1);
    }
    if(ANEXAR_EVIDENCIAS){
      body.appendParagraph("").setFontSize(8);
      body.appendParagraph("EVIDÊNCIAS DA ASSINATURA ELETRÔNICA").setFontSize(8).setBold(true);
      ["Documento: Termo de Entrega de Crachá · Registro "+e.id,
       "Empresa: "+e.empresa+" — CNPJ "+e.cnpj,
       "Assinante: "+e.nome+" — CPF "+cpfFmt_(e.cpf),
       "Data/hora da assinatura: "+ev.dataHora+" (America/Sao_Paulo)",
       "Identificador da assinatura: "+ev.assinaturaId,
       "IP informado pelo navegador: "+ev.ip,
       "Dispositivo: "+ev.userAgent,
       "Assinatura eletrônica simples, coletada por captura gráfica com confirmação de identidade por dígitos do CPF."
      ].forEach(function(t){ body.appendParagraph(t).setFontSize(8).setBold(false); });
    }
    doc.saveAndClose();

    var nomeArq = "TERMO_CRACHA_"+String(e.nome).replace(/[^A-Za-z0-9]+/g,"_")+"_"+e.id+".pdf";
    var pdfBlob = DriveApp.getFileById(copia.getId()).getAs(MimeType.PDF).setName(nomeArq);
    var arq = DriveApp.getFolderById(cfg.PASTA_PDF_ID).createFile(pdfBlob);
    return {fileId:arq.getId(), hash:sha256Bytes_(pdfBlob.getBytes()), nome:nomeArq};
  } finally { try{ copia.setTrashed(true); }catch(err){} }
}

/* Teste no editor: gera um PDF de exemplo e loga o link. */
function testarGeracaoPdf(){
  var e = {id:"TESTE-0001", nome:"TESTE DE GERACAO", cpf:"11144477735",
    empresa:EMPRESAS[0].nome, cnpj:EMPRESAS[0].cnpj, cidade:"Volta Redonda",
    item:ITEM_PADRAO, valorReposicao:VALOR_PADRAO};
  var png = Utilities.base64Decode("iVBORw0KGgoAAAANSUhEUgAAAGQAAAAZCAYAAAAaKDCBAAAAOklEQVR4nO3BAQ0AAADCoPdPbQ8HFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAeDwABmQkGAAAAAElFTkSuQmCC");
  var r = gerarPdf_(e, Utilities.newBlob(png,"image/png","assin.png"),
    {assinaturaId:"ASS-TESTE", dataHora:agora_(), ip:"0.0.0.0", userAgent:"editor"});
  Logger.log("PDF: https://drive.google.com/file/d/"+r.fileId+"/view\nHash: "+r.hash);
  return r;
}

/* ===================== DIAGNÓSTICO =====================
   Rode no editor e leia o "Registro de execução".            */
function diagnosticar(){
  var cfg = CFG();

  Logger.log("— BANCO —");
  Logger.log("SHEET_ID: "+(cfg.SHEET_ID||"VAZIO — rode instalar()"));
  try{ Logger.log("Planilha: "+SpreadsheetApp.openById(cfg.SHEET_ID).getName()); }
  catch(e){ Logger.log("Planilha inacessível"); }
  Object.keys(ABAS).forEach(function(a){
    try{ Logger.log("  aba "+a+": "+Math.max(0, sh_(a).getLastRow()-1)+" registro(s)"); }
    catch(e){ Logger.log("  aba "+a+": NÃO EXISTE — rode instalar()"); }
  });

  Logger.log("— SITE DE ASSINATURA —");
  Logger.log("APP_BASE: "+(cfg.APP_BASE||"VAZIO"));
  try{ Logger.log("  link de exemplo: "+linkDe_("EXEMPLODETOKEN123456789012345678")); }
  catch(e){ Logger.log("  PROBLEMA: "+e.message); }

  Logger.log("— MODELO DO TERMO —");
  Logger.log("TEMPLATE_DOC_ID: "+(cfg.TEMPLATE_DOC_ID||"VAZIO"));
  try{
    var n = textoModelo_().length;
    Logger.log("  modelo lido: "+n+" linha(s)");
    Logger.log("  "+conferirModelo());
  }catch(e){ Logger.log("  PROBLEMA: "+e.message); }

  Logger.log("— PASTAS DO DRIVE —");
  [["PASTA_PDF_ID",cfg.PASTA_PDF_ID],["PASTA_ASSIN_ID",cfg.PASTA_ASSIN_ID],["PASTA_TMP_ID",cfg.PASTA_TMP_ID]]
    .forEach(function(p){
      if(!p[1]){ Logger.log("  "+p[0]+": VAZIO — rode instalar()"); return; }
      try{ Logger.log("  "+p[0]+": "+DriveApp.getFolderById(p[1]).getName()); }
      catch(e){ Logger.log("  "+p[0]+": INACESSÍVEL"); }
    });

  Logger.log("— ACESSO DO RH —");
  var u = buscar_("usuarios", function(x){ return String(x.email).toLowerCase()===ADMIN_LOGIN; });
  if(!u){ Logger.log("Usuário "+ADMIN_LOGIN+" NÃO existe — rode instalar() ou resetarAdmin()"); }
  else{
    Logger.log("Usuário "+ADMIN_LOGIN+" | ativo: "+u.ativo+
               " | senha padrão confere: "+(hashSenha_(ADMIN_SENHA,u.salt)===u.senhaHash));
    Logger.log("Tentativas bloqueadas: "+(CacheService.getScriptCache().get("rl_"+sha256_(ADMIN_LOGIN).substr(0,20))||0));
    try{ Logger.log("apiLogin: OK, token "+apiLogin(ADMIN_LOGIN,ADMIN_SENHA).token.substr(0,8)+"..."); }
    catch(e){ Logger.log("apiLogin FALHOU: "+e.message); }
  }
  return "ver Registro de execução";
}
