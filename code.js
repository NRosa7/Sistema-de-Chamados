/**
 * ==========================================================================
 * Kanban HelpDesk Geomembrana - Google Apps Script Backend (PRODUÇÃO)
 * ==========================================================================
 */

const NOME_ABA_CHAMADOS = 'Chamados';
const NOME_ABA_CONFIG = 'Config';
const NOME_ABA_FORMULARIO = 'Respostas ao formulário 1'; 

function doGet(e) {
  if (e.parameter.tela === 'formulario') {
    let template = HtmlService.createTemplateFromFile('formulario');
    try {
      template.emailLogado = Session.getActiveUser().getEmail() || '';
    } catch(err) {
      template.emailLogado = '';
    }
    
    return template.evaluate()
      .setTitle('Geomembrana - Novo Chamado')
      .setFaviconUrl('https://www.google.com/images/branding/product/ico/googleg_lodp.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Geomembrana - Help Desk')
    .setFaviconUrl('https://www.google.com/images/branding/product/ico/googleg_lodp.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function receberDadosFormularioHTML(dados) {
  try {
    let fileUrls = [];
    
    if (dados.arquivos && dados.arquivos.length > 0) {
       const pastaId = '1vA5S_K4nPARJIWx_XM8X0p3fPMbWEe77'; 
       const folder = DriveApp.getFolderById(pastaId);
       
       for (let i = 0; i < dados.arquivos.length; i++) {
         let arq = dados.arquivos[i];
         let base64Data = arq.base64;
         if (base64Data.indexOf(',') > -1) {
           base64Data = base64Data.split(',')[1];
         }
         
         const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), arq.mimeType, arq.name);
         const file = folder.createFile(blob);
         
         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         fileUrls.push(file.getUrl());
       }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let abaOrigem = ss.getSheetByName(NOME_ABA_FORMULARIO);
    if (!abaOrigem) abaOrigem = ss.insertSheet(NOME_ABA_FORMULARIO);

    const novaLinha = [
      new Date(),                     
      dados.email || '',              
      dados.nome || '',               
      dados.departamento || '',       
      dados.categoria || 'Outros',    
      dados.descricao || '',          
      fileUrls.length > 0 ? fileUrls.join(', ') : '', 
      dados.tipo || '',               
      '',                             
      ''                              
    ];

    abaOrigem.appendRow(novaLinha);
    SpreadsheetApp.flush();

    processarChamadosParaKanban();

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getUsuarioAtual() {
  try {
    return Session.getActiveUser().getEmail() || 'Sistema/Admin';
  } catch(e) {
    return 'Sistema/Admin';
  }
}

function getAbaChamados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(NOME_ABA_CHAMADOS);
  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_CHAMADOS);
    const cabecalhos = [
      'Timestamp', 'ID Chamado', 'Solicitante Nome', 'Solicitante Email', 'Setor',
      'Categoria', 'Prioridade', 'Título', 'Descrição', 'Responsável Email',
      'Status', 'Data Previsão', 'Data Atualização', 'Histórico (JSON)', 'Observações'
    ];
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]).setFontWeight('bold').setBackground('#047857').setFontColor('#FFFFFF');
  }
  return aba;
}

function getAbaConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(NOME_ABA_CONFIG);
  
  const categoriasCorretas = [
    "Cigam", "Impressora", "Internet", "Login/senha", "E-mail", "Drive", "Arquivos", 
    "Planilhas", "Licenças", "Câmeras", "Vivo Voz Negócio", "MaaxTalk", 
    "Pré-nota", "Melhorias", "Celular/Notebook", "Outros"
  ];

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_CONFIG);
    aba.getRange(1, 1, 1, 3).setValues([['Responsável Nome', 'Responsável Email', 'Categorias']]).setFontWeight('bold');
    aba.getRange(2, 1, 2, 2).setValues([
      ['Nicolas Rosa', 'nicolas.rosa@geomembrana.com.br'],
      ['Guilherme Jeronymo', 'guilherme.jeronymo@geomembrana.com.br']
    ]);
    for(let c = 0; c < categoriasCorretas.length; c++) {
      aba.getRange(c + 2, 3).setValue(categoriasCorretas[c]);
    }
  }
  return aba;
}

function encontrarLinhaPorId(aba, idProcurado) {
  const lastRow = aba.getLastRow();
  if (lastRow <= 1) return -1;
  
  const dados = aba.getRange(2, 2, lastRow - 1, 7).getValues();
  const numBusca = parseInt(String(idProcurado).replace(/\D/g, ''), 10);
  if (isNaN(numBusca)) return -1; 
  
  for (let k = dados.length - 1; k >= 0; k--) {
    let valId = String(dados[k][0]).trim();
    let valTit = String(dados[k][6]).trim();
    
    let numId = parseInt(valId.replace(/\D/g, ''), 10);
    let matchTit = valTit.match(/os\s*(\d+)/i);
    let numTit = parseInt(matchTit ? matchTit[1] : valTit.replace(/\D/g, ''), 10);
    
    if (!isNaN(numTit) && numTit === numBusca) return k + 2;
    if (!isNaN(numId) && numId === numBusca) return k + 2;
  }
  return -1;
}

function getTodosChamados() {
  try {
    const aba = getAbaChamados();
    const lastRow = aba.getLastRow();
    if (lastRow <= 1) return JSON.stringify([]);
    
    const dados = aba.getRange(2, 1, lastRow - 1, 15).getValues();
    const lista = [];
    const osLidas = new Set(); 
    
    for (let i = dados.length - 1; i >= 0; i--) {
      const linha = dados[i];
      const linhaSheetNum = i + 2;
      
      if (!linha[1] && !linha[7]) continue; 
      
      let rawId = String(linha[1] || '').trim();
      let title = String(linha[7] || '').trim();
      
      let numOS = 0;
      let titleMatch = title.match(/OS\s*(\d+)/i);
      
      if (titleMatch) {
          numOS = parseInt(titleMatch[1], 10);
      } else {
          numOS = parseInt(rawId.replace(/\D/g, ''), 10);
      }
      
      if (isNaN(numOS) || numOS <= 0) continue; 
      
      let idLimpo = 'OS ' + numOS;

      if (osLidas.has(idLimpo)) continue;
      osLidas.add(idLimpo);

      let historico = [];
      if (linha[13]) {
        try { historico = JSON.parse(linha[13]); } 
        catch (e) { historico = [{ data: new Date().toISOString(), acao: String(linha[13]), usuario: 'sistema' }]; }
      }

      let statusOriginal = String(linha[10] || 'NOVO').trim();
      if (statusOriginal === 'AGUARDANDO USUÁRIO') {
        statusOriginal = 'AGUARDANDO RETORNO';
      }

      let respEmail = String(linha[9] || '').trim();
      let respNomeFormatado = '';
      if (respEmail) {
        if (respEmail.toLowerCase().includes('guilherme')) {
          respNomeFormatado = 'Guilherme Jeronymo';
        } else if (respEmail.toLowerCase().includes('nicolas')) {
          respNomeFormatado = 'Nicolas Rosa';
        } else {
          respNomeFormatado = respEmail.split('@')[0];
        }
      }

      const item = {
        id: idLimpo,
        solicitante: {
          nome: String(linha[2] || 'Solicitante').trim(),
          email: String(linha[3] || '').trim(),
          setor: String(linha[4] || '').trim()
        },
        categoria: String(linha[5] || 'Outros').trim(),
        prioridade: String(linha[6] || 'Média').trim(),
        titulo: title || 'Sem título',
        descricao: String(linha[8] || '').trim(),
        responsavel: respEmail ? { nome: respNomeFormatado, email: respEmail } : null,
        status: statusOriginal,
        dataPrevisao: linha[11] ? (linha[11] instanceof Date ? linha[11].toISOString().split('T')[0] : String(linha[11])) : '',
        dataAtualizacao: linha[12] ? (linha[12] instanceof Date ? linha[12].toISOString() : String(linha[12])) : new Date().toISOString(),
        dataCriacao: linha[0] ? (linha[0] instanceof Date ? linha[0].toISOString() : String(linha[0])) : new Date().toISOString(),
        historico: historico,
        observacoes: String(linha[14] || '').trim(),
        linhaSheet: linhaSheetNum
      };
      
      lista.unshift(item); 
    }
    return JSON.stringify(lista);
  } catch (err) {
    throw new Error('Falha ao buscar chamados da planilha: ' + err.message);
  }
}

function atualizarChamado(dadosChamado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA_CHAMADOS);
    
    const linhaIndex = encontrarLinhaPorId(aba, dadosChamado.id);
    if (linhaIndex === -1) throw new Error('Chamado não localizado na planilha pelo ID: ' + dadosChamado.id);
    
    const agora = new Date();
    const historicoJson = JSON.stringify(dadosChamado.historico || []);
    const respEmail = dadosChamado.responsavel ? dadosChamado.responsavel.email : '';
    const solNome = dadosChamado.solicitante ? dadosChamado.solicitante.nome : '';
    const solEmail = dadosChamado.solicitante ? dadosChamado.solicitante.email : '';
    const solSetor = dadosChamado.solicitante ? dadosChamado.solicitante.setor : '';
    
    aba.getRange(linhaIndex, 2, 1, 14).setValues([[ 
      dadosChamado.id || '', solNome, solEmail, solSetor, 
      dadosChamado.categoria || 'Outros', dadosChamado.prioridade || 'Média', 
      dadosChamado.titulo || '', dadosChamado.descricao || '', 
      respEmail, dadosChamado.status || 'NOVO', dadosChamado.dataPrevisao || '', 
      agora, historicoJson, dadosChamado.observacoes || '' 
    ]]);
    
    try {
      const abaForm = ss.getSheetByName(NOME_ABA_FORMULARIO);
      if (abaForm) {
        const numOSBusca = parseInt(String(dadosChamado.id).replace(/\D/g, ''), 10);
        if (!isNaN(numOSBusca)) {
          const formLastRow = abaForm.getLastRow();
          if (formLastRow > 1) {
            const osValues = abaForm.getRange(2, 10, formLastRow - 1, 1).getValues();
            for (let f = osValues.length - 1; f >= 0; f--) {
              let formValPuro = parseInt(String(osValues[f][0]).replace(/\D/g, ''), 10);
              if (formValPuro === numOSBusca) {
                abaForm.getRange(f + 2, 9).setValue(dadosChamado.status);
                break;
              }
            }
          }
        }
      }
    } catch(e) { }

    if (dadosChamado.status === 'RESOLVIDO' && solEmail && dadosChamado.notificarResolucao === true) {
      enviarEmailNotificacao('RESOLVIDO', dadosChamado);
    }
    
    SpreadsheetApp.flush();
    return { success: true, id: dadosChamado.id };
  } catch (err) {
    throw new Error('Falha ao atualizar chamado: ' + err.message);
  }
}

function criarChamado(dadosChamado) {
  try {
    const aba = getAbaChamados();
    const nextRow = aba.getLastRow() + 1;
    let newId = dadosChamado.id;

    if (!newId) {
      let maxIdNum = 0;
      if (aba.getLastRow() > 1) {
        const ids = aba.getRange(2, 2, aba.getLastRow() - 1, 7).getValues();
        ids.forEach(r => {
          let matchTit = String(r[6] || '').match(/os\s*(\d+)/i);
          let num = matchTit ? parseInt(matchTit[1], 10) : parseInt(String(r[0] || '').replace(/\D/g, ''), 10);
          if (!isNaN(num) && num > maxIdNum) maxIdNum = num;
        });
      }
      newId = 'OS ' + (maxIdNum + 1);
    }

    const agora = new Date();
    const historicoInicial = [{ data: agora.toISOString(), acao: 'Chamado aberto na plataforma', usuario: getUsuarioAtual() }];
    
    const novaLinha = [
      agora, newId, dadosChamado.solicitanteNome, dadosChamado.solicitanteEmail, dadosChamado.solicitanteSetor,
      dadosChamado.categoria, dadosChamado.prioridade, dadosChamado.titulo, dadosChamado.descricao, '',
      'NOVO', '', agora, JSON.stringify(historicoInicial), ''
    ];
    
    aba.getRange(nextRow, 1, 1, novaLinha.length).setValues([novaLinha]);
    
    const objetoCriado = { id: newId, titulo: dadosChamado.titulo, solicitante: { nome: dadosChamado.solicitanteNome, email: dadosChamado.solicitanteEmail, setor: dadosChamado.solicitanteSetor }, categoria: dadosChamado.categoria, descricao: dadosChamado.descricao };
    
    if (dadosChamado.solicitanteEmail) enviarEmailNotificacao('NOVO', objetoCriado);
    
    SpreadsheetApp.flush();
    return { success: true, id: newId };
  } catch (err) {
    throw new Error('Falha ao criar novo chamado: ' + err.message);
  }
}

function excluirChamado(idChamado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(NOME_ABA_CHAMADOS);
    const lastRow = aba.getLastRow();
    
    if (lastRow <= 1) return { success: true, msg: 'Planilha vazia' };
    
    const dados = aba.getRange(2, 2, lastRow - 1, 7).getValues(); 
    const numBusca = parseInt(String(idChamado).replace(/\D/g, ''), 10);
    
    let apagouAlgum = false;
    
    for (let k = dados.length - 1; k >= 0; k--) {
      let idCol = String(dados[k][0]).trim();
      let titCol = String(dados[k][6]).trim();
      
      let numId = parseInt(idCol.replace(/\D/g, ''), 10);
      let matchTit = titCol.match(/os\s*(\d+)/i);
      let numTit = parseInt(matchTit ? matchTit[1] : titCol.replace(/\D/g, ''), 10);
      
      let match = false;
      if (!isNaN(numBusca) && !isNaN(numTit) && numTit === numBusca) match = true;
      else if (!isNaN(numBusca) && !isNaN(numId) && numId === numBusca) match = true;
      
      if (match) {
        aba.deleteRow(k + 2);
        apagouAlgum = true;
      }
    }
    
    try {
       const abaForm = ss.getSheetByName(NOME_ABA_FORMULARIO);
       if (abaForm && !isNaN(numBusca)) {
         const formLastRow = abaForm.getLastRow();
         if (formLastRow > 1) {
           const osValues = abaForm.getRange(2, 10, formLastRow - 1, 1).getValues();
           for (let f = osValues.length - 1; f >= 0; f--) {
             let formValPuro = parseInt(String(osValues[f][0]).replace(/\D/g, ''), 10);
             if (formValPuro === numBusca) {
               abaForm.getRange(f + 2, 9).setValue('EXCLUÍDO KANBAN');
               break;
             }
           }
         }
       }
    } catch(e) { }

    if (apagouAlgum) SpreadsheetApp.flush(); 
    return { success: true, id: idChamado };
  } catch (err) {
    throw new Error('Falha ao excluir chamado no Sheets: ' + err.message);
  }
}

function getResponsaveis() {
  const responsaveisPadrao = [
    { nome: 'Nicolas Rosa', email: 'nicolas.rosa@geomembrana.com.br' },
    { nome: 'Guilherme Jeronymo', email: 'guilherme.jeronymo@geomembrana.com.br' }
  ];
  try {
    const aba = getAbaConfig();
    const dados = aba.getRange(2, 1, Math.max(1, aba.getLastRow() - 1), 2).getValues();
    const lista = [];
    dados.forEach(r => { 
      if (r[1]) {
        let emailResp = String(r[1]).trim();
        let nomeResp = String(r[0] || r[1]).trim();
        if (emailResp.toLowerCase().includes('guilherme') || nomeResp.toLowerCase().includes('guilherme')) {
          nomeResp = 'Guilherme Jeronymo';
        } else if (emailResp.toLowerCase().includes('nicolas') || nomeResp.toLowerCase().includes('nicolas')) {
          nomeResp = 'Nicolas Rosa';
        }
        lista.push({ nome: nomeResp, email: emailResp }); 
      }
    });
    return lista.length > 0 ? lista : responsaveisPadrao;
  } catch (err) { 
    return responsaveisPadrao; 
  }
}

function getCategorias() {
  const categoriasCorretas = [
    "Cigam", "Impressora", "Internet", "Login/senha", "E-mail", "Drive", "Arquivos", 
    "Planilhas", "Licenças", "Câmeras", "Vivo Voz Negócio", "MaaxTalk", 
    "Pré-nota", "Melhorias", "Celular/Notebook", "Outros"
  ];
  try {
    const aba = getAbaConfig();
    const dados = aba.getRange(2, 3, Math.max(1, aba.getLastRow() - 1), 1).getValues();
    const lista = [];
    dados.forEach(r => { if (r[0]) lista.push(String(r[0])); });
    return lista.length > 0 ? lista : categoriasCorretas;
  } catch (err) { 
    return categoriasCorretas; 
  }
}

function enviarEmailNotificacao(tipo, chamado) {
  try {
    let solEmail = chamado.solicitante && chamado.solicitante.email ? chamado.solicitante.email : '';
    let solNome = chamado.solicitante && chamado.solicitante.nome ? chamado.solicitante.nome : 'Usuário';
    let solSetor = chamado.solicitante && chamado.solicitante.setor ? chamado.solicitante.setor : '-';
    
    if (!solEmail) return;
    
    let idLimpo = String(chamado.id || '').replace(/GEOM-0*(\d+)/gi, 'OS $1');
    let tituloLimpo = String(chamado.titulo || '').replace(/GEOM-0*(\d+)/gi, 'OS $1');
    
    let assunto = tituloLimpo; 
    let mensagemTopo = '';
    
    if (tipo === 'NOVO') {
      mensagemTopo = `Olá <strong>${solNome}</strong>,<br><br>Seu chamado foi aberto com sucesso e nossa equipe já foi notificada.`;
    } else if (tipo === 'RESOLVIDO') {
      assunto = `[RESOLVIDO] ${tituloLimpo}`;
      mensagemTopo = `Olá <strong>${solNome}</strong>,<br><br>Informamos que o atendimento do seu chamado foi <strong>concluído</strong> com sucesso.`;
      
      if (chamado.comentarioResolucao && chamado.comentarioResolucao.trim() !== '') {
        mensagemTopo += `<br><br>
        <div style="background-color: #f1f5f9; border-left: 4px solid #047857; padding: 12px; border-radius: 4px; color: #333; font-size: 15px;">
          <strong>Comentário da TI:</strong><br>
          ${chamado.comentarioResolucao.replace(/\n/g, '<br>')}
        </div>`;
      }
    } else {
      return; 
    }

    let formatDateTime = function(dateObj) {
      if(!dateObj) return '-';
      var d = new Date(dateObj);
      if (isNaN(d.getTime())) return '-';
      return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear() + " " + ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
    };
    
    let dataCriacao = chamado.dataCriacao ? new Date(chamado.dataCriacao) : new Date();
    let dataHoraFormatada = formatDateTime(dataCriacao);
    let cat = chamado.categoria || 'Outros';
    let desc = chamado.descricao || 'Nenhuma descrição detalhada informada.';

    let htmlTemplate = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #047857; padding: 20px; text-align: center;">
          <h2 style="color: #ffffff; margin: 0; font-size: 22px;">Help Desk Geomembrana</h2>
        </div>
        <div style="padding: 20px; background-color: #f8fafc;">
          <p style="font-size: 16px; margin-bottom: 20px;">${mensagemTopo}</p>
          <div style="background-color: #ffffff; padding: 15px; border-radius: 6px; border-left: 4px solid #047857; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <h3 style="margin-top: 0; color: #047857; font-size: 18px;">Detalhes da Solicitação</h3>
            <p style="margin: 8px 0;"><strong>ID do Chamado:</strong> ${idLimpo}</p>
            <p style="margin: 8px 0;"><strong>Data/Hora:</strong> ${dataHoraFormatada}</p>
            <p style="margin: 8px 0;"><strong>Departamento:</strong> ${solSetor}</p>
            <p style="margin: 8px 0;"><strong>Classificação:</strong> ${cat}</p>
            <p style="margin: 8px 0;"><strong>Descrição:</strong><br><span style="display: inline-block; margin-top: 4px; padding: 10px; background: #f1f5f9; border-radius: 4px; width: 100%; box-sizing: border-box;">${desc.replace(/\n/g, '<br>')}</span></p>
          </div>
          <p style="margin-top: 20px; font-size: 14px; color: #666; text-align: center;">
            Este é um e-mail automático do sistema Kanban Help Desk Geomembrana.
          </p>
        </div>
      </div>
    `;

    MailApp.sendEmail({ to: solEmail, subject: assunto, htmlBody: htmlTemplate });
  } catch (e) { Logger.log('Erro email: ' + e); }
}

function processarChamadosParaKanban() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return; // Aumentado para 15s para garantir concorrência segura

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOrigem = ss.getSheetByName(NOME_ABA_FORMULARIO);
    var abaChamados = getAbaChamados();
    
    if (!abaOrigem) return;

    var lastRowForm = abaOrigem.getLastRow();
    if (lastRowForm <= 1) return;

    var dadosForm = abaOrigem.getDataRange().getValues();
    var dadosChamados = abaChamados.getDataRange().getValues();
    var maxOS = 0;
    
    for (var c = 1; c < dadosChamados.length; c++) {
      var idNum = parseInt(String(dadosChamados[c][1] || '').replace(/\D/g, ''), 10);
      var matchTit = String(dadosChamados[c][7] || '').match(/os\s*(\d+)/i);
      var titNum = matchTit ? parseInt(matchTit[1], 10) : 0;
      
      if (!isNaN(idNum) && idNum > maxOS) maxOS = idNum;
      if (!isNaN(titNum) && titNum > maxOS) maxOS = titNum;
    }
    
    for (var r = 1; r < dadosForm.length; r++) {
      var val = parseInt(String(dadosForm[r][9] || '').replace(/\D/g, ''), 10);
      if (!isNaN(val) && val > maxOS) maxOS = val;
    }

    for (var i = 1; i < dadosForm.length; i++) { 
      var linhaAtual = i + 1;
      var dataHoraCarimbo = dadosForm[i][0]; 
      var statusIntegracao = String(dadosForm[i][8] || '').trim();

      if (!dataHoraCarimbo || statusIntegracao !== '') continue;

      var solicitanteNome = String(dadosForm[i][2] || '').trim();
      var classificacaoProblema = String(dadosForm[i][4] || '').trim();
      var descricaoProblema = String(dadosForm[i][5] || '').trim();

      var isDuplicado = false;
      var strNome = solicitanteNome.toLowerCase();
      var strDesc = descricaoProblema.toLowerCase();

      if (strNome && strDesc) {
        for (var r = 1; r < i; r++) {
          var prevNome = String(dadosForm[r][2] || '').toLowerCase().trim();
          var prevDesc = String(dadosForm[r][5] || '').toLowerCase().trim();
          if (prevNome === strNome && prevDesc === strDesc) {
            isDuplicado = true;
            break;
          }
        }
      }

      if (!isDuplicado && strNome && strDesc) {
        for (var k = 1; k < dadosChamados.length; k++) {
           var nomeDB = String(dadosChamados[k][2] || '').toLowerCase().trim(); 
           var descDB = String(dadosChamados[k][8] || '').toLowerCase().trim(); 
           if (nomeDB === strNome && descDB === strDesc) {
               isDuplicado = true;
               break;
           }
        }
      }

      if (isDuplicado) {
         abaOrigem.getRange(linhaAtual, 9).setValue('DUPLICADO CLONE');
         abaOrigem.getRange(linhaAtual, 10).setValue('IGNORADO');
         continue; 
      }

      var osExistenteStr = String(dadosForm[i][9] || '').replace(/\D/g, ''); 
      var osDestaLinha = parseInt(osExistenteStr, 10);
      
      if (isNaN(osDestaLinha) || osDestaLinha <= 0) {
        maxOS++;
        osDestaLinha = maxOS;
        abaOrigem.getRange(linhaAtual, 10).setValue(osDestaLinha);
      }

      abaOrigem.getRange(linhaAtual, 9).setValue('PROCESSANDO...');
      SpreadsheetApp.flush(); 

      var imagemProblema = String(dadosForm[i][6] || '');
      var problemaNovoRecorrente = String(dadosForm[i][7] || '');
      var enderecoEmail = String(dadosForm[i][1] || '');
      var departamento = String(dadosForm[i][3] || '');

      var novoID = "OS " + osDestaLinha;
      var tituloCerto = novoID + " - " + classificacaoProblema;
      var historicoInicial = [{ data: new Date().toISOString(), acao: 'Chamado aberto na plataforma', usuario: 'Sistema' }];
      
      var novaLinha = [
        dataHoraCarimbo || new Date(), novoID, solicitanteNome, enderecoEmail, departamento,
        classificacaoProblema, "Média", tituloCerto, 
        "Descrição: " + descricaoProblema + "\nLink/Imagem: " + (imagemProblema || "N/A") + "\nTipo: " + problemaNovoRecorrente, 
        '', 'NOVO', '', new Date(), JSON.stringify(historicoInicial), ''
      ];

      try {
        abaChamados.getRange(abaChamados.getLastRow() + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
        abaOrigem.getRange(linhaAtual, 9).setValue('NOVO'); 
        dadosChamados.push(novaLinha); 
        SpreadsheetApp.flush();

        try {
          var assuntoEmail = `[Geomembrana HelpDesk] Chamado ABERTO: ${tituloCerto}`; 
          
          var formatDateTime = function(dateObj) {
            if(!dateObj) return '-';
            var d = new Date(dateObj);
            if(isNaN(d.getTime())) return '-';
            return ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear() + " " + ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2);
          };
          var dataHoraFormatada = formatDateTime(dataHoraCarimbo || new Date());
          
          var linksImagem = [];
          if (imagemProblema && imagemProblema.includes('http')) {
            var urls = imagemProblema.split(',');
            urls.forEach(function(url) {
                var cleanUrl = url.trim();
                if(cleanUrl) {
                    linksImagem.push(`<a href="${cleanUrl}" target="_blank" style="color: #047857; font-weight: bold; text-decoration: underline;">Visualizar Anexo</a>`);
                }
            });
          }
          var linkImagemHTML = linksImagem.length > 0 ? linksImagem.join(' <br> ') : '<span style="color: #666; font-style: italic;">Nenhum anexo.</span>';

          var gerarTemplateHTML = function(mensagemIntroducao) {
            return `
            <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
              <div style="background-color: #047857; padding: 20px; text-align: center;">
                <h2 style="color: #ffffff; margin: 0; font-size: 22px;">Help Desk Geomembrana</h2>
              </div>
              <div style="padding: 20px; background-color: #f8fafc;">
                <p style="font-size: 16px; margin-bottom: 20px;">${mensagemIntroducao}</p>
                <div style="background-color: #ffffff; padding: 15px; border-radius: 6px; border-left: 4px solid #047857; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                  <h3 style="margin-top: 0; color: #047857; font-size: 18px;">Resumo da Solicitação</h3>
                  <p style="margin: 8px 0;"><strong>ID do Chamado:</strong> ${novoID}</p>
                  <p style="margin: 8px 0;"><strong>Data/Hora:</strong> ${dataHoraFormatada}</p>
                  <p style="margin: 8px 0;"><strong>Departamento:</strong> ${departamento}</p>
                  <p style="margin: 8px 0;"><strong>Classificação do Problema:</strong> ${classificacaoProblema}</p>
                  <p style="margin: 8px 0;"><strong>Descrição do Problema:</strong><br><span style="display: inline-block; margin-top: 4px; padding: 10px; background: #f1f5f9; border-radius: 4px; width: 100%; box-sizing: border-box;">${descricaoProblema.replace(/\n/g, '<br>')}</span></p>
                  <p style="margin: 8px 0;"><strong>Anexos:</strong><br>${linkImagemHTML}</p>
                </div>
                <p style="margin-top: 20px; font-size: 14px; color: #666; text-align: center;">
                  Este é um e-mail automático do sistema Kanban Help Desk Geomembrana.
                </p>
              </div>
            </div>
            `;
          };

          var adminEmail = 'nicolas.rosa@geomembrana.com.br, guilherme.jeronymo@geomembrana.com.br';
          var msgAdmin = `Olá <strong>Equipe de TI</strong>,<br><br>Um novo chamado foi aberto no sistema por <strong>${solicitanteNome}</strong> e aguarda atendimento.`;
          MailApp.sendEmail({ 
            to: adminEmail, 
            subject: assuntoEmail, 
            htmlBody: gerarTemplateHTML(msgAdmin) 
          });

          if (enderecoEmail) {
            var msgUser = `Olá <strong>${solicitanteNome}</strong>,<br><br>Sua solicitação foi registrada com sucesso! A nossa equipe de TI já foi notificada e em breve iniciaremos o atendimento.`;
            MailApp.sendEmail({ 
              to: enderecoEmail, 
              subject: assuntoEmail, 
              htmlBody: gerarTemplateHTML(msgUser) 
            });
          }

        } catch (emailErr) {
          Logger.log('Erro ao enviar e-mails HTML do formulário: ' + emailErr);
        }

      } catch(e) {
        abaOrigem.getRange(linhaAtual, 9).setValue('ERRO DE SCRIPT'); 
      }
    } 
  } finally {
    lock.releaseLock(); 
  }
}