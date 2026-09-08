const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const root = __dirname;
const dataFile = path.join(root, 'data', 'chamados.json');
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY;
const categorias = [
  'Cigam', 'Impressora', 'Internet', 'Login/senha', 'E-mail', 'Drive', 'Arquivos',
  'Planilhas', 'Licenças', 'Câmeras', 'Vivo Voz Negócio', 'MaaxTalk', 'Pré-nota',
  'Melhorias', 'Celular/Notebook', 'Outros'
];

function ensureDataFile() {
  const dir = path.dirname(dataFile);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(dataFile)) {
    fs.writeFileSync(dataFile, JSON.stringify([
      {
        id: 'OS 101',
        solicitante: { nome: 'Maria Silva', email: 'maria.silva@geomembrana.com.br', setor: 'Financeiro' },
        categoria: 'E-mail',
        prioridade: 'Alta',
        titulo: 'E-mail com lentidão no Outlook',
        descricao: 'O cliente está com atraso na sincronização do Outlook e imagens não carregam.',
        responsavel: { nome: 'Nicolas Rosa', email: 'nicolas.rosa@geomembrana.com.br' },
        status: 'EM ATENDIMENTO',
        dataPrevisao: new Date(Date.now() + 2 * 86400000).toISOString(),
        dataAtualizacao: new Date().toISOString(),
        dataCriacao: new Date(Date.now() - 86400000).toISOString(),
        historico: [{ data: new Date().toISOString(), acao: 'Chamado criado em ambiente local.', usuario: 'Sistema' }],
        observacoes: 'Verificar regra de correio e testar no cliente.'
      },
      {
        id: 'OS 102',
        solicitante: { nome: 'José Pereira', email: 'jose.pereira@geomembrana.com.br', setor: 'Logística' },
        categoria: 'Internet',
        prioridade: 'Média',
        titulo: 'Rede está instável em fábrica',
        descricao: 'A internet cai em momentos de uso intenso e trava a operação da linha de produção.',
        responsavel: { nome: 'Guilherme Jeronymo', email: 'guilherme.jeronymo@geomembrana.com.br' },
        status: 'NOVO',
        dataPrevisao: new Date(Date.now() + 3 * 86400000).toISOString(),
        dataAtualizacao: new Date().toISOString(),
        dataCriacao: new Date(Date.now() - 2 * 86400000).toISOString(),
        historico: [{ data: new Date().toISOString(), acao: 'Solicitação registrada em ambiente local.', usuario: 'Sistema' }],
        observacoes: 'Verificar link e diagnóstico de rede.'
      }
    ], null, 2));
  }
}

async function supabaseRequest(pathname, options = {}) {
  if (!supabaseUrl || !supabaseKey) {
    throw new Error('Supabase não configurado');
  }

  const url = `${supabaseUrl.replace(/\/$/, '')}${pathname}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Supabase error (${response.status}): ${errorText}`);
  }

  return response.status === 204 ? null : response.json();
}

async function readData() {
  if (supabaseUrl && supabaseKey) {
    try {
      const rows = await supabaseRequest('/rest/v1/chamados?select=*');
      return Array.isArray(rows) ? rows.map(item => item.payload ?? item) : [];
    } catch (error) {
      console.warn('Falha ao ler Supabase, usando fallback local:', error.message);
    }
  }

  ensureDataFile();
  return JSON.parse(fs.readFileSync(dataFile, 'utf8'));
}

async function writeData(data) {
  if (supabaseUrl && supabaseKey) {
    try {
      await supabaseRequest('/rest/v1/chamados', { method: 'DELETE' });
      if (Array.isArray(data) && data.length) {
        const payload = data.map(item => ({ id: item.id, payload: item }));
        await supabaseRequest('/rest/v1/chamados', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
      }
      return;
    } catch (error) {
      console.warn('Falha ao gravar Supabase, usando fallback local:', error.message);
    }
  }

  fs.writeFileSync(dataFile, JSON.stringify(data, null, 2));
}

function gerarProximoId(data) {
  const maiorNumero = data.reduce((maior, chamado) => {
    const numero = Number.parseInt(String(chamado.id || '').replace(/\D/g, ''), 10);
    return Number.isNaN(numero) ? maior : Math.max(maior, numero);
  }, 0);
  return `OS ${maiorNumero + 1}`;
}

function ensureEmailLogsFile() {
  const emailFile = path.join(root, 'data', 'emails.json');
  const dir = path.dirname(emailFile);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(emailFile)) fs.writeFileSync(emailFile, JSON.stringify([], null, 2));
  return emailFile;
}

function registrarEmailConfirmacao(chamado) {
  const emailFile = ensureEmailLogsFile();
  const log = JSON.parse(fs.readFileSync(emailFile, 'utf8'));
  const mensagem = {
    para: chamado.solicitante && chamado.solicitante.email ? chamado.solicitante.email : '',
    assunto: `Confirmação de abertura do chamado ${chamado.id}`,
    corpo: `Prezado(a) ${chamado.solicitante?.nome || 'Solicitante'},\n\nSeu chamado foi aberto com sucesso.\n\nID: ${chamado.id}\nCategoria: ${chamado.categoria || 'Outros'}\nStatus: ${chamado.status || 'NOVO'}\nDescrição: ${chamado.descricao || ''}\n\nA equipe de TI irá analisar sua solicitação e responder em breve.`,
    enviadoEm: new Date().toISOString()
  };

  log.push(mensagem);
  fs.writeFileSync(emailFile, JSON.stringify(log, null, 2));
  return mensagem;
}

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(data));
}

function serveStaticFile(res, filePath) {
  const safePath = path.normalize(filePath);
  const allowed = ['index.html', 'formulario.html', 'style.html', 'api.html', 'modal.html', 'app.html', 'utils.html'];

  const extension = path.extname(safePath).toLowerCase();
  const mime = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.svg': 'image/svg+xml'
  };

  if (!allowed.includes(path.basename(safePath))) {
    res.writeHead(403); res.end('Acesso negado'); return;
  }

  const fullPath = path.join(root, safePath);
  fs.readFile(fullPath, (err, content) => {
    if (err) {
      res.writeHead(404); res.end('Arquivo não encontrado');
      return;
    }
    res.writeHead(200, { 'Content-Type': mime[extension] || 'application/octet-stream' });
    res.end(content);
  });
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');

  const handleRequest = async () => {
    if (url.pathname === '/api/usuario') {
      sendJson(res, 200, { email: 'admin@geomembrana.com.br' });
      return;
    }

    if (url.pathname === '/api/categorias') {
      sendJson(res, 200, categorias);
      return;
    }

    if (url.pathname === '/api/responsaveis') {
      sendJson(res, 200, [
        { nome: 'Nicolas Rosa', email: 'nicolas.rosa@geomembrana.com.br' },
        { nome: 'Guilherme Jeronymo', email: 'guilherme.jeronymo@geomembrana.com.br' }
      ]);
      return;
    }

    if (url.pathname === '/api/chamados') {
      if (req.method === 'GET') {
        sendJson(res, 200, await readData());
        return;
      }

      if (req.method === 'POST') {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', async () => {
          try {
            const payload = JSON.parse(body || '{}');
            const data = await readData();
            const generatedId = payload.id || gerarProximoId(data);
            const isNewTicket = !payload.id;
            const categoria = payload.categoria || 'Outros';
            const descricaoBase = payload.descricao || '';
            const descricao = isNewTicket
              ? `Descrição: ${descricaoBase}\nLink/Imagem: N/A\nTipo: ${payload.tipo || 'Não informado'}`
              : descricaoBase;
            const item = {
              id: generatedId,
              solicitante: payload.solicitante || {
                nome: payload.nome || 'Solicitante',
                email: payload.email || '',
                setor: payload.departamento || ''
              },
              categoria,
              prioridade: payload.prioridade || 'Média',
              titulo: isNewTicket ? `${generatedId} - ${categoria}` : (payload.titulo || 'Solicitação de suporte'),
              descricao,
              responsavel: payload.responsavel || null,
              status: payload.status || 'NOVO',
              dataPrevisao: payload.dataPrevisao || '',
              dataAtualizacao: payload.dataAtualizacao || new Date().toISOString(),
              dataCriacao: payload.dataCriacao || new Date().toISOString(),
              historico: Array.isArray(payload.historico) && payload.historico.length ? payload.historico : [{
                data: new Date().toISOString(),
                acao: 'Chamado criado no sistema local.',
                usuario: 'Sistema'
              }],
              observacoes: payload.observacoes || '',
              linhaSheet: payload.linhaSheet || null
            };

            const existingIndex = data.findIndex(ch => ch.id === generatedId);
            if (existingIndex >= 0) data[existingIndex] = item;
            else data.push(item);

            await writeData(data);
            const emailLog = registrarEmailConfirmacao(item);
            sendJson(res, 200, {
              success: true,
              id: generatedId,
              emailSent: false,
              emailQueued: !!emailLog.para,
              email: emailLog.para
            });
          } catch (error) {
            sendJson(res, 500, { success: false, error: error.message });
          }
        });
        return;
      }
    }

    if (url.pathname.startsWith('/api/chamados/')) {
      const id = decodeURIComponent(url.pathname.replace('/api/chamados/', ''));

      if (req.method === 'DELETE') {
        const data = (await readData()).filter(item => item.id !== id);
        await writeData(data);
        sendJson(res, 200, { success: true });
        return;
      }
    }

    const requestedPath = url.pathname === '/' ? '/index.html' : url.pathname;
    serveStaticFile(res, requestedPath.replace(/^\//, ''));
  };

  handleRequest().catch(error => {
    sendJson(res, 500, { success: false, error: error.message });
  });
});

const port = process.env.PORT || 8000;
server.listen(port, () => {
  console.log(`Servidor local rodando em http://localhost:${port}`);
});
