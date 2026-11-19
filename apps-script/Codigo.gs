/**
 * ════════════════════════════════════════════════════════════════
 *  PORTAL ASPIRANTES 2027 — API (Google Apps Script)  ·  v1.0
 * ════════════════════════════════════════════════════════════════
 *  O MESMO arquivo serve para a planilha de dev e a de produção:
 *  o script descobre sozinho em qual planilha está rodando.
 *
 *  O que ele faz:
 *   • Lê as pastas de disciplinas no Drive → cada subpasta vira uma aba
 *     (subpastas dentro de uma disciplina viram subgrupos)
 *   • Deixa cada arquivo novo como "qualquer pessoa com o link"
 *   • Pasta "quadro" → mostra a imagem mais recente como Quadro de Trabalho
 *   • Avisos: publicados pelo próprio site, expiram sozinhos
 *   • Aniversariantes: lidos da aba "aniversariantes"
 *   • Materiais antigos da aba "materiais" continuam aparecendo
 *     (menos os marcados com ativo = nao)
 *   • A senha da turma é conferida aqui no servidor
 *
 *  Para esconder um arquivo do site: apague-o ou mova para uma pasta
 *  cujo nome começa com "_" (ex.: "_arquivados").
 * ════════════════════════════════════════════════════════════════
 */

// ── AMBIENTES (identificados pelo ID da planilha) ───────────────
const AMBIENTES = {
  // DEV
  '1HevQE_BQXM6jUFXMlJcdu73LK8sDQ8VCD7aOCS9Gkr8': {
    nome: 'dev',
    pasta: '1wG9uDFFROa2ziGjiU3ZEyFmNH9NV8YN_',
    senhaTurma: 'dev',
    senhaAviso: 'novidadenova',
  },
  // PRODUÇÃO
  '1C4BekWGg4Gd01eAIIUexQ2QUxJRMIRtpYYJXcfADH_4': {
    nome: 'prod',
    pasta: '1aXj39m_1eyd3kmlQTYdfSUPXZx2gZt2E',
    senhaTurma: 'fenamissao',
    senhaAviso: 'novidadenova',
  },
};

// Nome bonito das abas (pastas que não estão aqui usam o próprio nome)
const NOMES_DISCIPLINAS = {
  fluidos:     'Mecânica dos Fluidos',
  direito:     'Direito Penal Militar',
  estatistica: 'Estatística',
  fisica:      'Física 3',
  ensino:      'Téc. Prática de Ensino',
  estruturais: 'Sistemas Estruturais',
  incendio2:   'Incêndio II',
  comunicacao: 'Comunicação e Oratória',
};

const PASTA_QUADRO     = 'quadro';   // nome da pasta do Quadro de Trabalho
const DIAS_AVISO       = 7;          // validade padrão de um aviso
const CACHE_SEGUNDOS   = 300;        // 5 minutos
const FUSO             = 'America/Recife';


// ════════════════════════════════════════════════════════════════
//  ENTRADAS DO WEB APP
// ════════════════════════════════════════════════════════════════

function doGet(e) {
  const p   = (e && e.parameter) || {};
  const amb = ambiente_();
  if (!amb) return json_({ ok: false, erro: 'ambiente' });
  if (p.senha !== amb.senhaTurma) return json_({ ok: false, erro: 'senha' });

  const cache = CacheService.getScriptCache();
  const chave = 'dados_' + amb.nome;
  if (!p.fresco) {
    const salvo = cache.get(chave);
    if (salvo) return ContentService.createTextOutput(salvo).setMimeType(ContentService.MimeType.JSON);
  }

  const dados = montarDados_(amb);
  const texto = JSON.stringify(dados);
  try { cache.put(chave, texto, CACHE_SEGUNDOS); } catch (err) { /* grande demais p/ cache */ }
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const amb = ambiente_();
  let corpo = {};
  try { corpo = JSON.parse(e.postData.contents); } catch (err) {}
  if (!amb) return json_({ ok: false, erro: 'ambiente' });
  if (corpo.senha !== amb.senhaAviso) return json_({ ok: false, erro: 'senha' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (corpo.acao === 'verificar') return json_({ ok: true });
    if (corpo.acao === 'novoAviso')   return json_(novoAviso_(corpo));
    if (corpo.acao === 'removerAviso') return json_(removerAviso_(corpo));
    return json_({ ok: false, erro: 'acao' });
  } finally {
    lock.releaseLock();
    CacheService.getScriptCache().remove('dados_' + amb.nome);
  }
}


// ════════════════════════════════════════════════════════════════
//  MONTAGEM DOS DADOS
// ════════════════════════════════════════════════════════════════

function montarDados_(amb) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultado = {
    ok: true,
    ambiente: amb.nome,
    geradoEm: Date.now(),
    disciplinas: [],
    quadro: null,
    avisos: [],
    aniversariantes: [],
  };

  // ── Drive ────────────────────────────────────────────────────
  const porDisciplina = {};
  const idsVistos = {};
  try {
    const raiz = DriveApp.getFolderById(amb.pasta);
    const pastas = raiz.getFolders();
    while (pastas.hasNext()) {
      const pasta = pastas.next();
      const nome  = pasta.getName().trim();
      if (ignorar_(nome)) continue;

      if (slug_(nome) === PASTA_QUADRO) {
        resultado.quadro = quadroDaPasta_(pasta);
        continue;
      }

      const id = slug_(nome);
      const disc = porDisciplina[id] || (porDisciplina[id] = {
        id: id,
        nome: NOMES_DISCIPLINAS[id] || nome,
        itens: [],
      });
      lerArquivos_(pasta, '', disc.itens, idsVistos);

      // subpastas viram subgrupos
      const subs = pasta.getFolders();
      while (subs.hasNext()) {
        const sub = subs.next();
        if (ignorar_(sub.getName())) continue;
        lerArquivos_(sub, sub.getName().trim(), disc.itens, idsVistos);
      }
    }
  } catch (err) {
    resultado.avisoDrive = String(err);
  }

  // ── Materiais antigos da planilha ────────────────────────────
  linhas_(ss, 'materiais').forEach(function (r) {
    if (!r.materia || !r.nome || !r.id) return;
    if ((r.ativo || '').toLowerCase() === 'nao' || (r.ativo || '').toLowerCase() === 'não') return;
    const id = slug_(r.materia);
    const disc = porDisciplina[id] || (porDisciplina[id] = {
      id: id, nome: NOMES_DISCIPLINAS[id] || r.materia, itens: [],
    });
    const ehLink = /^https?:\/\//.test(r.id);
    const arquivoId = ehLink ? r.id : r.id.trim();
    if (idsVistos[arquivoId]) return;
    idsVistos[arquivoId] = true;
    disc.itens.push({
      nome: r.nome,
      tipo: r.tipo || 'Arquivo',
      grupo: '',
      ver: ehLink ? r.id : 'https://drive.google.com/file/d/' + arquivoId + '/view',
      baixar: ehLink ? '' : 'https://drive.google.com/uc?export=download&id=' + arquivoId,
      data: dataDe_(r.data),
    });
  });

  // ordem: disciplinas conhecidas primeiro (na ordem do mapa), depois as novas
  const ordem = Object.keys(NOMES_DISCIPLINAS);
  resultado.disciplinas = Object.keys(porDisciplina)
    .map(function (k) { return porDisciplina[k]; })
    .sort(function (a, b) {
      const ia = ordem.indexOf(a.id), ib = ordem.indexOf(b.id);
      if (ia !== -1 || ib !== -1) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
      return a.nome.localeCompare(b.nome, 'pt-BR');
    });
  resultado.disciplinas.forEach(function (d) {
    d.itens.sort(function (a, b) { return (b.data || 0) - (a.data || 0); });
  });

  // ── Quadro antigo (planilha) se não houver pasta "quadro" ────
  if (!resultado.quadro) {
    const q = linhas_(ss, 'quadro').filter(function (r) { return r.nome && r.link; });
    if (q.length) {
      const u = q[q.length - 1];
      const m = u.link.match(/\/d\/([\w-]+)/);
      resultado.quadro = {
        nome: u.nome,
        imagem: m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w2400' : '',
        ver: u.link.replace('/preview', '/view'),
        data: null,
      };
    }
  }

  // ── Avisos ───────────────────────────────────────────────────
  const agora = Date.now();
  resultado.avisos = linhas_(ss, 'avisos')
    .filter(function (r) { return r.aviso && (r.ativo || '').toLowerCase() === 'sim'; })
    .map(function (r) {
      const data = dataDe_(r.data);
      let expira = dataDe_(r.expira);
      if (!expira && data) expira = data + DIAS_AVISO * 86400000;
      return { linha: r._linha, texto: r.aviso, data: data, expira: expira };
    })
    .filter(function (a) { return !a.expira || a.expira > agora; })
    .sort(function (a, b) { return (b.data || 0) - (a.data || 0); });

  // ── Aniversariantes ──────────────────────────────────────────
  resultado.aniversariantes = linhas_(ss, 'aniversariantes')
    .filter(function (r) { return r.nome && r.data; })
    .map(function (r) {
      const p = String(r.data).split(/[\/\-.]/);
      return { nome: r.nome, dia: parseInt(p[0], 10), mes: parseInt(p[1], 10) };
    })
    .filter(function (a) { return a.dia && a.mes; });

  return resultado;
}

function lerArquivos_(pasta, grupo, destino, idsVistos) {
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (ignorar_(f.getName())) continue;
    const id = f.getId();
    if (idsVistos[id]) continue;
    idsVistos[id] = true;
    garantirPublico_(f);
    const info = tipo_(f);
    destino.push({
      nome: f.getName().replace(/\.[^/.]+$/, ''),
      tipo: info.tipo,
      grupo: grupo,
      ver: info.nativo ? f.getUrl() : 'https://drive.google.com/file/d/' + id + '/view',
      baixar: info.nativo ? '' : 'https://drive.google.com/uc?export=download&id=' + id,
      data: f.getDateCreated().getTime(),
    });
  }
}

function quadroDaPasta_(pasta) {
  let melhor = null;
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    if (ignorar_(f.getName())) continue;
    const mime = f.getMimeType();
    if (mime.indexOf('image/') !== 0 && mime !== MimeType.PDF) continue;
    if (!melhor || f.getDateCreated() > melhor.getDateCreated()) melhor = f;
  }
  if (!melhor) return null;
  garantirPublico_(melhor);
  const id = melhor.getId();
  return {
    nome: melhor.getName().replace(/\.[^/.]+$/, ''),
    imagem: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w2400',
    ver: 'https://drive.google.com/file/d/' + id + '/view',
    data: melhor.getDateCreated().getTime(),
  };
}


// ════════════════════════════════════════════════════════════════
//  AVISOS (escrita)
// ════════════════════════════════════════════════════════════════

function novoAviso_(corpo) {
  const texto = String(corpo.texto || '').trim();
  if (!texto) return { ok: false, erro: 'vazio' };
  if (texto.length > 1000) return { ok: false, erro: 'longo' };

  const aba = abaAvisos_();
  const cab = cabecalho_(aba);
  const agora = new Date();
  const linha = new Array(cab.length).fill('');
  linha[cab.indexOf('aviso')] = texto;
  linha[cab.indexOf('ativo')] = 'sim';
  linha[cab.indexOf('data')]  = Utilities.formatDate(agora, FUSO, 'dd/MM/yyyy HH:mm');
  const dias = parseInt(corpo.dias, 10);
  if (dias > 0) {
    linha[cab.indexOf('expira')] = Utilities.formatDate(
      new Date(agora.getTime() + dias * 86400000), FUSO, 'dd/MM/yyyy HH:mm');
  } else if (corpo.dias === 0 || corpo.dias === '0') {
    linha[cab.indexOf('expira')] = 'nunca';
  }
  aba.appendRow(linha);
  return { ok: true };
}

function removerAviso_(corpo) {
  const aba = abaAvisos_();
  const cab = cabecalho_(aba);
  const n = parseInt(corpo.linha, 10);
  if (!n || n < 2 || n > aba.getLastRow()) return { ok: false, erro: 'linha' };
  aba.getRange(n, cab.indexOf('ativo') + 1).setValue('nao');
  return { ok: true };
}

function abaAvisos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('avisos');
  const cab = cabecalho_(aba);
  if (cab.indexOf('expira') === -1) aba.getRange(1, cab.length + 1).setValue('expira');
  return aba;
}


// ════════════════════════════════════════════════════════════════
//  UTILITÁRIOS
// ════════════════════════════════════════════════════════════════

function ambiente_() {
  try {
    return AMBIENTES[SpreadsheetApp.getActiveSpreadsheet().getId()] || null;
  } catch (err) { return null; }
}

function linhas_(ss, nomeAba) {
  const aba = ss.getSheetByName(nomeAba);
  if (!aba || aba.getLastRow() < 2) return [];
  const valores = aba.getDataRange().getDisplayValues();
  const cab = valores[0].map(function (h) { return String(h).trim().toLowerCase(); });
  return valores.slice(1).map(function (v, i) {
    const o = { _linha: i + 2 };
    cab.forEach(function (h, j) { if (h) o[h] = String(v[j] || '').trim(); });
    return o;
  });
}

function cabecalho_(aba) {
  return aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
}

// "dd/MM/yyyy HH:mm" → timestamp (ms). "nunca" ou vazio → null
function dataDe_(s) {
  if (!s || /nunca/i.test(s)) return null;
  const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  let ano = parseInt(m[3], 10); if (ano < 100) ano += 2000;
  const iso = ano + '-' + pad_(m[2]) + '-' + pad_(m[1]) + 'T' + pad_(m[4] || 0) + ':' + pad_(m[5] || 0) + ':00-03:00';
  const t = new Date(iso).getTime();
  return isNaN(t) ? null : t;
}

function pad_(n) { return ('0' + n).slice(-2); }

function slug_(s) {
  return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '').trim();
}

function ignorar_(nome) { return /^[_.]/.test(String(nome).trim()); }

function garantirPublico_(f) {
  try {
    if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK &&
        f.getSharingAccess() !== DriveApp.Access.ANYONE) {
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (err) { /* sem permissão para alterar: segue */ }
}

function tipo_(f) {
  const mime = f.getMimeType();
  const nome = f.getName().toLowerCase();
  const nativos = {};
  nativos[MimeType.GOOGLE_DOCS]   = 'Documento Google';
  nativos[MimeType.GOOGLE_SHEETS] = 'Planilha';
  nativos[MimeType.GOOGLE_SLIDES] = 'Apresentação';
  nativos[MimeType.GOOGLE_FORMS]  = 'Formulário';
  if (nativos[mime]) return { tipo: nativos[mime], nativo: true };
  if (mime === MimeType.PDF) return { tipo: 'PDF', nativo: false };
  if (/word/.test(mime) || /\.docx?$/.test(nome)) return { tipo: 'Word', nativo: false };
  if (/presentation|powerpoint/.test(mime) || /\.pptx?$/.test(nome)) return { tipo: 'PowerPoint', nativo: false };
  if (/sheet|excel/.test(mime) || /\.xlsx?$/.test(nome)) return { tipo: 'Excel', nativo: false };
  if (mime === MimeType.PLAIN_TEXT || /\.txt$/.test(nome)) return { tipo: 'Texto', nativo: false };
  if (mime.indexOf('image/') === 0) return { tipo: 'Imagem', nativo: false };
  if (mime.indexOf('video/') === 0) return { tipo: 'Vídeo', nativo: false };
  if (mime.indexOf('audio/') === 0) return { tipo: 'Áudio', nativo: false };
  if (/zip|rar|7z/.test(mime)) return { tipo: 'Compactado', nativo: false };
  return { tipo: 'Arquivo', nativo: false };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════════
//  GATILHO ANTIGO: data automática ao editar a planilha à mão
// ════════════════════════════════════════════════════════════════

function onEdit(e) {
  const aba = e.source.getActiveSheet();
  const linha = e.range.getRow();
  const coluna = e.range.getColumn();
  if (coluna !== 1 || linha === 1) return;
  const formato = Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm');
  if (aba.getName() === 'avisos')    aba.getRange(linha, 3).setValue(formato);
  if (aba.getName() === 'materiais') aba.getRange(linha, 5).setValue(formato);
}


// ════════════════════════════════════════════════════════════════
//  TESTE MANUAL (rode no editor para conferir no log)
// ════════════════════════════════════════════════════════════════

function testar() {
  const amb = ambiente_();
  Logger.log('Ambiente: ' + (amb ? amb.nome : 'NÃO RECONHECIDO'));
  if (!amb) return;
  const d = montarDados_(amb);
  Logger.log('Disciplinas: ' + d.disciplinas.map(function (x) { return x.nome + ' (' + x.itens.length + ')'; }).join(', '));
  Logger.log('Quadro: ' + (d.quadro ? d.quadro.nome : '—'));
  Logger.log('Avisos ativos: ' + d.avisos.length);
  Logger.log('Aniversariantes: ' + d.aniversariantes.length);
  if (d.avisoDrive) Logger.log('Erro no Drive: ' + d.avisoDrive);
}
