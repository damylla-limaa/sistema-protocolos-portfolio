/**
 * ATUALIZAR ANDAMENTO – Protocolos da Prefeitura de Foz do Iguaçu (portal Oxy)
 *
 * Menu "Protocolos" na planilha:
 *   - Atualizar andamento agora
 *   - Ativar atualização automática (2x por dia: 7h e 13h) / Desativar
 *
 * Quando um protocolo passa a "Atendido" ou "Respondido", o script cria um
 * lembrete na sua Agenda Google (agenda principal) com o link do processo.
 *
 * Para cada linha da aba "Protocolos" que tenha o link do processo na coluna K,
 * o script consulta o portal e preenche: protocolo, requerente, tipo, objeto,
 * aberto em, setor atual, última movimentação, situação e última consulta.
 * A aba "Movimentações" é refeita com o histórico completo de todos os protocolos.
 *
 * Protocolo novo: cole o link do QR code (https://foz.oxy.elotech.com.br/...)
 * na coluna K de uma linha vazia e clique em "Atualizar andamento agora".
 * O resto da linha é preenchido sozinho. A coluna L (comprovante) não é alterada.
 */

const API = 'https://foz.oxy.elotech.com.br/portal-govdig-api/api/oxy/processo-digital/process/';
const PORTAL = 'https://foz.oxy.elotech.com.br/governo-digital/processo-digital/';
const ABA_PROT = 'Protocolos';
const ABA_MOV = 'Movimentações';
const PASTA_COMPROVANTES = 'ID_PASTA_COMPROVANTES';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Protocolos')
    .addItem('Atualizar andamento agora (inclui novos comprovantes)', 'atualizarAndamento')
    .addItem('Só verificar novos comprovantes na pasta', 'verificarNovosComprovantes')
    .addSeparator()
    .addItem('Ativar atualização automática (7h e 13h)', 'ativarDiario')
    .addItem('Desativar atualização automática', 'desativarDiario')
    .addSeparator()
    .addItem('Ver prontos dos últimos 7 dias', 'verProntosRecentes')
    .addItem('Baixar anexos da Prefeitura agora', 'baixarAnexos')
    .addItem('Arquivar histórico antigo (Movimentações)', 'arquivarMovimentacoes')
    .addItem('Abrir o sistema (página)', 'abrirSistema')
    .addItem('Configurar sistema (rodar uma vez)', 'configurarSistema')
    .addToUi();
}

function atualizarAndamento_base_() {
  const ss = planilha_();
  const aba = ss.getSheetByName(ABA_PROT);
  const novos = buscarNovosComprovantes_(aba);
  const ultima = aba.getLastRow();
  if (ultima < 2) return;

  const n = ultima - 1;
  const formulasK = aba.getRange(2, 11, n, 1).getFormulas();
  const valoresK = aba.getRange(2, 11, n, 1).getValues();
  const agora = new Date();
  const todasMov = [];
  const mudancas = [];
  const erros = [];
  const respostas = [];

  for (let i = 0; i < n; i++) {
    const linha = i + 2;
    const fonte = String(formulasK[i][0] || valoresK[i][0] || '');
    const m = fonte.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
    if (!m) continue;
    const uuid = m[0];

    let p;
    try {
      const r = UrlFetchApp.fetch(API + uuid, { muteHttpExceptions: true, headers: { Accept: 'application/json' } });
      if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
      p = JSON.parse(r.getContentText());
    } catch (e) {
      erros.push('Linha ' + linha + ': ' + e.message);
      continue;
    }

    const movs = movimentos_(p);
    const req = movs.length ? movs[0].de : '';
    const ult = movs[movs.length - 1];
    const setor = setorAtual_(p, ult);
    const situacao = situacao_(p, movs, ult);

    const sitAnterior = String(aba.getRange(linha, 10).getValue());
    if (/^(Atendid|Respondido)/.test(situacao) && !/^(Atendid|Respondido)/.test(sitAnterior)) {
      respostas.push({ protocolo: p.numberYear, requerente: titulo_(req), situacao: situacao,
                       resumo: ult ? ult.resumo : '', link: PORTAL + uuid,
                       comprovante: (String(aba.getRange(linha, 12).getFormula()).match(/https?:\/\/[^"]+/) || [''])[0] });
    }

    const anterior = aba.getRange(linha, 7).getValue();
    const dataUlt = ult ? ult.data : null;
    if (anterior && dataUlt && new Date(anterior).getTime() !== dataUlt.getTime()) {
      mudancas.push(p.numberYear + ' (' + req + '): ' + (ult ? ult.resumo : ''));
    }

    aba.getRange(linha, 1, 1, 8).setValues([[
      p.numberYear,
      titulo_(req),
      limpa_(p.subject && p.subject.name),
      limpa_(p.title),
      new Date(p.createdAt),
      setor,
      dataUlt || '',
      ult ? ult.resumo : ''
    ]]);
    aba.getRange(linha, 1).setNumberFormat('@');
    if (!aba.getRange(linha, 9).getFormula()) {
      aba.getRange(linha, 9).setFormula('=IF(OR(G' + linha + '="";LEFT(J' + linha + ';7)="Atendid");"";TODAY()-INT(G' + linha + '))');
    }
    aba.getRange(linha, 10).setValue(situacao);
    if (!formulasK[i][0]) {
      aba.getRange(linha, 11).setFormula('=HYPERLINK("' + PORTAL + uuid + '";"Abrir processo")');
    }
    aba.getRange(linha, 13).setValue(agora);

    movs.forEach(mv => todasMov.push([p.numberYear, titulo_(req), mv.data, mv.de, mv.resumo]));
  }

  gravaMovimentacoes_(ss, todasMov.concat(movsFinalizados_(ss, todasMov)), agora);
  try { baixarAnexos_(ss, 90000); } catch (e) { Logger.log('Anexos: ' + e); }
  moverFinalizados_(aba, false);
  try { organizarPlanilha_(ss); } catch (e) { Logger.log('Organizar: ' + e); }   // atendidos/respondidos vão para a aba Finalizados
  const lembretes = criarLembretes_(respostas, ss);

  const msg = (novos.adicionados.length ? 'Novos protocolos encontrados na pasta: ' + novos.adicionados.join(', ') + '\n\n' : '') +
              (novos.falhas.length ? 'Comprovantes sem QR legível (cole o link na coluna K): ' + novos.falhas.join(', ') + '\n\n' : '') +
              (lembretes ? 'Lembretes criados na Agenda: ' + lembretes + '\n\n' : '') +
              (mudancas.length ? 'Movimentações novas:\n' + mudancas.join('\n') : 'Nenhuma movimentação nova.') +
              (erros.length ? '\n\nNão consegui consultar:\n' + erros.join('\n') : '');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Andamento atualizado', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
  PropertiesService.getScriptProperties().setProperty('ultimaAtualizacaoOk', String(Date.now()));
  limparCache_();
  return msg;
}

/** Menu: só procura comprovantes novos na pasta e acrescenta as linhas. */
function verificarNovosComprovantes() {
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_PROT);
  const r = buscarNovosComprovantes_(aba);
  const msg = (r.adicionados.length ? 'Novos protocolos adicionados: ' + r.adicionados.join(', ') : 'Nenhum comprovante novo na pasta.') +
              (r.falhas.length ? '\n\nNão consegui ler o QR de: ' + r.falhas.join(', ') + '\nCole o link do processo na coluna K dessas linhas.' : '') +
              (r.adicionados.length ? '\n\nClique em "Atualizar andamento agora" para preencher o andamento.' : '');
  SpreadsheetApp.getUi().alert('Comprovantes', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Procura na pasta do Drive arquivos (PDF ou imagem) que ainda não estão na planilha,
 * lê o QR code de cada um e acrescenta uma linha com o link do processo (K) e do comprovante (L).
 */
function buscarNovosComprovantes_(aba) {
  const res = { adicionados: [], falhas: [] };
  const ultima = aba.getLastRow();
  const conhecidos = {};
  const uuids = {};
  if (ultima >= 2) {
    aba.getRange(2, 11, ultima - 1, 2).getFormulas().forEach((l, i) => {
      const id = (String(l[1]).match(/\/d\/([A-Za-z0-9_-]{20,})/) || [])[1];
      if (id) conhecidos[id] = true;
      const u = (String(l[0]).match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i) || [])[0];
      if (u) uuids[u.toLowerCase()] = i + 2;
    });
  }
  const fin = conhecidosFinalizados_(aba.getParent());   // comprovantes que já estão na aba Finalizados
  Object.assign(conhecidos, fin.arquivos);
  const it = DriveApp.getFolderById(PASTA_COMPROVANTES).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const tipo = f.getMimeType();
    if (conhecidos[f.getId()] || !(tipo === 'application/pdf' || /^image\/(png|jpeg)$/.test(tipo))) continue;
    let url = '';
    try {
      url = qrDoArquivo_(f);
    } catch (e) {
      if (e.message === 'SEM_MINIATURA') continue;          // o Drive ainda está gerando a prévia; tenta na próxima
      Logger.log(f.getName() + ': ' + e);
    }
    const u = (url.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i) || [])[0];
    if (u && fin.uuids[u.toLowerCase()]) { conhecidos[f.getId()] = true; continue; }   // já está em Finalizados
    const linkArq = '=HYPERLINK("https://drive.google.com/file/d/' + f.getId() + '/view";"Abrir comprovante")';
    if (u && uuids[u.toLowerCase()]) {                      // protocolo já está na planilha: só liga o comprovante
      const lin = uuids[u.toLowerCase()];
      if (!aba.getRange(lin, 12).getFormula()) aba.getRange(lin, 12).setFormula(linkArq);
      continue;
    }
    const lin = aba.getLastRow() + 1;
    if (u) {
      aba.getRange(lin, 11).setFormula('=HYPERLINK("' + PORTAL + u + '";"Abrir processo")');
      aba.getRange(lin, 12).setFormula(linkArq);
      uuids[u.toLowerCase()] = lin;
      res.adicionados.push(f.getName().replace(/\.(pdf|png|jpe?g)$/i, '').replace(/^0*\s*COMPROVANTE DE PROTOCOLO\s*-\s*/i, ''));
    } else {
      aba.getRange(lin, 2).setValue(f.getName());
      aba.getRange(lin, 10).setValue('Não consegui ler o QR: cole o link do processo na coluna K');
      aba.getRange(lin, 12).setFormula(linkArq);
      res.falhas.push(f.getName());
    }
    conhecidos[f.getId()] = true;
  }
  return res;
}

/** Devolve o texto do QR code do arquivo (usa a prévia do Drive para PDFs). */
function qrDoArquivo_(f) {
  if (/^image\//.test(f.getMimeType())) return lerQR_(f.getBlob().getBytes());
  const token = ScriptApp.getOAuthToken();
  const meta = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '?fields=thumbnailLink',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }).getContentText());
  if (!meta.thumbnailLink) throw new Error('SEM_MINIATURA');
  for (const tam of [1600, 2400]) {
    const link = meta.thumbnailLink.replace(/=s\d+$/, '') + '=s' + tam;
    const r = UrlFetchApp.fetch(link, { headers: { Authorization: 'Bearer ' + token, Accept: 'image/png,image/jpeg' }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) continue;
    try {
      const txt = lerQR_(r.getBlob().getBytes());
      if (txt) return txt;
    } catch (e) { Logger.log('Leitura ' + tam + ': ' + e); }
  }
  return '';
}

/** Movimentos válidos, sem duplicados, em ordem cronológica. */
function movimentos_(p) {
  const vistos = {};
  const lista = [];
  (p.movements || []).forEach((m, i) => {
    if (m.disabled) return;
    const dest = (m.recipients || []).map(r => {
      const c = r.processCommunicant || {};
      return limpa_((c.sector && c.sector.name) || (c.user && c.user.name) || '');
    }).filter(String);
    const corpo = texto_(m.body);
    if (!dest.length && !corpo) return;
    const de = limpa_(m.sender && m.sender.user && m.sender.user.name);
    const chave = String(m.createdAt).slice(0, 16) + '|' + de + '|' + corpo;
    const docs = (m.documents || []).filter(d => !d.disabled).length;
    if (vistos[chave]) {                       // mesmo movimento enviado a outro destino
      const v = vistos[chave];
      dest.forEach(d => { if (v.dest.indexOf(d) < 0) v.dest.push(d); });
      v.docs = Math.max(v.docs, docs);
      v.resumo = resumo_(v);
      return;
    }
    const mv = { data: new Date(m.createdAt), de: de, dest: dest, corpo: corpo, docs: docs,
                 abertura: lista.length === 0, deUsuario: (m.recipients || []).some(r => r.processCommunicant && r.processCommunicant.user) };
    mv.resumo = resumo_(mv);
    vistos[chave] = mv;
    lista.push(mv);
  });
  return lista.sort((a, b) => a.data - b.data);
}

function resumo_(mv) {
  let s = mv.abertura
    ? 'Protocolo aberto e encaminhado para ' + mv.dest.join(', ')
    : titulo_(mv.de) + ' enviou para ' + mv.dest.join(', ');
  if (!mv.abertura && mv.corpo) s += ': ' + mv.corpo.slice(0, 300);
  if (!mv.abertura && mv.docs) s += ' (' + mv.docs + ' documento(s) anexado(s))';
  return s;
}

function setorAtual_(p, ult) {
  let aFazer = (p.sectorsToDo || []).map(s => limpa_(s.sector && s.sector.name)).filter(String);
  const longos = aFazer.filter(s => !/^[A-Z\/]{2,10}$/.test(s));   // descarta siglas soltas (ex.: DVGRH)
  if (longos.length) aFazer = longos;
  if (aFazer.length) return Array.from(new Set(aFazer)).join(', ');
  if (ult && ult.deUsuario) return 'Com o requerente (resposta enviada)';
  const fazendo = (p.sectorsDoing || []).map(s => limpa_(s.sector && s.sector.name)).filter(s => s && s !== 'Externo');
  return fazendo.length ? Array.from(new Set(fazendo)).join(', ') : 'Sem setor pendente';
}

function situacao_(p, movs, ult) {
  const aFazer = (p.sectorsToDo || []).length;
  if (!aFazer && ult && ult.deUsuario) return ult.docs ? 'Atendido: documento enviado ao requerente' : 'Respondido ao requerente';
  if (movs.length <= 1) return 'Aguardando análise';
  return 'Em andamento';
}

function gravaMovimentacoes_(ss, linhas, agora) {
  const aba = ss.getSheetByName(ABA_MOV);
  const antigas = {};
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 6).getValues().forEach(l => {
      const k = l[0] + '|' + (l[2] instanceof Date ? l[2].getTime() : l[2]) + '|' + l[3];
      if (l[5]) antigas[k] = l[5];
    });
    aba.getRange(2, 1, aba.getLastRow() - 1, 6).clearContent();
  }
  if (!linhas.length) return;
  linhas.sort((a, b) => String(a[0]).localeCompare(String(b[0])) || a[2] - b[2]);
  const saida = linhas.map(l => {
    const k = l[0] + '|' + l[2].getTime() + '|' + l[3];
    return [l[0], l[1], l[2], l[3], l[4], antigas[k] || agora];
  });
  aba.getRange(2, 1, saida.length, 6).setValues(saida);
  aba.getRange(2, 1, saida.length, 1).setNumberFormat('@');
}

/** Cria um evento de 30 minutos na agenda principal para cada protocolo que acabou de ser respondido. */
function criarLembretes_(respostas, ss) {
  if (!respostas.length) return 0;
  const urlPlanilha = ss.getUrl();
  const cal = CalendarApp.getDefaultCalendar();
  const inicio = new Date();
  inicio.setMinutes(Math.ceil((inicio.getMinutes() + 5) / 15) * 15, 0, 0);
  respostas.forEach((r, i) => {
    const ini = new Date(inicio.getTime() + i * 5 * 60000);
    const fim = new Date(ini.getTime() + 30 * 60000);
    const verde = /^Atendid/.test(r.situacao);
    const titulo = (verde ? 'Protocolo administrativo atendido: ' : 'Protocolo administrativo respondido (verificar): ') + r.protocolo + ' - ' + r.requerente;
    const desc = r.situacao + '\n\n' + r.resumo + '\n\nProcesso no portal: ' + r.link + '\nPlanilha: ' + urlPlanilha;
    const ev = cal.createEvent(titulo, ini, fim, { description: desc });
    ev.removeAllReminders();
    ev.addPopupReminder(0);
    ev.addEmailReminder(0);
    ev.setColor(verde ? CalendarApp.EventColor.GREEN : CalendarApp.EventColor.YELLOW);
    anexar_(ev, ss, r);
  });
  return respostas.length;
}

/** Anexa a planilha (e o comprovante, se houver) ao evento. Usa o serviço avançado "Calendar". */
function anexar_(ev, ss, r) {
  try {
    const anexos = [{ fileUrl: ss.getUrl(), title: ss.getName(), mimeType: 'application/vnd.google-apps.spreadsheet' }];
    if (r.comprovante) anexos.push({ fileUrl: r.comprovante, title: 'Comprovante ' + r.protocolo + ' - ' + r.requerente, mimeType: 'application/pdf' });
    const id = ev.getId().replace(/@google\.com$/, '');
    Calendar.Events.patch({ attachments: anexos }, 'primary', id, { supportsAttachments: true });
  } catch (e) {
    Logger.log('Não foi possível anexar ao evento: ' + e);
  }
}

function texto_(html) {
  return String(html || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
}
function limpa_(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function titulo_(s) {
  s = limpa_(s);
  if (s !== s.toUpperCase()) return s;
  return s.toLowerCase().replace(/(^|\s)(\S)/g, (a, b, c) => b + c.toUpperCase())
          .replace(/\s(Da|De|Do|Das|Dos|E)\s/g, m => m.toLowerCase());
}

function ativarDiario() {
  desativarDiario();
  [7, 13].forEach(h => ScriptApp.newTrigger('atualizarAndamento').timeBased().everyDays(1).atHour(h).create());
  SpreadsheetApp.getUi().alert('Atualização automática ativada: duas vezes por dia, por volta das 7h e das 13h.');
}
function desativarDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizarAndamento')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** Lê o QR code de uma imagem (PNG ou JPEG) e devolve o texto, ou '' se não achar. */
function lerQR_(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : Uint8Array.from(bytes, b => b & 255);
  let img;
  if (u8[0] === 0x89 && u8[1] === 0x50) img = decodePNG_(u8);
  else if (u8[0] === 0xFF && u8[1] === 0xD8) img = QRLIB.jpegDecode(u8, { useTArray: true, formatAsRGBA: true, maxMemoryUsageInMB: 1024 });
  else throw new Error('formato de imagem não suportado');
  let r = QRLIB.jsQR(img.data, img.width, img.height, { inversionAttempts: 'dontInvert' });
  if (!r) {                                   // tenta de novo só no quarto superior direito, ampliado
    const rec = recorte_(img, 0.5, 0, 1, 0.4);
    r = QRLIB.jsQR(rec.data, rec.width, rec.height, { inversionAttempts: 'attemptBoth' });
  }
  return r ? r.data : '';
}

function recorte_(img, x0, y0, x1, y1) {
  const sx = Math.floor(img.width * x0), sy = Math.floor(img.height * y0);
  const w = Math.floor(img.width * (x1 - x0)), h = Math.floor(img.height * (y1 - y0));
  const out = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    const src = ((sy + y) * img.width + sx) * 4;
    out.set(img.data.subarray(src, src + w * 4), y * w * 4);
  }
  return { data: out, width: w, height: h };
}

function decodePNG_(u8) {
  let p = 8, width = 0, height = 0, depth = 0, ctype = 0, palette = null;
  const idat = [];
  let total = 0;
  while (p < u8.length) {
    const len = (u8[p] << 24 | u8[p + 1] << 16 | u8[p + 2] << 8 | u8[p + 3]) >>> 0;
    const type = String.fromCharCode(u8[p + 4], u8[p + 5], u8[p + 6], u8[p + 7]);
    const d = u8.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') {
      width = (d[0] << 24 | d[1] << 16 | d[2] << 8 | d[3]) >>> 0;
      height = (d[4] << 24 | d[5] << 16 | d[6] << 8 | d[7]) >>> 0;
      depth = d[8]; ctype = d[9];
      if (d[12]) throw new Error('PNG entrelaçado não suportado');
    } else if (type === 'PLTE') palette = d;
    else if (type === 'IDAT') { idat.push(d); total += d.length; }
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (depth !== 8) throw new Error('PNG com profundidade ' + depth + ' não suportado');
  const z = new Uint8Array(total);
  let o = 0; idat.forEach(c => { z.set(c, o); o += c.length; });
  const raw = QRLIB.inflate(z);
  const ch = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ctype];
  const stride = width * ch;
  const px = new Uint8Array(height * stride);
  for (let y = 0; y < height; y++) {
    const f = raw[y * (stride + 1)];
    const lin = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    const cur = px.subarray(y * stride, (y + 1) * stride);
    const prev = y ? px.subarray((y - 1) * stride, y * stride) : null;
    for (let x = 0; x < stride; x++) {
      const a = x >= ch ? cur[x - ch] : 0, b = prev ? prev[x] : 0, c = (prev && x >= ch) ? prev[x - ch] : 0;
      let v = lin[x];
      if (f === 1) v += a;
      else if (f === 2) v += b;
      else if (f === 3) v += (a + b) >> 1;
      else if (f === 4) { const pp = a + b - c, pa = Math.abs(pp - a), pb = Math.abs(pp - b), pc = Math.abs(pp - c); v += (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c); }
      cur[x] = v & 255;
    }
  }
  const rgba = new Uint8ClampedArray(width * height * 4);
  for (let i = 0, j = 0; i < width * height; i++, j += ch) {
    let r, g, b2;
    if (ctype === 3) { const k = px[j] * 3; r = palette[k]; g = palette[k + 1]; b2 = palette[k + 2]; }
    else if (ch <= 2) { r = g = b2 = px[j]; }
    else { r = px[j]; g = px[j + 1]; b2 = px[j + 2]; }
    rgba[i * 4] = r; rgba[i * 4 + 1] = g; rgba[i * 4 + 2] = b2; rgba[i * 4 + 3] = 255;
  }
  return { data: rgba, width: width, height: height };
}


/* =====================================================================
 * SISTEMA (acrescentado em 23/11/2026)
 * - Aba "Finalizados": protocolos atendidos/respondidos saem da aba
 *   Protocolos e vão para lá sozinhos, com a data em que ficaram prontos.
 * - Aviso ao abrir a planilha: quem ficou pronto desde o seu último acesso.
 * - Página própria (App da Web) com busca, filtros e histórico.
 * Para ligar tudo, rode uma vez: menu Protocolos > Configurar sistema.
 * ===================================================================== */
const ABA_FIN = 'Finalizados';

/** Cria a aba Finalizados, move os já prontos, ajusta o Painel e liga o aviso ao abrir. */
function configurarSistema() {
  const ss = SpreadsheetApp.getActive();
  const aba = ss.getSheetByName(ABA_PROT);
  const fin = abaFinalizados_(ss);
  const movidos = moverFinalizados_(aba, true);

  // aviso ao abrir (gatilho instalável: o onOpen simples não pode abrir janelas)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'aoAbrir')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('aoAbrir').forSpreadsheet(ss).onOpen().create();

  // Painel: "Atendidos" passa a contar a aba Finalizados
  const painel = ss.getSheetByName('Painel');
  if (painel) {
    painel.getRange('E5').setFormula('=COUNTIF(Finalizados!J2:J;"Atendido*")');
    painel.getRange('G2').setValue('As abas: Painel (resumo e gráficos), Protocolos (pedidos em andamento), Finalizados (atendidos ou respondidos, com a data em que ficaram prontos) e Movimentações (histórico completo).');
    painel.getRange('G8').setValue('5. Os protocolos atendidos ou respondidos saem da aba Protocolos e vão sozinhos para a aba Finalizados, e o PDF do comprovante vai para a pasta "Processos encerrados".');
    painel.getRange('G14').setValue('Prontos desde o último acesso').setFontWeight('bold');
  }
  PropertiesService.getUserProperties().setProperty('ultimoAcesso', String(Date.now()));
  SpreadsheetApp.getUi().alert('Sistema configurado',
    'Aba "Finalizados" pronta (' + movidos + ' protocolo(s) movido(s)).\n' +
    'A partir de agora, ao abrir a planilha aparece quem ficou pronto desde o seu último acesso.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function abaFinalizados_(ss) {
  let fin = ss.getSheetByName(ABA_FIN);
  if (fin) return fin;
  const prot = ss.getSheetByName(ABA_PROT);
  fin = ss.insertSheet(ABA_FIN, prot.getIndex());
  prot.getRange(1, 1, 1, 13).copyTo(fin.getRange(1, 1, 1, 13));
  prot.getRange(1, 13).copyTo(fin.getRange(1, 14), { formatOnly: true });
  fin.getRange(1, 14).setValue('Finalizado em');
  fin.setFrozenRows(1);
  fin.setFrozenColumns(2);
  for (let c = 1; c <= 13; c++) fin.setColumnWidth(c, prot.getColumnWidth(c));
  fin.setColumnWidth(14, 130);
  return fin;
}

/**
 * Move para a aba Finalizados as linhas com situação "Atendido..." ou "Respondido...".
 * Na primeira configuração (inicial = true) usa a data da última movimentação como
 * "Finalizado em"; depois, usa a hora em que a atualização percebeu a resposta.
 */
function moverFinalizados_(aba, inicial) {
  const ss = aba.getParent();
  const fin = abaFinalizados_(ss);
  const ultima = aba.getLastRow();
  if (ultima < 2) return 0;
  const sit = aba.getRange(2, 10, ultima - 1, 1).getValues();
  const linhas = [];
  sit.forEach((l, i) => { if (/^(Atendid|Respondido)/.test(String(l[0]))) linhas.push(i + 2); });
  const agora = new Date();
  linhas.slice().reverse().forEach(linha => {      // de baixo para cima, para não bagunçar os números
    const destino = fin.getLastRow() + 1;
    aba.getRange(linha, 1, 1, 13).copyTo(fin.getRange(destino, 1, 1, 13));
    const quando = inicial ? (aba.getRange(linha, 7).getValue() || agora) : agora;
    fin.getRange(destino, 14).setValue(quando).setNumberFormat('dd/MM/yyyy HH:mm');
    arquivarComprovante_(String(aba.getRange(linha, 12).getFormula()));
    aba.deleteRow(linha);
  });
  if (linhas.length && fin.getLastRow() > 2) {
    fin.getRange(2, 1, fin.getLastRow() - 1, 14).sort({ column: 14, ascending: false });
  }
  return linhas.length;
}

const PASTA_ENCERRADOS = 'ID_PASTA_ENCERRADOS';   // "Processos encerrados" (dentro da pasta de comprovantes)

function pastaEncerrados_() { return DriveApp.getFolderById(PASTA_ENCERRADOS); }

/** Move o PDF do comprovante para "Processos encerrados" (o link na planilha continua valendo). */
function arquivarComprovante_(formula) {
  const id = (formula.match(/\/d\/([A-Za-z0-9_-]{20,})/) || [])[1];
  if (!id) return;
  try {
    const f = DriveApp.getFileById(id);
    const destino = pastaEncerrados_();
    if (!f.getParents().hasNext() || f.getParents().next().getId() !== destino.getId()) f.moveTo(destino);
  } catch (e) {
    Logger.log('Não consegui mover o comprovante ' + id + ': ' + e);
  }
}

/** Comprovantes e processos que já estão na aba Finalizados (para não voltarem como "novos"). */
function conhecidosFinalizados_(ss) {
  const res = { arquivos: {}, uuids: {} };
  const fin = ss.getSheetByName(ABA_FIN);
  if (!fin || fin.getLastRow() < 2) return res;
  fin.getRange(2, 11, fin.getLastRow() - 1, 2).getFormulas().forEach(l => {
    const id = (String(l[1]).match(/\/d\/([A-Za-z0-9_-]{20,})/) || [])[1];
    if (id) res.arquivos[id] = true;
    const u = (String(l[0]).match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i) || [])[0];
    if (u) res.uuids[u.toLowerCase()] = true;
  });
  return res;
}

/** Mantém na aba Movimentações o histórico dos protocolos que já foram para Finalizados. */
function movsFinalizados_(ss, todasMov) {
  const mov = ss.getSheetByName(ABA_MOV);
  if (!mov || mov.getLastRow() < 2) return [];
  const atuais = {};
  todasMov.forEach(l => { atuais[l[0]] = true; });
  return mov.getRange(2, 1, mov.getLastRow() - 1, 5).getValues()
    .filter(l => l[0] && !atuais[l[0]])
    .map(l => [String(l[0]), l[1], paraData_(l[2]), l[3], l[4]]);
}

function paraData_(v) {
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : new Date(0);
}

/** Lista dos que ficaram prontos depois de "desde" (milissegundos). */
function prontosDesde_(desde) {
  const fin = planilha_().getSheetByName(ABA_FIN);
  if (!fin || fin.getLastRow() < 2) return [];
  const n = fin.getLastRow() - 1;
  const v = fin.getRange(2, 1, n, 14).getValues();
  const f = fin.getRange(2, 11, n, 1).getFormulas();
  const out = [];
  v.forEach((l, i) => {
    const quando = paraData_(l[13]);
    if (quando.getTime() > desde) out.push({
      protocolo: String(l[0]), requerente: String(l[1]), situacao: String(l[9]),
      quando: Utilities.formatDate(quando, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'),
      link: (String(f[i][0]).match(/HYPERLINK\("([^"]+)"/) || [])[1] || ''
    });
  });
  return out;
}

/** Gatilho instalável "ao abrir": mostra quem ficou pronto desde o último acesso. */
function aoAbrir() {
  const props = PropertiesService.getUserProperties();
  const desde = Number(props.getProperty('ultimoAcesso') || 0);
  const lista = prontosDesde_(desde);
  props.setProperty('ultimoAcesso', String(Date.now()));

  const painel = planilha_().getSheetByName('Painel');
  if (painel) {
    const quando = desde ? Utilities.formatDate(new Date(desde), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : '';
    painel.getRange('G14').setValue('Prontos desde o último acesso' + (quando ? ' (' + quando + ')' : '')).setFontWeight('bold');
    painel.getRange('G15:G34').clearContent();
    const linhas = lista.length
      ? lista.slice(0, 20).map(p => ['• ' + p.requerente + ' (' + p.protocolo + ') - ' + p.situacao + ' - ' + p.quando])
      : [['Nenhum protocolo ficou pronto desde o último acesso.']];
    painel.getRange(15, 7, linhas.length, 1).setValues(linhas);
  }
  if (!lista.length) return;
  mostrarProntos_(lista, 'Prontos desde o seu último acesso');
}

/** Menu: mostra de novo os prontos dos últimos 7 dias. */
function verProntosRecentes() {
  const lista = prontosDesde_(Date.now() - 7 * 864e5);
  if (!lista.length) {
    SpreadsheetApp.getUi().alert('Nenhum protocolo ficou pronto nos últimos 7 dias.');
    return;
  }
  mostrarProntos_(lista, 'Prontos nos últimos 7 dias');
}

function mostrarProntos_(lista, titulo) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const itens = lista.map(p =>
    '<li><b>' + esc(p.requerente) + '</b> (' + esc(p.protocolo) + ')<br>' +
    '<span class="' + (/^Atendid/.test(p.situacao) ? 'ok' : 'ver') + '">' + esc(p.situacao) + '</span>' +
    ' · ' + esc(p.quando) +
    (p.link ? ' · <a href="' + esc(p.link) + '" target="_blank">Abrir processo</a>' : '') + '</li>').join('');
  const html = '<style>body{font-family:Arial,sans-serif;font-size:14px}li{margin:0 0 10px}' +
    '.ok{color:#137333;font-weight:bold}.ver{color:#b06000;font-weight:bold}</style>' +
    '<p>' + lista.length + ' protocolo(s):</p><ul>' + itens + '</ul>' +
    '<p style="color:#666">Eles estão na aba Finalizados.</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(Math.min(120 + lista.length * 60, 520)), titulo);
}

/** Menu: link para a página do sistema. */
function abrirSistema() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? '<p style="font-family:Arial"><a href="' + url + '" target="_blank">Abrir o sistema de protocolos</a></p>'
    : '<p style="font-family:Arial">A página ainda não foi publicada. No editor do Apps Script: Implantar &gt; Nova implantação &gt; App da Web.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(360).setHeight(90), 'Sistema de protocolos');
}

/* ---------------------------- PÁGINA (App da Web) ---------------------------- */

function doGet() {
  return HtmlService.createHtmlOutput(PAGINA_HTML_)
    .setTitle('Protocolos - Prefeitura de Foz')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Dados para a página. */
function dadosSistema_base_() {
  const props = PropertiesService.getUserProperties();
  const dono = ehDono_();
  const desde = Number(props.getProperty('ultimoAcesso') || 0);
  if (dono) props.setProperty('ultimoAcesso', String(Date.now()));
  const d = lerCache_('dadosPagina') || salvarCache_('dadosPagina', dadosPlanilha_());
  d.finalizados.forEach(p => { p.novo = p.tFin > desde; });
  d.dono = dono;
  d.ultimoAcesso = desde ? Utilities.formatDate(new Date(desde), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') : '';
  return d;
}

/** Lê a planilha para a página (o resultado fica 10 min no cache). */
function dadosPlanilha_() {
  const ss = planilha_();

  const ler = (nome, fim) => {
    const sh = ss.getSheetByName(nome);
    if (!sh || sh.getLastRow() < 2) return [];
    const n = sh.getLastRow() - 1;
    const v = sh.getRange(2, 1, n, 14).getDisplayValues();
    const f = sh.getRange(2, 11, n, 2).getFormulas();
    const brutos = fim ? sh.getRange(2, 14, n, 1).getValues() : null;
    return v.filter(l => l[0]).map((l, i) => ({
      protocolo: l[0], requerente: l[1], tipo: l[2], objeto: l[3], aberto: l[4], setor: l[5],
      dataUlt: l[6], ultima: l[7], dias: l[8], situacao: l[9], consulta: l[12],
      finalizadoEm: fim ? l[13] : '',
      tFin: fim ? paraData_(brutos[i][0]).getTime() : 0,
      processo: (String(f[i][0]).match(/HYPERLINK\("([^"]+)"/) || [])[1] || '',
      comprovante: (String(f[i][1]).match(/HYPERLINK\("([^"]+)"/) || [])[1] || '',
      finalizado: fim
    }));
  };

  const historico = {};
  const mov = ss.getSheetByName(ABA_MOV);
  if (mov && mov.getLastRow() > 1) {
    mov.getRange(2, 1, mov.getLastRow() - 1, 5).getDisplayValues().forEach(l => {
      (historico[l[0]] = historico[l[0]] || []).push({ data: l[2], quem: l[3], texto: l[4] });
    });
  }
  return {
    andamento: comPasta_(ler(ABA_PROT, false)),
    finalizados: comPasta_(ler(ABA_FIN, true)),
    historico: historico,
    planilha: ss.getUrl()
  };
}

/** Botão "Atualizar agora" da página. */
function atualizarPelaPagina_base_() {
  if (!ehDono_()) throw new Error('Somente visualização: a atualização é feita pela conta do escritório.');
  return atualizarAndamento() || 'Atualizado.';
}

/** A planilha, também quando o código roda pela página (onde não há planilha ativa). */
function planilha_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById('ID_DA_PLANILHA');
}

const PAGINA_HTML_ = String.raw`<!doctype html>
<html><head><meta charset="utf-8">
<style>
  :root{--side:#1e3a8a;--side2:#2a4a9e;--blue:#2563eb;--bg:#f4f6fb;--card:#fff;--tx:#1e293b;--mut:#64748b;--bd:#e2e8f0;--ok:#15803d;--okbg:#ecfdf3;--ver:#c2410c;--verbg:#fff7ed;--red:#dc2626;--redbg:#fef2f2}
  *{box-sizing:border-box}
  body{margin:0;font-family:Inter,Roboto,Arial,sans-serif;background:var(--bg);color:var(--tx);display:flex;min-height:100vh}
  aside{width:215px;background:var(--side);color:#fff;flex-shrink:0;display:flex;flex-direction:column}
  .marca{padding:22px 18px 18px}
  .marca b{font-size:19px;letter-spacing:.5px;display:block}
  .marca span{font-size:12px;opacity:.85}
  nav a{display:flex;align-items:center;justify-content:space-between;padding:11px 18px;color:#fff;text-decoration:none;font-size:14px;cursor:pointer}
  nav a:hover{background:var(--side2)}
  nav a.on{background:var(--side2);font-weight:600}
  nav .cont{background:#fff;color:var(--side);border-radius:999px;font-size:11px;padding:1px 8px;font-weight:700}
  nav .cont.verm{background:#ef4444;color:#fff}
  .rodape{margin-top:auto;padding:16px 18px;font-size:12px;opacity:.85;line-height:1.5}
  .rodape a{color:#fff}
  .conteudo{flex:1;display:flex;gap:18px;padding:22px 26px;min-width:0}
  main{flex:1;min-width:0}
  .lateral{width:300px;flex-shrink:0}
  .titulo{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:16px}
  .titulo h1{margin:0;font-size:24px;color:var(--side)}
  .btn{background:var(--blue);color:#fff;border:0;border-radius:6px;padding:10px 16px;font-size:14px;cursor:pointer;font-weight:500}
  .btn:disabled{opacity:.6;cursor:wait}
  .kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px}
  .kpi{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:16px 18px}
  .kpi b{display:block;font-size:30px;font-weight:600;color:var(--side)}
  .kpi span{color:var(--mut);font-size:13px}
  .kpi.verm b{color:var(--red)} .kpi.lar b{color:var(--ver)} .kpi.verde b{color:var(--ok)}
  .painel{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:16px 18px;margin-bottom:16px}
  .painel h2{margin:0 0 12px;font-size:16px;color:var(--side)}
  .filtros{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
  .filtros input,.filtros select{padding:9px 10px;border:1px solid var(--bd);border-radius:6px;font-size:14px;background:#fff;color:var(--tx)}
  .filtros input{flex:1;min-width:200px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{text-align:left;background:#f1f5fb;color:var(--tx);font-weight:600;padding:10px;border-bottom:1px solid var(--bd)}
  td{padding:10px;border-bottom:1px solid var(--bd);vertical-align:top}
  tr.novo td{background:var(--okbg)}
  th.ordenavel{cursor:pointer;user-select:none;color:var(--blue)}
  tr.det td{background:#fafbfe;color:var(--mut)}
  .nome{font-weight:600;color:var(--tx)}
  .sub{color:var(--mut);font-size:12px}
  a.lkn{color:inherit;text-decoration:none}a.lkn:hover{color:var(--blue);text-decoration:underline}
  a.lks{color:var(--blue);text-decoration:none}a.lks:hover{text-decoration:underline}
  .lapis{cursor:pointer;margin-right:6px;font-size:12px;opacity:.65}.lapis:hover{opacity:1}
  .npasta{display:inline-block;background:#eef2ff;color:var(--side);border-radius:4px;font-size:11px;font-weight:700;padding:1px 6px;margin-right:6px}
  .mbox .btn2{background:#e2e8f0;color:var(--tx);margin-top:8px}
  #erPasta,#erEmail{color:var(--red)}
  #modalEmail,#modalPasta{position:fixed;inset:0;background:rgba(15,23,42,.55);display:flex;align-items:center;justify-content:center;z-index:50}
  #modalEmail{backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);background:rgba(15,23,42,.35)}
  .mbox{background:#fff;border-radius:10px;padding:24px;width:340px;max-width:90vw;box-shadow:0 10px 30px rgba(0,0,0,.25)}
  .mbox h3{margin:0 0 6px;color:var(--tx)}.mbox p{margin:0;color:var(--mut);font-size:13px}
  .mbox input{width:100%;padding:9px;margin:12px 0 0;border:1px solid #cbd5e1;border-radius:6px;box-sizing:border-box;font-size:16px}.mbox .btn{margin-top:12px;width:100%}
  .tag{display:inline-block;font-size:11px;border-radius:999px;padding:2px 8px;font-weight:600;white-space:nowrap}
  .t-ok{background:var(--okbg);color:var(--ok)} .t-ver{background:var(--verbg);color:var(--ver)}
  .t-and{background:#eef2ff;color:var(--side)} .t-red{background:var(--redbg);color:var(--red)}
  .dias{font-weight:700} .dias.alto{color:var(--red)}
  .t-anx{background:var(--verbg);color:var(--ver);cursor:help}
  a.lk,span.lk{color:var(--blue);cursor:pointer;text-decoration:underline;font-size:12px;margin-right:8px;white-space:nowrap}
  .hist div{margin-bottom:4px}
  .cal{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:16px}
  .cal h3{margin:0 0 4px;font-size:16px;color:var(--side)}
  .cal .data{color:var(--mut);font-size:13px;margin-bottom:12px}
  .cal h4{margin:14px 0 8px;font-size:13px;color:var(--blue)}
  .cal h4.verm{color:var(--red)}
  .item{border-left:4px solid var(--blue);background:#eef4ff;border-radius:6px;padding:9px 11px;margin-bottom:8px;font-size:13px}
  .item.ok{border-color:var(--ok);background:var(--okbg)}
  .item.red{border-color:var(--red);background:var(--redbg)}
  .item b{display:block;margin-bottom:2px}
  .item span{color:var(--mut);font-size:12px}
  .vazio{color:var(--mut);padding:16px;text-align:center;font-size:13px}
  #msg{white-space:pre-wrap;background:#fff;border:1px solid var(--bd);border-left:4px solid var(--blue);border-radius:8px;padding:12px;margin-bottom:16px;display:none;font-size:13px}
  @media (max-width:1000px){.conteudo{flex-direction:column}.lateral{width:auto}}
  @media (max-width:700px){
    body{flex-direction:column}
    aside{width:auto;position:sticky;top:0;z-index:20}
    .marca{padding:10px 14px 6px}.marca b{font-size:16px;display:inline}.marca span{display:none}
    nav{display:flex;overflow-x:auto;white-space:nowrap;-webkit-overflow-scrolling:touch}
    nav a{padding:9px 12px;font-size:13px;gap:6px;flex-shrink:0}
    .rodape{display:none}
    .conteudo{padding:12px;gap:12px}
    .titulo h1{font-size:20px}.titulo .btn{padding:8px 12px;font-size:13px}
    .kpis{grid-template-columns:1fr 1fr;gap:8px}
    .kpi{padding:10px 12px}.kpi b{font-size:22px}.kpi span{font-size:12px}
    .painel{padding:12px}
    .filtros input,.filtros select{flex:1 1 100%;min-width:0;font-size:16px}
    table,tbody,tr,td{display:block;width:100%}
    thead{display:none}
    tbody tr{border:1px solid var(--bd);border-radius:8px;margin-bottom:10px;padding:6px 0;background:#fff}
    tbody tr.det{margin-top:-8px;border-top:0;border-radius:0 0 8px 8px}
    td{border:0;padding:4px 12px}
    td:nth-child(3):before{content:"Anexos: ";color:var(--mut);font-size:12px}
    td:nth-child(5):before{content:"Dias parado: ";color:var(--mut);font-size:12px}
    td:nth-child(6){padding-top:8px}
    a.lk,span.lk{display:inline-block;padding:4px 0;font-size:13px}
    .lateral{width:auto}
  }
</style></head>
<body>
<aside>
  <div class="marca"><b>SALGAR</b><span>Processos Administrativos · Prefeitura de Foz</span></div>
  <nav id="menu">
    <a data-v="andamento" class="on">Em andamento <span class="cont" id="cAnd">0</span></a>
    <a data-v="todos">Todos</a>
    <a data-v="anexo">Com doc. anexado <span class="cont" id="cAnx">0</span></a>
    <a data-v="parados">Parados 15+ dias <span class="cont verm" id="cPar">0</span></a>
    <a data-v="finalizados">Finalizados <span class="cont" id="cFin">0</span></a>
    <a id="lkPlanilha" target="_blank">Abrir planilha</a>
  </nav>
  <div class="rodape" id="rodape"></div>
</aside>
<div class="conteudo">
<main>
  <div class="titulo"><h1>Protocolos</h1><button class="btn" id="btAtualizar" onclick="atualizar()">↻ Atualizar agora</button></div>
  <div id="msg"></div>
  <div class="kpis" id="kpis"></div>
  <div class="painel">
    <h2 id="tituloLista">Em andamento</h2>
    <div class="filtros">
      <input id="busca" placeholder="Buscar por nome, protocolo, objeto ou setor">
      <select id="fSit"><option value="">Todas as situações</option></select>
      <select id="fOrd">
        <option value="dias">Mais tempo parado primeiro</option>
        <option value="recente">Movimentação mais recente</option>
        <option value="nome">Nome (A-Z)</option>
      <option value="nomeZ">Nome (Z-A)</option><option value="sit">Situação (A-Z)</option><option value="sitZ">Situação (Z-A)</option>
      </select>
    </div>
    <div id="lista"><div class="vazio">Carregando...</div></div>
  </div>
</main>
<div class="lateral">
  <div class="cal" id="lateral"><div class="vazio">Carregando...</div></div>
</div>
</div>
<script>
let D = null, aba = 'andamento';
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const NOMES = { andamento: 'Em andamento', parados: 'Parados há 15 dias ou mais', finalizados: 'Finalizados', todos: 'Todos os protocolos', anexo: 'Em andamento com documento anexado' };
let registrado = false;
function registrar(d) {
  if (registrado) return;
  registrado = true;
  let s = null;
  try { s = JSON.parse(localStorage.getItem('acessoNomeEmail') || 'null'); } catch (x) {}
  if (s && s.n && s.e) { google.script.run.registrarAcesso(s.e, s.n); return; }
  const f = document.createElement('div');
  f.id = 'modalEmail';
  f.innerHTML = '<div class="mbox"><h3>Identifique-se</h3><p>Informe seu nome e e-mail para acessar o painel de protocolos.</p>' +
    '<input id="inNome" type="text" maxlength="80" placeholder="Seu nome">' +
    '<input id="inEmail" type="email" maxlength="120" placeholder="seu@email.com">' +
    '<button class="btn" id="btEntrar">Entrar</button><div id="erEmail" class="sub"></div></div>';
  document.body.appendChild(f);
  $('btEntrar').addEventListener('click', salvarEmail);
  ['inNome', 'inEmail'].forEach(id => $(id).addEventListener('keydown', ev => { if (ev.key === 'Enter') salvarEmail(); }));
  setTimeout(() => $('inNome').focus(), 50);
}
function salvarEmail() {
  const n = $('inNome').value.trim(), e = $('inEmail').value.trim();
  if (n.length < 2) { $('erEmail').textContent = 'Digite seu nome.'; return; }
  if (e.indexOf('@') < 1 || e.lastIndexOf('.') < e.indexOf('@') + 2 || /\s/.test(e)) { $('erEmail').textContent = 'Digite um e-mail válido.'; return; }
  try { localStorage.setItem('acessoNomeEmail', JSON.stringify({ n: n, e: e })); } catch (x) {}
  google.script.run.registrarAcesso(e, n);
  $('modalEmail').remove();
}
function numPasta(p) { const m = String(p.pastaNome || '').match(/^(\d{1,4})\s/); return m ? m[1] : ''; }
document.addEventListener('click', ev => { const t = ev.target.closest ? ev.target.closest('.lapis') : null; if (t) editarPasta(t.getAttribute('data-prot')); });
function editarPasta(prot) {
  const p = D.andamento.filter(x => x.protocolo === prot)[0];
  if (!p) return;
  const f = document.createElement('div');
  f.id = 'modalPasta';
  f.innerHTML = '<div class="mbox"><h3>Pasta do cliente</h3><p><b>' + esc(p.requerente) + '</b> · ' + esc(p.protocolo) + '</p>' +
    '<p style="margin-top:8px">Atual: ' + (p.pasta ? '<a class="lks" href="' + esc(p.pasta) + '" target="_blank">' + esc(p.pastaNome) + '</a>' : 'nenhuma pasta vinculada') + '</p>' +
    '<input id="inPasta" type="text" placeholder="Cole o link da pasta certa do Drive">' +
    '<input id="inSenha" type="password" placeholder="Senha do escritório">' +
    '<button class="btn" id="btPasta">Salvar</button><button class="btn btn2" id="btCancPasta">Cancelar</button><div id="erPasta" class="sub"></div></div>';
  document.body.appendChild(f);
  $('btCancPasta').addEventListener('click', () => f.remove());
  $('btPasta').addEventListener('click', () => {
    const link = $('inPasta').value.trim(), senha = $('inSenha').value;
    if (!link) { $('erPasta').textContent = 'Cole o link da pasta.'; return; }
    $('btPasta').disabled = true;
    $('erPasta').textContent = 'Salvando...';
    google.script.run.withSuccessHandler(r => { p.pasta = r.pasta; p.pastaNome = r.pastaNome; f.remove(); render(); })
      .withFailureHandler(e => { $('btPasta').disabled = false; $('erPasta').textContent = (e && e.message) || String(e); })
      .salvarPasta(p.protocolo, link, senha);
  });
  setTimeout(() => $('inPasta').focus(), 50);
}
function carregar() {
  google.script.run.withSuccessHandler(d => { D = d; if (!d.dono) $('btAtualizar').style.display = 'none'; registrar(d); montar(); })
    .withFailureHandler(e => { $('lista').innerHTML = '<div class="vazio">Erro ao carregar: ' + esc(e.message) + '</div>'; })
    .dadosSistema();
}
function tagSit(s) {
  const c = /^Atendid/.test(s) ? 't-ok' : /^Respondido/.test(s) ? 't-ver' : 't-and';
  return '<span class="tag ' + c + '">' + esc(s) + '</span>';
}
function montar() {
  $('lkPlanilha').href = D.planilha;
  const parados = D.andamento.filter(p => +p.dias >= 15);
  const dias = D.andamento.map(p => +p.dias).filter(n => !isNaN(n) && String(n) !== '');
  const media = dias.length ? (dias.reduce((a, b) => a + b, 0) / dias.length).toFixed(1).replace('.', ',') : '0';
  const novos = D.finalizados.filter(p => p.novo);
  $('cAnd').textContent = D.andamento.length; $('cPar').textContent = parados.length; $('cFin').textContent = D.finalizados.length;
  $('kpis').innerHTML = [
    ['', D.andamento.length, 'Protocolos em andamento'], ['verm', parados.length, 'Parados há 15 dias ou mais'], ['lar', media, 'Média de dias sem movimentação'],
    ['verde', novos.length, 'Prontos desde o seu último acesso'], ['verde', D.finalizados.length, 'Finalizados'], ['', Math.max(0, ...dias), 'Maior tempo parado (dias)']
  ].map(k => '<div class="kpi ' + k[0] + '"><b>' + k[1] + '</b><span>' + k[2] + '</span></div>').join('');
  const hoje = new Date();
  const semana = ['Domingo','Segunda-feira','Terça-feira','Quarta-feira','Quinta-feira','Sexta-feira','Sábado'][hoje.getDay()];
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const recentes = D.finalizados.slice(0, 6);
  const comAnexo = D.andamento.filter(p => anexosDe(p).length);
  $('cAnx').textContent = comAnexo.length;
  $('lateral').innerHTML = '<h3>' + hoje.getDate() + ' de ' + meses[hoje.getMonth()] + '</h3><div class="data">' + semana + '</div>' +
    '<h4>Prontos desde o seu último acesso' + (D.ultimoAcesso ? ' (' + esc(D.ultimoAcesso) + ')' : '') + '</h4>' +
    (novos.length ? novos.map(p => '<div class="item ok"><b>' + esc(p.requerente) + '</b><span>' + esc(p.protocolo) + ' · ' + esc(p.situacao) + '<br>Pronto em ' + esc(p.finalizadoEm) + '</span></div>').join('')
      : '<div class="vazio">Nenhum protocolo novo pronto.</div>') +
    '<h4>Em andamento com documento anexado</h4>' +
    (comAnexo.length ? comAnexo.map(p => { const a = anexosDe(p), u = a[a.length - 1]; return '<div class="item"><b>' + esc(p.requerente) + '</b><span>' + esc(u.data) + ' · ' + esc(u.quem) + ' anexou ' + u.n + ' doc(s)<br>Setor atual: ' + esc(p.setor) + '</span></div>'; }).join('')
      : '<div class="vazio">Nenhum.</div>') +
    '<h4 class="verm">Mais tempo parados</h4>' +
    (parados.length ? parados.slice().sort((a, b) => +b.dias - +a.dias).slice(0, 6).map(p => '<div class="item red"><b>' + esc(p.requerente) + '</b><span>' + esc(p.dias) + ' dias · ' + esc(p.setor) + '</span></div>').join('')
      : '<div class="vazio">Nenhum parado há 15 dias ou mais.</div>') +
    '<h4>Últimos finalizados</h4>' +
    (recentes.length ? recentes.map(p => '<div class="item"><b>' + esc(p.requerente) + '</b><span>' + esc(p.protocolo) + ' · ' + esc(p.finalizadoEm) + '</span></div>').join('') : '<div class="vazio">Nenhum ainda.</div>');
  $('rodape').innerHTML = 'Última consulta ao portal:<br>' + esc((D.andamento[0] || D.finalizados[0] || {}).consulta || '-') +
    '<br><a href="' + esc(D.planilha) + '" target="_blank">abrir planilha</a>';
  const sits = Array.from(new Set(D.andamento.concat(D.finalizados).map(p => p.situacao).filter(String))).sort();
  const atual = $('fSit').value;
  $('fSit').innerHTML = '<option value="">Todas as situações</option>' + sits.map(s => '<option>' + esc(s) + '</option>').join('');
  $('fSit').value = atual;
  render();
}
function anexosDe(p) {
  return (D.historico[p.protocolo] || []).map(m => {
    const r = String(m.texto).match(/\((\d+) documento\(s\) anexado\(s\)\)/);
    return r ? { n: +r[1], data: m.data, quem: m.quem } : null;
  }).filter(Boolean);
}
function celAnexos(p) {
  const a = anexosDe(p);
  if (!a.length) return '<span class="sub">-</span>';
  const u = a[a.length - 1], tot = a.reduce((s, x) => s + x.n, 0);
  return '<span class="tag t-anx" title="' + esc(a.map(x => x.data + ' - ' + x.quem + ': ' + x.n + ' doc(s)').join('\n')) + '">📎 ' + tot + ' doc(s)</span>' +
    '<div class="sub">' + esc(String(u.data).slice(0, 10)) + ' · ' + esc(u.quem) + '</div>';
}
function render() {
  if (!D) return;
  $('tituloLista').textContent = NOMES[aba];
  let l = aba === 'andamento' ? D.andamento : aba === 'parados' ? D.andamento.filter(p => +p.dias >= 15) : aba === 'anexo' ? D.andamento.filter(p => anexosDe(p).length)
    : aba === 'finalizados' ? D.finalizados : D.andamento.concat(D.finalizados);
  const q = $('busca').value.trim().toLowerCase(), s = $('fSit').value, o = $('fOrd').value;
  l = l.filter(p => (!s || p.situacao === s) &&
    (!q || [p.requerente, p.protocolo, p.objeto, p.setor, p.tipo].join(' ').toLowerCase().indexOf(q) >= 0));
  const dt = x => { const m = String(x).match(/(\d+)\/(\d+)\/(\d+)\s*(\d+)?:?(\d+)?/); return m ? new Date(m[3], m[2]-1, m[1], m[4]||0, m[5]||0).getTime() : 0; };
  l = l.slice().sort(o === 'nome' ? (a, b) => a.requerente.localeCompare(b.requerente)
    : o === 'nomeZ' ? (a, b) => b.requerente.localeCompare(a.requerente)
    : o === 'sit' ? (a, b) => a.situacao.localeCompare(b.situacao) || a.requerente.localeCompare(b.requerente)
    : o === 'sitZ' ? (a, b) => b.situacao.localeCompare(a.situacao) || a.requerente.localeCompare(b.requerente)
    : o === 'recente' ? (a, b) => dt(b.dataUlt) - dt(a.dataUlt)
    : (a, b) => (b.novo - a.novo) || ((+b.dias || -1) - (+a.dias || -1)));
  if (!l.length) { $('lista').innerHTML = '<div class="vazio">Nenhum protocolo encontrado.</div>'; return; }
  $('lista').innerHTML = '<table><thead><tr><th class="ordenavel" onclick="ordenarNome()" title="Clique para ordenar por nome">Requerente ' + (o === 'nome' ? '▲' : o === 'nomeZ' ? '▼' : '⇅') + '</th><th class="ordenavel" onclick="ordenarSit()" title="Clique para ordenar por situação">Situação ' + (o === 'sit' ? '▲' : o === 'sitZ' ? '▼' : '⇅') + '</th><th>Anexos</th><th>Setor atual</th><th>Dias parado</th><th></th></tr></thead><tbody>' +
    l.map((p, i) => {
      const h = (D.historico[p.protocolo] || []).map(m => '<div><b>' + esc(m.data) + '</b> · ' + esc(m.texto) + '</div>').join('');
      const temDias = p.dias !== '' && !isNaN(+p.dias);
      return '<tr class="' + (p.novo ? 'novo' : '') + '"><td><div class="nome">' + (p.finalizado ? '' : '<span class="lapis" title="Trocar a pasta vinculada" data-prot="' + esc(p.protocolo) + '">✏️</span>') + (numPasta(p) ? '<span class="npasta" title="Número da pasta">' + numPasta(p) + '</span>' : '') + (p.pasta ? '<a class="lkn" href="' + esc(p.pasta) + '" target="_blank" title="Abrir a pasta: ' + esc(p.pastaNome) + '">' + esc(p.requerente) + '</a>' : esc(p.requerente)) + (p.novo ? ' <span class="tag t-ok">Novo</span>' : '') + (!p.pasta && !p.finalizado ? ' <span class="tag t-red">sem pasta vinculada</span>' : '') + '</div>' +
        '<div class="sub">' + esc(p.protocolo) + ' · ' + esc(p.objeto || p.tipo) + '</div></td>' +
        '<td>' + tagSit(p.situacao) + (p.finalizadoEm ? '<div class="sub">Pronto em ' + esc(p.finalizadoEm) + '</div>' : '') + '</td>' +
        '<td>' + celAnexos(p) + '</td>' + '<td>' + (!p.finalizado && anexosDe(p).length ? '<span class="tag t-anx">Enviado</span> ' : '') + esc(p.setor) + '<div class="sub">Últ. mov.: ' + esc(p.dataUlt) + '</div></td>' +
        '<td>' + (temDias ? '<span class="dias' + (+p.dias >= 15 ? ' alto' : '') + '">' + esc(p.dias) + '</span>' : '-') + '</td>' +
        '<td>' + (p.processo ? '<a class="lk" href="' + esc(p.processo) + '" target="_blank">Processo</a>' : '') +
        (p.comprovante ? '<a class="lk" href="' + esc(p.comprovante) + '" target="_blank">Comprovante</a>' : '') +
        (h ? '<span class="lk" onclick="var e=document.getElementById(\'h' + i + '\');e.style.display=e.style.display===\'table-row\'?\'none\':\'table-row\'">Histórico</span>' : '') + '</td></tr>' +
        '<tr class="det" id="h' + i + '" style="display:none"><td colspan="6"><div class="hist"><div><b>Última movimentação:</b> ' + esc(p.ultima) + '</div>' + h + '</div></td></tr>';
    }).join('') + '</tbody></table>';
}
function ordenarSit() { $('fOrd').value = $('fOrd').value === 'sit' ? 'sitZ' : 'sit'; render(); }
function ordenarNome() { $('fOrd').value = $('fOrd').value === 'nome' ? 'nomeZ' : 'nome'; render(); }
function atualizar() {
  const b = $('btAtualizar'); b.disabled = true; b.textContent = 'Atualizando... (pode levar 1 min)';
  google.script.run.withSuccessHandler(m => { b.disabled = false; b.textContent = '↻ Atualizar agora'; $('msg').style.display = 'block'; $('msg').textContent = m; carregar(); })
    .withFailureHandler(e => { b.disabled = false; b.textContent = '↻ Atualizar agora'; $('msg').style.display = 'block'; $('msg').textContent = 'Erro: ' + e.message; })
    .atualizarPelaPagina();
}
$('menu').onclick = e => { const a = e.target.closest('a'); const v = a && a.dataset.v; if (!v) return; aba = v;
  Array.from($('menu').children).forEach(x => x.classList.toggle('on', x.dataset.v === v)); render(); };
$('busca').oninput = render; $('fSit').onchange = render; $('fOrd').onchange = render;
carregar();
</script>
</body></html>`;


/* ------------------- Verificação da pasta a cada 15 minutos ------------------- */

/** Olha só a pasta de comprovantes; se entrou protocolo novo, já preenche pelo portal. */
function verificarPastaAuto_base_() {
  const lock = LockService.getScriptLock();
  let falha = null;
  if (!lock.tryLock(5000)) return;                 // outra atualização em andamento
  try {
    const aba = planilha_().getSheetByName(ABA_PROT);
    const r = buscarNovosComprovantes_(aba);
    if (r.adicionados.length) {
      Logger.log('Novos comprovantes: ' + r.adicionados.join(', '));
      atualizarAndamento();
    }
    if (r.falhas.length) Logger.log('QR não lido: ' + r.falhas.join(', '));
  } catch (e) {
    falha = e;
  } finally {
    lock.releaseLock();
  }
  try { avisarErros_(falha); } catch (e) { Logger.log('Aviso por e-mail: ' + e); }
}

/** Liga a verificação da pasta a cada 15 minutos (rodar uma vez). */
function ativarVerificacaoPasta() {
  desativarVerificacaoPasta();
  ScriptApp.newTrigger('verificarPastaAuto').timeBased().everyMinutes(15).create();
}

function desativarVerificacaoPasta() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'verificarPastaAuto')
    .forEach(t => ScriptApp.deleteTrigger(t));
}


/* ---------------------- PASTAS DOS CLIENTES E ANEXOS DA PREFEITURA ---------------------- */

const PASTA_PEDIDOS = 'ID_PASTA_PEDIDOS';    // 101 Pedidos administrativos protocolados
const PASTA_COMPLETOS = 'ID_PASTA_COMPLETOS';  // 000 PROCESSOS ADM COMPLETOS

function semAcento_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); }

/** Subpastas de clientes (101 e 000 PROCESSOS ADM COMPLETOS), guardadas por 10 min. */
function pastasClientes_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('pastasClientes');
  if (c) return JSON.parse(c);
  const lista = [];
  [[PASTA_PEDIDOS, 2], [PASTA_COMPLETOS, 1]].forEach(([id, peso]) => {
    const it = DriveApp.getFolderById(id).getFolders();
    while (it.hasNext()) {
      const f = it.next(), nome = f.getName();
      if (/^0+\s/.test(nome)) continue;   // pastas de organização (000, 00 MODELOS...)
      lista.push({ id: f.getId(), nome: nome, w: semAcento_(nome).split(/[^a-z0-9]+/), peso: peso, num: parseInt(nome, 10) || 0 });
    }
  });
  try { cache.put('pastasClientes', JSON.stringify(lista), 600); } catch (e) {}
  return lista;
}

/** Pasta do cliente pelo nome do requerente (primeiro nome + quase todos os sobrenomes). */
function pastaDoCliente_(nome, lista) {
  const toks = semAcento_(nome).split(/\s+/).filter(t => t.length > 2);
  if (toks.length < 2) return null;
  let melhor = null, nota = -1;
  lista.forEach(p => {
    if (p.w.indexOf(toks[0]) < 0) return;
    const ac = toks.filter(t => p.w.indexOf(t) >= 0).length;
    if (ac < Math.max(2, toks.length - 1)) return;
    const n = ac * 100000 + p.peso * 10000 + p.num;
    if (n > nota) { nota = n; melhor = p; }
  });
  return melhor;
}

/** Acrescenta o link da pasta do cliente em cada protocolo (para a página). */
function comPasta_(lista) {
  let pastas = [];
  try { pastas = pastasClientes_(); } catch (e) { Logger.log('Pastas: ' + e); }
  const vinc = vinculos_(), antes = JSON.stringify(vinc);
  lista.forEach(p => {
    const x = pastaVinculada_(p.protocolo, p.requerente, pastas, vinc);
    p.pasta = x ? 'https://drive.google.com/drive/folders/' + x.id : '';
    p.pastaNome = x ? x.nome : '';
  });
  if (JSON.stringify(vinc) !== antes) salvarVinculos_(vinc);
  return lista;
}

/** Menu: baixa agora os anexos da Prefeitura. */
function baixarAnexos_base_() {
  const n = baixarAnexos_(planilha_());
  const msg = (n.parcial ? 'Ainda faltam anexos (limite de tempo): rode de novo.\n' : '') + (n.salvos.length ? 'Anexos salvos:\n' + n.salvos.join('\n') : 'Nenhum anexo novo.');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg + (n.semPasta.length ? '\n\nSem pasta encontrada: ' + n.semPasta.join(', ') : '')); } catch (e) {}
  return msg;
}

/**
 * Para cada protocolo em andamento, baixa os documentos que a Prefeitura anexou no portal
 * e salva na pasta do cliente. Pula: o requerimento inicial, o que o próprio requerente
 * anexou, pastas marcadas FINALIZADO, arquivos com o mesmo nome já na pasta e o que já foi baixado.
 */
function baixarAnexos_(ss, limite) {
  limite = limite || 240000;
  const res = { salvos: [], semPasta: [] };
  const nomes = [], forms = [], fins = [], protos = [];
  [ABA_PROT, ABA_FIN].forEach(nm => {
    const aba = ss.getSheetByName(nm);
    if (!aba || aba.getLastRow() < 2) return;
    const q = aba.getLastRow() - 1;
    nomes.push(...aba.getRange(2, 2, q, 1).getDisplayValues());
    protos.push(...aba.getRange(2, 1, q, 1).getDisplayValues());
    forms.push(...aba.getRange(2, 11, q, 1).getFormulas());
    for (let k = 0; k < q; k++) fins.push(nm === ABA_FIN);
  });
  const n = nomes.length;
  if (!n) return res;
  const props = PropertiesService.getScriptProperties();
  const feitos = JSON.parse(props.getProperty('anexosBaixados2') || '{}');
  const t0 = Date.now();   // limite de tempo: o resto fica para a próxima execução
  const urlArq = API.replace(/oxy\/processo-digital\/process\/?$/, 'files/url-presigned');
  const pastas = pastasClientes_();
  const vinc = vinculos_();

  for (let i = 0; i < n; i++) {
    const uuid = (String(forms[i][0]).match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i) || [])[0];
    if (!uuid || !nomes[i][0]) continue;
    if (Date.now() - t0 > limite) { res.parcial = true; break; }
    if (fins[i] && feitos['p:' + uuid]) continue;   // finalizado já conferido
    let movs;
    try {
      const r = UrlFetchApp.fetch(API + 'progress/' + uuid, { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) continue;
      movs = JSON.parse(r.getContentText());
    } catch (e) { Logger.log('Anexos ' + nomes[i][0] + ': ' + e); continue; }

    const docs = [];
    movs.forEach((m, k) => {
      if (k === 0 || (m.user && m.user.external) || /requerente anexou/i.test(m.description || '')) return;
      (m.documents || []).forEach(d => { if (!d.disabled && !feitos[d.id]) docs.push(d); });
    });
    if (!docs.length) { if (fins[i]) feitos['p:' + uuid] = 1; continue; }

    const p = pastaVinculada_(protos[i][0], nomes[i][0], pastas, vinc);
    if (!p) { res.semPasta.push(nomes[i][0]); continue; }
    const pasta = DriveApp.getFolderById(p.id);

    docs.forEach(d => {
      if (Date.now() - t0 > limite + 30000) { res.parcial = true; return; }
      try {
        if (pasta.getFilesByName(d.name).hasNext()) { feitos[d.id] = 1; return; }
        const j = JSON.parse(UrlFetchApp.fetch(urlArq + '?fileKey=' + encodeURIComponent(d.key) + '&secure=' + (d.secure ? 'true' : 'false')).getContentText());
        const h = j.customerKey ? {
          'x-amz-server-side-encryption-customer-algorithm': j.customerAlgorithm,
          'x-amz-server-side-encryption-customer-key': j.customerKey,
          'x-amz-server-side-encryption-customer-key-MD5': j.customerKeyMD5
        } : {};
        const arq = UrlFetchApp.fetch(j.url, { headers: h, muteHttpExceptions: true });
        if (arq.getResponseCode() !== 200) { Logger.log('Anexo ' + d.name + ': HTTP ' + arq.getResponseCode()); return; }
        const blob = arq.getBlob();
        if (mesmoTamanho_(pasta, blob.getBytes().length)) { feitos[d.id] = 1; return; }
        const nome = nomeArquivo_(d.name, nomes[i][0], pasta);
        pasta.createFile(blob.setName(nome));
        feitos[d.id] = 1;
        res.salvos.push(nomes[i][0] + ': ' + nome);
      } catch (e) { Logger.log('Anexo ' + d.name + ': ' + e); }
    });
  }
  props.setProperty('anexosBaixados2', JSON.stringify(feitos));
  salvarVinculos_(vinc);
  return res;
}

/* ------------------------- ACESSO PELA PÁGINA E AVISO DE ERRO ------------------------- */

/** true só quando quem abriu a página é a conta dona do script (as outras só visualizam). */
function ehDono_() {
  try {
    const a = Session.getActiveUser().getEmail();
    return !!a && a === Session.getEffectiveUser().getEmail();
  } catch (e) { return false; }
}

const EMAIL_AVISO = 'seu-email@exemplo.com';

/**
 * Procura erros de fórmula (#ERROR!, #REF!...) nas abas, falha da atualização ou atualização
 * parada há mais de 20 h, e manda e-mail. Não repete o mesmo aviso antes de 24 h.
 */
function avisarErros_(falha) {
  const ss = planilha_();
  const probs = [];
  if (falha) probs.push('A atualização parou com o erro: ' + (falha.message || falha));
  [ABA_PROT, ABA_FIN, 'Painel'].forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (!sh || sh.getLastRow() < 1) return;
    sh.getDataRange().getDisplayValues().forEach((l, i) => l.forEach((c, j) => {
      if (/^#(ERROR!|REF!|N\/A|VALUE!|NAME\?|DIV\/0!|NUM!|NULL!)/.test(c)) probs.push(nome + '!' + sh.getRange(i + 1, j + 1).getA1Notation() + ': ' + c);
    }));
  });
  const props = PropertiesService.getScriptProperties();
  const ok = Number(props.getProperty('ultimaAtualizacaoOk') || 0);
  if (!ok) props.setProperty('ultimaAtualizacaoOk', String(Date.now()));
  else if (Date.now() - ok > 20 * 3600 * 1000) probs.push('A atualização automática não termina com sucesso desde ' + Utilities.formatDate(new Date(ok), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + '.');

  const chave = probs.slice(0, 30).join('\n');
  const ant = props.getProperty('ultimoAviso') || '';
  if (chave && chave !== ant) registrarSaude_('alerta', 'ALERTA', Date.now(), chave);
  if (!chave) { if (ant) props.deleteProperty('ultimoAviso'); return; }
  if (chave === ant && Date.now() - Number(props.getProperty('ultimoAvisoEm') || 0) < 24 * 3600 * 1000) return;
  MailApp.sendEmail({
    to: EMAIL_AVISO,
    subject: 'Planilha de protocolos: erro encontrado',
    body: 'O script da planilha "Acompanhamento de Protocolos - Prefeitura de Foz" encontrou problema(s):\n\n' +
      chave + (probs.length > 30 ? '\n... e mais ' + (probs.length - 30) + '.' : '') +
      '\n\nPlanilha: ' + ss.getUrl() + '\n\nAviso automático (verificação a cada 15 minutos).'
  });
  props.setProperty('ultimoAviso', chave);
  props.setProperty('ultimoAvisoEm', String(Date.now()));
}

/** Rodar pelo editor quando aparecer erro de permissão: pede de novo todas as autorizações. */
function autorizarTudo() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  DriveApp.getFolderById(PASTA_PEDIDOS).getName();
  return 'ok';
}

/* ------------------------------ CACHE E ORGANIZAÇÃO ------------------------------ */

/** Lê do cache um objeto guardado em partes (o cache aceita até 100 KB por chave). */
function lerCache_(k) {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(k + '_n') || 0);
    if (!n) return null;
    const chaves = Array.from({ length: n }, (_, i) => k + '_' + i);
    const partes = c.getAll(chaves);
    if (chaves.some(x => partes[x] == null)) return null;
    return JSON.parse(chaves.map(x => partes[x]).join(''));
  } catch (e) { return null; }
}

function salvarCache_(k, obj) {
  try {
    const s = JSON.stringify(obj), m = {};
    let n = 0;
    for (let i = 0; i < s.length; i += 40000) m[k + '_' + n++] = s.slice(i, i + 40000);
    m[k + '_n'] = String(n);
    CacheService.getScriptCache().putAll(m, 600);
  } catch (e) { Logger.log('Cache: ' + e); }
  return obj;
}

/** Depois de cada atualização a página volta a ler a planilha. */
function limparCache_() {
  try { CacheService.getScriptCache().remove('dadosPagina_n'); } catch (e) {}
}

/** Mantém as abas em ordem: Protocolos por nome (linhas vazias vão para o fim), Finalizados do mais recente. */
function organizarPlanilha_(ss) {
  const p = ss.getSheetByName(ABA_PROT);
  if (p && p.getLastRow() > 2) p.getRange(2, 1, p.getLastRow() - 1, p.getLastColumn()).sort({ column: 2, ascending: true });
  const f = ss.getSheetByName(ABA_FIN);
  if (f && f.getLastRow() > 2) f.getRange(2, 1, f.getLastRow() - 1, Math.max(14, f.getLastColumn())).sort({ column: 14, ascending: false });
}

/** Algum arquivo da pasta com o mesmo tamanho (evita baixar duas vezes o mesmo documento). */
function mesmoTamanho_(pasta, tam) {
  const it = pasta.getFiles();
  while (it.hasNext()) if (it.next().getSize() === tam) return true;
  return false;
}

/** "FRANCISCO - TERMO DE EXONERAÇÃO.pdf" -> "100 Termo de exoneração - Francisco de Assis Angelo.pdf" (próximo número da pasta, nome no fim). */
function nomeArquivo_(orig, cliente, pasta) {
  const m = String(orig).match(/\.[A-Za-z0-9]{2,4}$/);
  const ext = m ? m[0].toLowerCase() : '.pdf';
  let base = m ? String(orig).slice(0, -m[0].length) : String(orig);
  const toks = semAcento_(cliente).split(/\s+/).filter(t => t.length > 2);
  base = base.split(/[\s_]+/).filter(w => toks.indexOf(semAcento_(w)) < 0).join(' ');
  const limpa = s => s.replace(/^[\s\-–.,]+|[\s\-–.,]+$/g, '').replace(/^(d[aeo]s?\s+)+/i, '').replace(/(\s+d[aeo]s?)+$/i, '').replace(/^[\s\-–.,]+|[\s\-–.,]+$/g, '');
  base = limpa(limpa(base)).replace(/\s*-\s*(-\s*)+/g, ' - ').replace(/\s+/g, ' ');
  if (base && !/[a-zà-ÿ]/.test(base)) base = base.toLowerCase();
  base = base ? base.charAt(0).toUpperCase() + base.slice(1) : 'Documento da Prefeitura';
  let maior = 99;   // arquivos novos numerados a partir de 100
  const it = pasta.getFiles();
  while (it.hasNext()) { const x = it.next().getName().match(/^(\d{3})\s/); if (x && +x[1] > maior) maior = +x[1]; }
  return (maior + 1) + ' ' + base + ' - ' + cliente + ext;
}

/** Grava quem abriu a página (aba Acessos, mais recente no topo). Visitantes informam nome e e-mail na entrada. */
function registrarAcesso(email, nome) {
  try {
    const dono = ehDono_();
    const mail = String(email || '').trim().slice(0, 120) || (dono ? Session.getActiveUser().getEmail() : '');
    const quem = String(nome || '').trim().slice(0, 80) || (dono ? 'Escritório' : '');
    if (!mail && !quem) return;
    const ss = planilha_();
    let sh = ss.getSheetByName('Acessos');
    if (!sh) { sh = ss.insertSheet('Acessos'); sh.setFrozenRows(1); }
    sh.getRange(1, 1, 1, 4).setValues([['Data e hora', 'Nome', 'E-mail', 'Como']]).setFontWeight('bold');
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 4).setValues([[new Date(), quem, mail, dono ? 'conta do escritório' : 'digitado na página']]).setFontWeight('normal');
    sh.getRange(2, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  } catch (e) { Logger.log('Acesso: ' + e); }
}

/* ------------------------- PASTA DO CLIENTE (vínculo por protocolo) ------------------------- */

/** Vínculos protocolo -> ID da pasta, guardados nas propriedades do script (não mexe na planilha). */
function vinculos_() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('pastasVinculadas') || '{}');
}
function salvarVinculos_(v) {
  PropertiesService.getScriptProperties().setProperty('pastasVinculadas', JSON.stringify(v));
}

/** Pasta do protocolo: usa o vínculo salvo; se não houver, procura pelo nome UMA vez e grava o vínculo. */
function pastaVinculada_(protocolo, nome, lista, vinc) {
  const id = vinc[protocolo];
  if (id) {
    const x = lista.filter(p => p.id === id)[0];
    if (x) return x;
    try { return { id: id, nome: DriveApp.getFolderById(id).getName() }; } catch (e) { return null; }
  }
  const p = pastaDoCliente_(nome, lista);
  if (p && protocolo) vinc[protocolo] = p.id;
  return p;
}

/** Página: troca a pasta vinculada a um protocolo (exige a senha cadastrada em SENHA_EDICAO). */
function salvarPasta(protocolo, link, senha) {
  const certa = PropertiesService.getScriptProperties().getProperty('SENHA_EDICAO');
  if (!certa) throw new Error('A senha de edição ainda não foi cadastrada no Apps Script.');
  if (String(senha || '') !== certa) { Utilities.sleep(1500); throw new Error('Senha incorreta.'); }
  const s = String(link || '').trim();
  const m = s.match(/folders\/([-\w]{20,})/) || s.match(/[?&]id=([-\w]{20,})/) || s.match(/^([-\w]{20,})$/);
  if (!m) throw new Error('Cole o link de uma PASTA do Google Drive.');
  let f;
  try { f = DriveApp.getFolderById(m[1]); f.getName(); } catch (e) { throw new Error('Pasta não encontrada no Drive (confira o link).'); }
  const vinc = vinculos_();
  vinc[String(protocolo)] = f.getId();
  salvarVinculos_(vinc);
  limparCache_();
  return { pasta: 'https://drive.google.com/drive/folders/' + f.getId(), pastaNome: f.getName() };
}

/* ------------------------- RESUMO SEMANAL (PDF anexado na Agenda) ------------------------- */

/** Rodar UMA vez pelo editor: cria o gatilho de toda segunda, 8h. */
function ativarResumoSemanal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'resumoSemanal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  return 'Resumo semanal ativado: toda segunda às 8h.';
}
function desativarResumoSemanal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'resumoSemanal').forEach(t => ScriptApp.deleteTrigger(t));
}

/* ------------------------- ARQUIVO MENSAL DE ACESSOS */
function ativarArquivoAcessos() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'arquivarAcessos').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('arquivarAcessos').timeBased().onMonthDay(1).atHour(6).create();
  return 'Arquivo de acessos ativado: todo dia 1º às 6h.';
}
function desativarArquivoAcessos() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'arquivarAcessos').forEach(t => ScriptApp.deleteTrigger(t));
}

function pastaAcessos_() {
  const mae = DriveApp.getFolderById(PASTA_COMPROVANTES);
  const it = mae.getFoldersByName('Relatórios de acessos');
  return it.hasNext() ? it.next() : mae.createFolder('Relatórios de acessos');
}

/** Monta o PDF de um mês de acessos e salva em "Relatórios de acessos". */
function relatorioAcessos_(chave, linhas, prefixo) {
  const tz = Session.getScriptTimeZone();
  const MESES = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const [ano, mes] = chave.split('-').map(Number);
  const h = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const fmt = d => Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm');
  const ord = linhas.slice().sort((a, b) => a[0] - b[0]);
  const pessoas = {};
  ord.forEach(r => {
    const k = (String(r[1]).trim() + '|' + String(r[2]).trim()).toLowerCase();
    const p = pessoas[k] = pessoas[k] || { nome: r[1], email: r[2], qtd: 0, primeiro: r[0], ultimo: r[0] };
    p.qtd++; p.ultimo = r[0];
  });
  const lista = Object.keys(pessoas).map(k => pessoas[k]).sort((a, b) => b.qtd - a.qtd);
  const css = 'body{font-family:Arial,sans-serif;font-size:11px;color:#222}h1{font-size:17px;color:#1a3d7c;margin:0 0 4px}' +
    'h2{font-size:13px;color:#1a3d7c;margin:18px 0 6px}p{margin:2px 0}table{border-collapse:collapse;width:100%}' +
    'th{color:#1a3d7c;border-bottom:2px solid #1a3d7c;text-align:left;padding:4px 6px}td{border-bottom:1px solid #ddd;padding:3px 6px}';
  let html = '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<h1>Acessos ao Sistema de Protocolos — ' + MESES[mes - 1] + ' de ' + ano + '</h1>' +
    '<p>' + ord.length + ' acesso(s) de ' + lista.length + ' pessoa(s). Gerado em ' + fmt(new Date()) + '.</p>' +
    '<h2>Resumo por pessoa</h2><table><tr><th>Nome</th><th>E-mail</th><th>Acessos</th><th>Primeiro</th><th>Último</th></tr>' +
    lista.map(p => '<tr><td>' + h(p.nome) + '</td><td>' + h(p.email) + '</td><td>' + p.qtd + '</td><td>' + fmt(p.primeiro) + '</td><td>' + fmt(p.ultimo) + '</td></tr>').join('') +
    '</table><h2>Todos os acessos</h2><table><tr><th>Data e hora</th><th>Nome</th><th>E-mail</th><th>Como</th></tr>' +
    ord.map(r => '<tr><td>' + fmt(r[0]) + '</td><td>' + h(r[1]) + '</td><td>' + h(r[2]) + '</td><td>' + h(r[3]) + '</td></tr>').join('') +
    '</table></body></html>';
  const nome = (prefixo || '') + 'Acessos ' + chave + ' (' + MESES[mes - 1] + ' ' + ano + ').pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'acessos.html').getAs('application/pdf').setName(nome);
  return pastaAcessos_().createFile(pdf);
}

/** Todo dia 1º: gera um PDF por mês já encerrado e só depois apaga essas linhas da aba Acessos. */
function arquivarAcessos_base_() {
  const sh = planilha_().getSheetByName('Acessos');
  if (!sh || sh.getLastRow() < 2) return 'Nada para arquivar.';
  const agora = new Date();
  const corte = new Date(agora.getFullYear(), agora.getMonth(), 1);
  const tz = Session.getScriptTimeZone();
  const dados = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const grupos = {}, apagar = [];
  dados.forEach((r, i) => {
    if (!(r[0] instanceof Date) || r[0] >= corte) return;
    const k = Utilities.formatDate(r[0], tz, 'yyyy-MM');
    (grupos[k] = grupos[k] || []).push(r);
    apagar.push(i + 2);
  });
  const meses = Object.keys(grupos).sort();
  if (!meses.length) return 'Nada para arquivar.';
  const arquivos = meses.map(k => relatorioAcessos_(k, grupos[k]));
  if (arquivos.some(f => !f || !f.getId())) throw new Error('PDF não foi salvo; nenhuma linha apagada.');
  apagar.sort((a, b) => b - a);
  let i = 0;
  while (i < apagar.length) {
    let j = i;
    while (j + 1 < apagar.length && apagar[j + 1] === apagar[j] - 1) j++;
    sh.deleteRows(apagar[j], j - i + 1);
    i = j + 1;
  }
  return 'Arquivados ' + apagar.length + ' acesso(s) em ' + arquivos.length + ' PDF(s): ' + arquivos.map(f => f.getName()).join(', ');
}

/** Teste: gera o PDF do mês atual (prefixo TESTE) sem apagar nada. */
function testeRelatorioAcessos() {
  const sh = planilha_().getSheetByName('Acessos');
  const tz = Session.getScriptTimeZone();
  const k = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const linhas = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] instanceof Date && Utilities.formatDate(r[0], tz, 'yyyy-MM') === k);
  return relatorioAcessos_(k, linhas, 'TESTE ').getId();
}

/* ------------------------- SAÚDE DO SISTEMA (aba "Saúde")
 * As funções abaixo são as "portas de entrada" (gatilhos, menu, página).
 * Cada uma chama a versão *_base_ e grava em "Saúde" se terminou OK ou com ERRO.
 * verificarPastaAuto (a cada 15 min) e dadosSistema (cada abertura da página) só gravam erros. */
let SAUDE_NIVEL_ = 0;
const SAUDE_SO_ERRO_ = { verificarPastaAuto: 1, dadosSistema: 1 };
function atualizarAndamento() { return comSaude_('atualizarAndamento', atualizarAndamento_base_, arguments); }
function verificarPastaAuto() { return comSaude_('verificarPastaAuto', verificarPastaAuto_base_, arguments); }
function baixarAnexos() { return comSaude_('baixarAnexos', baixarAnexos_base_, arguments); }
function resumoSemanal() { return comSaude_('resumoSemanal', resumoSemanal_base_, arguments); }
function arquivarAcessos() { try { verificarTamanho_(); } catch (e) { Logger.log('Tamanho: ' + e); } return comSaude_('arquivarAcessos', arquivarAcessos_base_, arguments); }
function dadosSistema() { return comSaude_('dadosSistema', dadosSistema_base_, arguments); }
function atualizarPelaPagina() { return comSaude_('atualizarPelaPagina', atualizarPelaPagina_base_, arguments); }

function comSaude_(nome, fn, args) {
  const t0 = Date.now();
  SAUDE_NIVEL_++;
  try {
    const r = fn.apply(null, args);
    if (!SAUDE_SO_ERRO_[nome]) registrarSaude_(nome, 'OK', t0, typeof r === 'string' ? r : '');
    return r;
  } catch (e) {
    if (!(e && e.saudeGravada)) {
      registrarSaude_(nome, 'ERRO', t0, (e && (e.stack || e.message)) || String(e));
      try { e.saudeGravada = true; } catch (x) {}
    }
    throw e;
  } finally {
    SAUDE_NIVEL_--;
  }
}

function registrarSaude_(nome, status, t0, detalhe) {
  try {
    const ss = planilha_();
    let sh = ss.getSheetByName('Saúde');
    if (!sh) {
      sh = ss.insertSheet('Saúde');
      sh.getRange(1, 1, 1, 5).setValues([['Data e hora', 'Função', 'Status', 'Duração (s)', 'Detalhe']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 5).setValues([[new Date(), nome, status, Math.round((Date.now() - t0) / 100) / 10, String(detalhe || '').slice(0, 1000)]])
      .setFontWeight('normal').setFontColor(status === 'OK' ? '#1b5e20' : '#b71c1c');
    sh.getRange(2, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    if (sh.getLastRow() > 1001) sh.deleteRows(1002, sh.getLastRow() - 1001);   // guarda as 1000 mais recentes
  } catch (e) { Logger.log('Saúde: ' + e); }
}

/* ------------------------- TAMANHO DA PLANILHA E ARQUIVO DE MOVIMENTAÇÕES
 * O Google limita a planilha inteira a 10 milhões de células (contando as vazias reservadas).
 * verificarTamanho_ roda 1x por mês (dia 1º, junto com arquivarAcessos): avisa por e-mail e grava ALERTA na Saúde.
 * arquivarMovimentacoes só roda pelo menu e sempre pede confirmação antes de mexer. */
const LIMITE_CELULAS_ = 10000000;
const LIMITE_MOV_ = 20000;

function tamanhoPlanilha_() {
  const ss = planilha_();
  let celulas = 0;
  ss.getSheets().forEach(sh => { celulas += sh.getMaxRows() * sh.getMaxColumns(); });
  const mov = ss.getSheetByName(ABA_MOV);
  return { celulas: celulas, pct: celulas / LIMITE_CELULAS_ * 100, linhasMov: mov ? Math.max(0, mov.getLastRow() - 1) : 0 };
}

function avisosTamanho_() {
  const t = tamanhoPlanilha_();
  const out = [];
  if (t.pct >= 50) out.push('A planilha passou de ' + Math.floor(t.pct / 10) * 10 + '% do limite de células do Google (10 milhões). Peça ao Claude para arquivar o histórico antigo.');
  if (t.linhasMov >= LIMITE_MOV_) out.push('A aba Movimentações passou de ' + LIMITE_MOV_ + ' linhas. Para arquivar o histórico antigo: menu Protocolos > Arquivar histórico antigo (pede sua confirmação).');
  return out;
}

function verificarTamanho_() {
  const avisos = avisosTamanho_();
  if (!avisos.length) return;
  registrarSaude_('tamanho', 'ALERTA', Date.now(), avisos.join(' | '));
  MailApp.sendEmail({
    to: EMAIL_AVISO,
    subject: 'Planilha de protocolos: espaço',
    body: avisos.join('\n\n') + '\n\nPlanilha: ' + planilha_().getUrl() + '\n\nAviso automático (verificação mensal, dia 1º).'
  });
}

function arquivoMovimentacoes_() {
  const mae = DriveApp.getFolderById(PASTA_COMPROVANTES);
  const it = mae.getFilesByName('Arquivo de movimentações');
  if (it.hasNext()) return SpreadsheetApp.open(it.next());
  const novo = SpreadsheetApp.create('Arquivo de movimentações');
  DriveApp.getFileById(novo.getId()).moveTo(mae);
  const sh = novo.getSheets()[0];
  sh.setName('Movimentações arquivadas');
  sh.getRange(1, 1, 1, 7).setValues([['Protocolo', 'Requerente', 'Data', 'Quem', 'Movimentação', 'Registrado na planilha em', 'Arquivado em']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return novo;
}

/** Menu: leva para a planilha "Arquivo de movimentações" o histórico dos protocolos finalizados há mais de 1 ano. Pede confirmação. */
function arquivarMovimentacoes() { return comSaude_('arquivarMovimentacoes', arquivarMovimentacoes_base_, arguments); }
function arquivarMovimentacoes_base_() {
  const ss = planilha_();
  const ui = SpreadsheetApp.getUi();
  const fin = ss.getSheetByName(ABA_FIN), mov = ss.getSheetByName(ABA_MOV);
  if (!fin || !mov || mov.getLastRow() < 2 || fin.getLastRow() < 2) { ui.alert('Nada para arquivar.'); return 'Nada para arquivar.'; }
  const corte = Date.now() - 365 * 864e5;
  const antigos = {};
  fin.getRange(2, 1, fin.getLastRow() - 1, 14).getValues().forEach(l => {
    const t = paraData_(l[13]).getTime();
    if (l[0] && t > 0 && t < corte) antigos[String(l[0])] = true;
  });
  const n = mov.getLastRow() - 1;
  const dados = mov.getRange(2, 1, n, 6).getValues();
  const sai = [], fica = [], prots = {};
  dados.forEach(l => {
    if (antigos[String(l[0])]) { sai.push(l); prots[l[0]] = true; } else fica.push(l);
  });
  if (!sai.length) { ui.alert('Nenhum protocolo finalizado há mais de 1 ano. Nada para arquivar.'); return 'Nada para arquivar.'; }
  const qtd = Object.keys(prots).length;
  const resp = ui.alert('Arquivar histórico antigo',
    sai.length + ' linha(s) de ' + qtd + ' protocolo(s) finalizado(s) há mais de 1 ano vão para a planilha "Arquivo de movimentações" ' +
    '(na pasta de comprovantes) e saem da aba Movimentações. Primeiro copia, confere e só depois tira daqui. Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return 'Cancelado pelo usuário.';
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const sa = arquivoMovimentacoes_().getSheets()[0];
    const ini = sa.getLastRow() + 1;
    const agora = new Date();
    sa.getRange(ini, 1, sai.length, 7).setValues(sai.map(l => l.concat([agora])));
    sa.getRange(ini, 1, sai.length, 1).setNumberFormat('@');
    SpreadsheetApp.flush();
    if (sa.getLastRow() < ini + sai.length - 1) throw new Error('Cópia incompleta: nada foi tirado da aba Movimentações.');
    mov.getRange(2, 1, n, 6).clearContent();
    if (fica.length) {
      mov.getRange(2, 1, fica.length, 6).setValues(fica);
      mov.getRange(2, 1, fica.length, 1).setNumberFormat('@');
    }
    const sobra = mov.getMaxRows() - (fica.length + 1) - 500;   // devolve as linhas vazias, deixando 500 de folga
    if (sobra > 0) mov.deleteRows(fica.length + 2 + 500, sobra);
  } finally {
    lock.releaseLock();
  }
  limparCache_();
  const msg = 'Arquivadas ' + sai.length + ' linha(s) de ' + qtd + ' protocolo(s) na planilha "Arquivo de movimentações".';
  ui.alert(msg);
  return msg;
}

function dataBr_(s) {
  const m = String(s).match(/(\d+)\/(\d+)\/(\d+)\s*(\d+)?:?(\d+)?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : null;
}

function pastaResumos_() {
  const mae = DriveApp.getFolderById(PASTA_COMPROVANTES);
  const it = mae.getFoldersByName('Resumos semanais');
  return it.hasNext() ? it.next() : mae.createFolder('Resumos semanais');
}

/** Gera o PDF da semana, salva em "Resumos semanais" e cria o evento na agenda principal com o PDF anexado. */
function resumoSemanal_base_() {
  const tz = 'America/Sao_Paulo';
  const agora = new Date();
  const desde = new Date(agora.getTime() - 7 * 864e5);
  const d = dadosPlanilha_();
  const todos = d.andamento.concat(d.finalizados);
  const prontos = d.finalizados.filter(p => p.tFin >= desde.getTime());
  const anexos = [];
  todos.forEach(p => (d.historico[p.protocolo] || []).forEach(m => {
    const r = String(m.texto).match(/\((\d+) documento\(s\) anexado\(s\)\)/);
    const dt = dataBr_(m.data);
    if (r && dt && dt >= desde && !/requerente anexou/i.test(m.texto)) anexos.push({ p: p, m: m, n: +r[1] });
  }));
  const parados = d.andamento.filter(p => +p.dias >= 15).sort((a, b) => +b.dias - +a.dias);
  const e = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const dia = Utilities.formatDate(agora, tz, 'dd/MM/yyyy');
  const titulo = 'Resumo semanal de protocolos - ' + dia;
  const tabela = (cab, linhas) => linhas.length
    ? '<table><tr>' + cab.map(c => '<th>' + c + '</th>').join('') + '</tr>' + linhas.map(l => '<tr>' + l.map(c => '<td>' + e(c) + '</td>').join('') + '</tr>').join('') + '</table>'
    : '<p class="vazio">Nada nesta semana.</p>';
  const html = '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:11px;color:#1e293b}h1{font-size:18px;color:#1e3a8a;margin:0 0 4px}' +
    'h2{font-size:13px;color:#1e3a8a;margin:18px 0 6px;border-bottom:1px solid #cbd5e1;padding-bottom:3px}' +
    '.sub{color:#64748b;margin-bottom:10px}.kp td{border:1px solid #cbd5e1;padding:6px 10px;text-align:center}.kp b{font-size:16px;display:block}' +
    'table{border-collapse:collapse;width:100%}th{background:#eef2fb;text-align:left;padding:5px;border:1px solid #cbd5e1}td{padding:5px;border:1px solid #e2e8f0;vertical-align:top}' +
    '.vazio{color:#64748b}</style></head><body>' +
    '<h1>' + e(titulo) + '</h1><div class="sub">Semana de ' + Utilities.formatDate(desde, tz, 'dd/MM') + ' a ' + Utilities.formatDate(agora, tz, 'dd/MM/yyyy') + ' · Prefeitura de Foz do Iguaçu</div>' +
    '<table class="kp"><tr><td><b>' + d.andamento.length + '</b>Em andamento</td><td><b>' + prontos.length + '</b>Prontos na semana</td><td><b>' + anexos.length + '</b>Envios com documento</td><td><b>' + parados.length + '</b>Parados 15+ dias</td></tr></table>' +
    '<h2>Prontos na semana</h2>' + tabela(['Requerente', 'Protocolo', 'Objeto', 'Pronto em', 'Situação'], prontos.map(p => [p.requerente, p.protocolo, p.objeto || p.tipo, p.finalizadoEm, p.situacao])) +
    '<h2>Documentos anexados pela Prefeitura</h2>' + tabela(['Requerente', 'Protocolo', 'Data', 'Quem', 'Docs'], anexos.map(a => [a.p.requerente, a.p.protocolo, a.m.data, a.m.quem, a.n])) +
    '<h2>Parados há 15 dias ou mais</h2>' + tabela(['Requerente', 'Protocolo', 'Objeto', 'Setor atual', 'Dias'], parados.map(p => [p.requerente, p.protocolo, p.objeto || p.tipo, p.setor, p.dias])) +
    '<p class="vazio" style="margin-top:16px">Gerado automaticamente em ' + Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm') + '.</p></body></html>';
  const pdf = Utilities.newBlob(html, 'text/html', titulo + '.html').getAs('application/pdf').setName(titulo + '.pdf');
  const arq = pastaResumos_().createFile(pdf);

  const desc = 'Em andamento: ' + d.andamento.length + '\nProntos na semana: ' + prontos.length +
    (prontos.length ? ' (' + prontos.map(p => p.requerente).join(', ') + ')' : '') +
    '\nEnvios com documento: ' + anexos.length + '\nParados 15+ dias: ' + parados.length +
    '\n\nPDF completo em anexo e na pasta "Resumos semanais".';
  const ini = new Date(agora); ini.setHours(8, 0, 0, 0);
  const fim = new Date(ini.getTime() + 30 * 60000);
  Calendar.Events.insert({
    summary: titulo,
    description: desc,
    start: { dateTime: ini.toISOString(), timeZone: tz },
    end: { dateTime: fim.toISOString(), timeZone: tz },
    attachments: [{ fileUrl: arq.getUrl(), title: arq.getName(), mimeType: 'application/pdf' }]
  }, 'primary', { supportsAttachments: true });
  return titulo;
}
