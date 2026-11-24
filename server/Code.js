/**
 * PONTO DA CASA — servidor (Google Apps Script)
 * Guarda as batidas numa planilha do Google Drive, com hora do servidor,
 * valida localização (raio do endereço) + celular cadastrado (+ selfie opcional) e fecha o mês (PDF + Excel).
 * GPS impreciso que pode estar dentro do raio: a batida entra como "pendente" para o empregador aprovar.
 *
 * INSTALAÇÃO (uma vez):
 * 1. script.google.com › Novo projeto › cole este arquivo em Code.gs.
 * 2. Rode a função setup() e autorize.
 * 3. Implantar › Nova implantação › Tipo: App da Web
 *    Executar como: Eu · Quem pode acessar: Qualquer pessoa.
 * 4. Copie a URL do App da Web e cole em API_URL no index.html.
 * 5. Notificações: rode configurarNotificacoes('APP_ID', 'REST_API_KEY') com os dados do OneSignal
 *    e cole o mesmo APP_ID em ONESIGNAL_APP_ID no index.html.
 */

// Estrutura no Drive: Claude ( G drive) › eSocial › { Arquivos do Claude | Espelhos de ponto › <funcionária> | DAE › <AAAA-MM> › <funcionária> }
const PASTA_ARQUIVOS_CLAUDE = '1TiBQ2r6qxVnOEDYekl75LqZd-LQxuOov'; // eSocial › Arquivos do Claude
const PASTA_ESPELHOS = '14bms1HPs9CRh0fHWNxKMmlEqmsqNLbJt';        // eSocial › Espelhos de ponto
const PASTA_DAE = '1m8XHiKRkyh6ecUVxiydX8TY66RQDKESj';             // eSocial › DAE
const TZ = 'America/Sao_Paulo';
const TIPOS = ['Entrada', 'Saída p/ almoço', 'Volta do almoço', 'Saída'];
const COLS = {
  Funcionarias: ['id','nome','cpf','cargo','localId','pinHash','deviceId','token','ativo','jornada'],
  Locais: ['id','nome','endereco','lat','lng','raio'],
  Batidas: ['id','funcId','nome','data','hora','ts','tipo','localId','lat','lng','precisao','distancia','selfieId','deviceId','hashAnterior','hash','status','revisadoEm'],
  Confirmacoes: ['funcId','ym','ts','acao','motivo','ip','userAgent','deviceId','hashEspelho','hash','termoId'],
  Config: ['chave','valor']
};
const P = PropertiesService.getScriptProperties();

/* ===================== INSTALAÇÃO ===================== */
function setup() {
  if (P.getProperty('SHEET_ID')) { Logger.log('Já instalado. Planilha: ' + P.getProperty('SHEET_ID')); return; }
  const pasta = DriveApp.getFolderById(PASTA_ARQUIVOS_CLAUDE).createFolder('Ponto — sistema');
  const selfies = pasta.createFolder('Selfies');
  const ss = SpreadsheetApp.create('Ponto — registros');
  DriveApp.getFileById(ss.getId()).moveTo(pasta);
  Object.keys(COLS).forEach((nome, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(nome) : ss.insertSheet(nome);
    sh.getRange(1, 1, 1, COLS[nome].length).setValues([COLS[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // tudo como texto: datas/horas não são convertidas
  });
  P.setProperties({ SHEET_ID: ss.getId(), PASTA_ID: pasta.getId(), SELFIES_ID: selfies.getId(), SALT: Utilities.getUuid() });
  // valores iniciais — troque tudo pela área do empregador
  setCfg('adminPinHash', hashPin('1234'));
  setCfg('tolerancia', 10); setCfg('empregador', ''); setCfg('cpfEmpregador', ''); setCfg('exigirSelfie', false);
  setCfg('alertaDia', 4); setCfg('alertaHora', 18); setCfg('email', '');
  append('Locais', { id: 'casa', nome: 'Casa principal', endereco: 'A definir', lat: -23.3217, lng: -46.7266, raio: 100 });
  const j = JSON.stringify({ 1: ['08:00','12:00','13:00','17:00'], 2: ['08:00','12:00','13:00','17:00'], 3: ['08:00','12:00','13:00','17:00'], 4: ['08:00','12:00','13:00','17:00'], 5: ['08:00','12:00','13:00','17:00'] });
  append('Funcionarias', { id: 'f1', nome: 'Funcionária 1', cargo: 'Funcionária', localId: 'casa', pinHash: hashPin('1111'), ativo: true, jornada: j });
  append('Funcionarias', { id: 'f2', nome: 'Funcionária 2', cargo: 'Funcionária', localId: 'casa', pinHash: hashPin('2222'), ativo: true, jornada: j });
  // fechamento automático no dia 1 de cada mês, às 6h
  ScriptApp.newTrigger('fecharMesAnterior').timeBased().onMonthDay(1).atHour(6).inTimezone(TZ).create();
  instalarAlertaSemanal();
  ScriptApp.newTrigger('lembretes').timeBased().everyMinutes(15).create();
  Logger.log('Pronto. Planilha: ' + ss.getUrl());
}

/* ===================== API ===================== */
function doGet() { return json({ ok: true, data: 'Ponto da Casa — servidor ativo' }); }
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    return json({ ok: true, data: rota(p.action, p) });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function rota(a, p) {
  switch (a) {
    case 'listFuncs': return rows('Funcionarias').filter(f => isTrue(f.ativo)).map(f => ({ id: f.id, nome: f.nome }));
    case 'login': return login(p);
    case 'punch': return punch(p);
    case 'myPunches': { const f = auth(p); return batidasDe(f.id, p.ym); }
    case 'inicio': { const f = auth(p); return { punches: batidasDe(f.id, p.ym), info: { func: pub(f), tol: +cfg('tolerancia') || 0, exigirSelfie: isTrue(cfg('exigirSelfie')), confirmacoes: rows('Confirmacoes').filter(c => c.funcId === f.id) } }; }
    case 'myInfo': { const f = auth(p); return { func: pub(f), tol: +cfg('tolerancia') || 0, exigirSelfie: isTrue(cfg('exigirSelfie')), confirmacoes: rows('Confirmacoes').filter(c => c.funcId === f.id) }; }
    case 'confirm': return confirmar(p);
    case 'adminLogin': admin(p); return true;
    case 'adminAll': return adminAll(p);
    case 'selfie': admin(p); return selfieData(p.selfieId);
    case 'saveFunc': return saveFunc(p);
    case 'resetDevice': { admin(p); updateById('Funcionarias', p.funcId, { deviceId: '', token: '' }); return true; }
    case 'saveLocal': return saveLocal(p);
    case 'review': { admin(p); updateById('Batidas', p.id, { status: p.aprovar ? 'aprovada' : 'recusada', revisadoEm: Date.now() }); return true; }
    case 'saveAdmin': return saveAdmin(p);
    case 'publico': return { pushAppId: P.getProperty('ONESIGNAL_APP_ID') || '' };
    case 'setPush': { admin(p); if (!p.appId || !p.key) throw new Error('Informe o App ID e a REST API Key.'); configurarNotificacoes(String(p.appId).trim(), String(p.key).trim()); return true; }
  }
  throw new Error('Ação desconhecida: ' + a);
}

/* ===================== FUNCIONÁRIA ===================== */
function login(p) {
  const f = rows('Funcionarias').find(x => x.id === p.funcId);
  if (!f || !isTrue(f.ativo) || f.pinHash !== hashPin(p.pin)) throw new Error('PIN incorreto.');
  if (f.deviceId && f.deviceId !== p.device) throw new Error('Este celular não é o cadastrado para você. Peça ao empregador para liberar o novo celular.');
  const token = rnd(24);
  updateById('Funcionarias', f.id, { deviceId: p.device, token: token });
  f.token = token; f.deviceId = p.device;
  const loc = rows('Locais').find(l => l.id === f.localId);
  return { token: token, func: pub(f), local: loc ? { id: loc.id, nome: loc.nome, lat: +loc.lat, lng: +loc.lng, raio: +loc.raio || 100 } : null };
}
function auth(p) {
  const f = rows('Funcionarias').find(x => x.id === p.funcId);
  if (!f || !isTrue(f.ativo) || !p.token || f.token !== p.token || f.deviceId !== p.device) throw new Error('Sessão expirada. Entre de novo.');
  return f;
}
function punch(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = auth(p);
    const loc = rows('Locais').find(l => l.id === f.localId);
    if (!loc) throw new Error('Nenhum local de trabalho cadastrado para você.');
    const lat = +p.lat, lng = +p.lng, acc = +p.acc || 0;
    if (!isFinite(lat) || !isFinite(lng)) throw new Error('Localização não recebida.');
    const dist = distM(lat, lng, +loc.lat, +loc.lng), raio = +loc.raio || 100;
    let status;
    if (dist <= raio + Math.min(acc, 50)) status = 'ok';
    else if (acc > 60 && dist - acc <= raio) status = 'pendente'; // GPS impreciso: empregador aprova
    else throw new Error('Fora do local de trabalho: você está a ' + (dist >= 1000 ? (dist / 1000).toFixed(1).replace('.', ',') + ' km' : Math.round(dist) + ' m') + ' da casa. O ponto só pode ser batido no local.');
    const temSelfie = p.selfie && String(p.selfie).indexOf('data:image') === 0;
    if (isTrue(cfg('exigirSelfie')) && !temSelfie) throw new Error('Falta a selfie.');

    const ts = Date.now(); // HORA DO SERVIDOR
    const data = Utilities.formatDate(new Date(ts), TZ, 'yyyy-MM-dd');
    const hora = Utilities.formatDate(new Date(ts), TZ, 'HH:mm');
    const todas = rows('Batidas');
    const hoje = todas.filter(b => b.funcId === f.id && b.data === data && b.status !== 'recusada');
    const ult = hoje[hoje.length - 1];
    if (ult && ts - (+ult.ts) < 120000) throw new Error('Você acabou de bater o ponto. Aguarde 2 minutos.');
    const n = hoje.length;
    const tipo = n < 4 ? TIPOS[n] : (n % 2 === 0 ? 'Entrada (extra)' : 'Saída (extra)');
    const id = Utilities.getUuid();
    let fotoId = '';
    if (temSelfie) {
      const bytes = Utilities.base64Decode(String(p.selfie).split(',')[1]);
      fotoId = DriveApp.getFolderById(P.getProperty('SELFIES_ID')).createFile(Utilities.newBlob(bytes, 'image/jpeg', data + ' ' + hora.replace(':', 'h') + ' ' + f.nome + ' ' + tipo + '.jpg')).getId();
    }
    const prev = todas.length ? todas[todas.length - 1].hash : '0';
    const hash = sha256([prev, id, f.id, ts, tipo, lat, lng].join('|'));
    const b = { id: id, funcId: f.id, nome: f.nome, data: data, hora: hora, ts: ts, tipo: tipo, localId: loc.id, lat: lat, lng: lng, precisao: Math.round(acc), distancia: Math.round(dist), selfieId: fotoId, deviceId: p.device, hashAnterior: prev, hash: hash, status: status };
    append('Batidas', b);
    if (status === 'pendente') push(['empregador'], 'Ponto — aprovar localização', f.nome + ' bateu ' + tipo + ' às ' + hora + ' com GPS impreciso (a ' + Math.round(dist) + ' m). Abra o app para aprovar.');
    return front(b);
  } finally { lock.releaseLock(); }
}
function confirmar(p) {
  const f = auth(p); const ym = String(p.ym || '');
  if (!(ym < Utilities.formatDate(new Date(), TZ, 'yyyy-MM'))) throw new Error('Só dá para confirmar um mês já encerrado.');
  const ult = rows('Confirmacoes').filter(c => c.funcId === f.id && c.ym === ym).pop();
  if (ult && ult.acao !== 'contestado') return ult;
  const func = pub(f); const E = espelho(func, batidasDe(f.id, ym), ym, +cfg('tolerancia') || 0);
  const ts = Date.now(); const acao = p.acao === 'contestado' ? 'contestado' : 'confirmado';
  const c = { funcId: f.id, ym: ym, ts: ts, acao: acao, motivo: String(p.motivo || '').slice(0, 500), ip: String(p.ip || ''), userAgent: String(p.ua || '').slice(0, 200), deviceId: p.device, hashEspelho: hashEspelho(E) };
  c.hash = sha256([c.funcId, c.ym, c.ts, c.acao, c.motivo, c.ip, c.deviceId, c.hashEspelho].join('|'));
  // Termo em PDF na pasta da funcionária
  const quando = Utilities.formatDate(new Date(ts), TZ, "dd/MM/yyyy 'às' HH:mm:ss");
  const html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial;font-size:11px;line-height:1.5}h1{font-size:16px}td{padding:3px 8px;border-bottom:1px solid #ddd}</style></head><body>' +
    '<h1>Termo de ' + (acao === 'confirmado' ? 'confirmação' : 'contestação') + ' do espelho de ponto — ' + ym + '</h1>' +
    '<p>' + (acao === 'confirmado' ? 'A funcionária abaixo declarou, no aplicativo de ponto, com seu PIN pessoal e em seu celular cadastrado, que conferiu e concorda com as marcações do espelho de ponto do período.' : 'A funcionária abaixo registrou, no aplicativo de ponto, que NÃO concorda com o espelho de ponto do período, pelo motivo informado.') + '</p>' +
    '<table><tr><td>Empregador</td><td>' + (cfg('empregador') || '') + ' · CPF ' + (cfg('cpfEmpregador') || '') + '</td></tr><tr><td>Funcionária</td><td>' + f.nome + (f.cpf ? ' · CPF ' + f.cpf : '') + '</td></tr>' +
    '<tr><td>Período</td><td>' + ym + '</td></tr><tr><td>Totais do espelho</td><td>trabalhado ' + mm(E.tot.trab) + ' · extras ' + mm(E.tot.extra) + ' · débito ' + mm(E.tot.deb) + ' · faltas ' + E.tot.faltas + ' · noturno ' + mm(E.tot.not) + '</td></tr>' +
    (acao === 'contestado' ? '<tr><td>Motivo</td><td>' + c.motivo.replace(/</g, '&lt;') + '</td></tr>' : '') +
    '<tr><td>Data e hora (servidor)</td><td>' + quando + ' (horário de Brasília)</td></tr><tr><td>IP de origem</td><td>' + (c.ip || 'não informado') + '</td></tr>' +
    '<tr><td>Aparelho cadastrado</td><td>' + c.deviceId + '</td></tr><tr><td>Navegador</td><td>' + c.userAgent.replace(/</g, '&lt;') + '</td></tr>' +
    '<tr><td>Código do espelho</td><td>' + c.hashEspelho + '</td></tr><tr><td>Código deste registro</td><td>' + c.hash + '</td></tr></table>' +
    '<p>O código do espelho é o mesmo impresso no PDF do espelho de ponto deste período: se uma marcação mudar, o código muda.</p></body></html>';
  const pdf = Utilities.newBlob(html, 'text/html', 'termo.html').getAs('application/pdf').setName('Termo ' + ym + ' — ' + f.nome + ' (' + acao + ').pdf');
  c.termoId = subpasta(DriveApp.getFolderById(PASTA_ESPELHOS), f.nome).createFile(pdf).getId();
  append('Confirmacoes', c);
  push(['empregador'], acao === 'confirmado' ? 'Ponto — espelho confirmado' : 'Ponto — espelho CONTESTADO', f.nome + (acao === 'confirmado' ? ' confirmou o espelho de ' + ym + '.' : ' contestou o espelho de ' + ym + ': ' + c.motivo));
  return c;
}
function hashEspelho(E) { return sha256(JSON.stringify([E.dias.filter(d => !d.futuro).map(d => [d.key, d.bs.map(b => b.hora + ' ' + b.tipo)]), [E.tot.trab, E.tot.extra, E.tot.deb, E.tot.faltas, E.tot.not]])); }

/* ===================== EMPREGADOR ===================== */
function admin(p) { if (hashPin(p.adminPin) !== cfg('adminPinHash')) throw new Error('PIN do empregador incorreto.'); }
function adminAll(p) {
  admin(p);
  const hoje = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const bs = rows('Batidas').filter(b => (!p.ym || String(b.data).indexOf(p.ym) === 0) && (!p.desde || String(b.data) >= p.desde)).map(b => {
    const x = front(b); if (b.data === hoje && b.selfieId) x.selfie = selfieData(b.selfieId); return x; });
  return {
    admin: { pushAppId: P.getProperty('ONESIGNAL_APP_ID') || '', pushKey: P.getProperty('ONESIGNAL_KEY') ? 'configurada' : '', tolerancia: +cfg('tolerancia') || 0, exigirSelfie: isTrue(cfg('exigirSelfie')), alertaDia: +cfg('alertaDia') || 4, alertaHora: +cfg('alertaHora') || 18, email: String(cfg('email') || ''), empregador: cfg('empregador'), cpfEmpregador: cfg('cpfEmpregador') },
    locais: rows('Locais').map(l => ({ id: l.id, nome: l.nome, endereco: l.endereco, lat: +l.lat, lng: +l.lng, raio: +l.raio })),
    funcs: rows('Funcionarias').map(f => Object.assign(pub(f), { device: f.deviceId || null })),
    batidas: bs,
    confirmacoes: rows('Confirmacoes')
  };
}
function saveFunc(p) {
  admin(p);
  const f = p.func;
  const o = { nome: f.nome, cpf: f.cpf, cargo: f.cargo, localId: f.localId, ativo: !!f.ativo, jornada: JSON.stringify(f.jornada || {}) };
  if (f.pin) { if (!/^\d{4,6}$/.test(f.pin)) throw new Error('O PIN precisa ter de 4 a 6 números.'); o.pinHash = hashPin(f.pin); }
  if (f.id && rows('Funcionarias').some(x => x.id === f.id)) updateById('Funcionarias', f.id, o);
  else { if (!f.pin) throw new Error('Defina um PIN.'); o.id = 'f' + rnd(6); append('Funcionarias', o); }
  return true;
}
function saveLocal(p) {
  admin(p);
  const l = p.local;
  const o = { nome: l.nome, endereco: l.endereco, lat: +l.lat, lng: +l.lng, raio: +l.raio || 100 };
  if (l.id && rows('Locais').some(x => x.id === l.id)) updateById('Locais', l.id, o);
  else { o.id = 'l' + rnd(6); append('Locais', o); }
  return true;
}
function saveAdmin(p) {
  admin(p);
  const c = p.cfg;
  setCfg('empregador', c.empregador || ''); setCfg('cpfEmpregador', c.cpfEmpregador || ''); setCfg('tolerancia', +c.tolerancia || 0); setCfg('exigirSelfie', !!c.exigirSelfie);
  const mudouAlerta = +c.alertaDia !== +cfg('alertaDia') || +c.alertaHora !== +cfg('alertaHora');
  setCfg('alertaDia', +c.alertaDia || 4); setCfg('alertaHora', +c.alertaHora || 18);
  setCfg('email', String(c.email || ''));
  if (mudouAlerta) instalarAlertaSemanal();
  if (c.newPin) { if (!/^\d{4,8}$/.test(c.newPin)) throw new Error('PIN do empregador: 4 a 8 números.'); setCfg('adminPinHash', hashPin(c.newPin)); }
  return true;
}
function selfieData(id) {
  if (!id) return null;
  const b = DriveApp.getFileById(id).getBlob();
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes());
}

/* ===================== ALERTA SEMANAL DE HORAS ===================== */
// Calcula a semana (seg–dom) de cada funcionária e a previsão até domingo pela jornada contratada.
// Se a previsão passar do contratado, avisa a tempo de compensar (liberar mais cedo).
function instalarAlertaSemanal() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alertaSemanal').forEach(t => ScriptApp.deleteTrigger(t));
  const dias = { 3: ScriptApp.WeekDay.WEDNESDAY, 4: ScriptApp.WeekDay.THURSDAY, 5: ScriptApp.WeekDay.FRIDAY };
  ScriptApp.newTrigger('alertaSemanal').timeBased().onWeekDay(dias[+cfg('alertaDia')] || ScriptApp.WeekDay.THURSDAY).atHour(+cfg('alertaHora') || 18).inTimezone(TZ).create();
}
function addDias(key, n) { const d = new Date(key + 'T12:00:00-03:00'); d.setUTCDate(d.getUTCDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function semana(func, batidas, tol, hojeKey) {
  const dow = new Date(hojeKey + 'T12:00:00-03:00').getUTCDay(); const seg = addDias(hojeKey, -((dow + 6) % 7));
  const agora = toMin(Utilities.formatDate(new Date(), TZ, 'HH:mm'));
  let trab = 0, contratado = 0, restante = 0; const diasLongos = [];
  for (let i = 0; i < 7; i++) {
    const key = addDias(seg, i); const d = new Date(key + 'T12:00:00-03:00').getUTCDay();
    const jj = func.jornada && func.jornada[d]; const esp = esperado(jj && jj[0] ? jj : null); contratado += esp;
    const bs = batidas.filter(b => b.data === key && b.status !== 'recusada').sort((a, b) => a.ts - b.ts);
    let t = 0; for (let k = 0; k + 1 < bs.length; k += 2) { let a = toMin(bs[k].hora), b = toMin(bs[k + 1].hora); if (b < a) b += 1440; t += b - a; }
    if (bs.length % 2 === 1 && key === hojeKey) t += Math.max(0, agora - toMin(bs[bs.length - 1].hora));
    if (key <= hojeKey) trab += t;
    if (key > hojeKey) restante += esp; else if (key === hojeKey) restante += Math.max(0, esp - t);
    if (esp > 0 && t > esp + 120) diasLongos.push(key);
  }
  const previsto = trab + restante; let estouro = previsto - contratado; if (Math.abs(estouro) <= tol) estouro = 0;
  return { seg: seg, fim: addDias(seg, 6), trab: trab, contratado: contratado, previsto: previsto, estouro: estouro, diasLongos: diasLongos };
}
function alertaSemanal() {
  const hoje = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); const tol = +cfg('tolerancia') || 0;
  const desde = addDias(hoje, -7); const todas = rows('Batidas').filter(b => String(b.data) >= desde).map(front);
  let algumEstouro = false, s0 = null;
  const linhas = rows('Funcionarias').filter(f => isTrue(f.ativo)).map(f => {
    const func = pub(f); const s = semana(func, todas.filter(b => b.funcId === f.id), tol, hoje); s0 = s0 || s;
    let t = '• ' + f.nome + ': ' + mm(s.trab) + ' trabalhadas de ' + mm(s.contratado) + ' contratadas. Previsão até domingo: ' + mm(s.previsto) + '.';
    if (s.estouro > 0) { algumEstouro = true; t += ' ⚠️ Estouro previsto de ' + mm(s.estouro) + ' — compense liberando mais cedo nos próximos dias.'; } else t += ' Dentro do contratado.';
    if (s.diasLongos.length) t += ' Atenção: ' + s.diasLongos.length + ' dia(s) com mais de 2h extras (limite legal diário).';
    return t; });
  if (!s0) return;
  const pend = rows('Batidas').filter(b => b.status === 'pendente').length;
  const msg = 'Ponto — semana de ' + s0.seg.split('-').reverse().slice(0, 2).join('/') + ' a ' + s0.fim.split('-').reverse().slice(0, 2).join('/') + '\n\n' + linhas.join('\n') + (pend ? '\n\n' + pend + ' batida(s) com localização aguardando sua aprovação.' : '');
  avisar((algumEstouro ? '⚠️ ' : '') + 'Ponto — alerta semanal de horas', msg);
}
function avisar(assunto, msg) {
  push(['empregador'], assunto, msg.length > 400 ? msg.slice(0, 397) + '…' : msg);
  MailApp.sendEmail(String(cfg('email') || '') || Session.getEffectiveUser().getEmail(), assunto, msg);
}

/* ===================== NOTIFICAÇÕES (OneSignal) ===================== */
// ids: 'empregador' ou o id da funcionária (o app faz login no OneSignal com esse id).
function configurarNotificacoes(appId, restKey) { P.setProperties({ ONESIGNAL_APP_ID: appId, ONESIGNAL_KEY: restKey }); Logger.log('Notificações configuradas.'); }
function push(ids, titulo, texto) {
  const appId = P.getProperty('ONESIGNAL_APP_ID'), key = P.getProperty('ONESIGNAL_KEY');
  if (!appId || !key) return;
  try {
    UrlFetchApp.fetch('https://api.onesignal.com/notifications', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Key ' + key },
      payload: JSON.stringify({ app_id: appId, target_channel: 'push', include_aliases: { external_id: ids }, headings: { en: titulo, pt: titulo }, contents: { en: texto, pt: texto } }) });
  } catch (e) { Logger.log('push falhou: ' + e.message); }
}
// Roda a cada 15 min: lembra a funcionária que esqueceu a entrada ou a saída (uma vez por dia cada aviso).
function lembretes() {
  const agora = new Date(); const hoje = Utilities.formatDate(agora, TZ, 'yyyy-MM-dd'); const min = toMin(Utilities.formatDate(agora, TZ, 'HH:mm'));
  const dow = +Utilities.formatDate(agora, TZ, 'u') % 7; // 1=seg … 7=dom → 0=dom
  let log = {}; try { log = JSON.parse(P.getProperty('LEMBRETES') || '{}'); } catch (e) {}
  if (log.data !== hoje) log = { data: hoje, enviados: [] };
  const bs = rows('Batidas').filter(b => String(b.data) === hoje && b.status !== 'recusada');
  rows('Funcionarias').filter(f => isTrue(f.ativo)).forEach(f => {
    const func = pub(f); const j = func.jornada && func.jornada[dow]; if (!j || !j[0]) return;
    const meus = bs.filter(b => b.funcId === f.id); const nome = String(f.nome).split(' ')[0];
    const ent = toMin(j[0]), sai = toMin(j[3] || j[1]);
    const enviar = (tipo, titulo, texto) => { const k = f.id + ':' + tipo; if (log.enviados.indexOf(k) >= 0) return; push([f.id], titulo, texto); log.enviados.push(k); };
    if (!meus.length && min >= ent + 15 && min < sai) enviar('entrada', 'Ponto', nome + ', você ainda não bateu a entrada de hoje.');
    if (meus.length % 2 === 1 && min >= sai + 15) enviar('saida', 'Ponto', nome + ', esqueceu de bater a saída?');
  });
  P.setProperty('LEMBRETES', JSON.stringify(log));
}

/* ===================== FECHAMENTO DO MÊS ===================== */
// Roda sozinho todo dia 1 às 6h (gatilho criado no setup). Pode rodar à mão também.
function fecharMesAnterior() {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
  fecharMes(Utilities.formatDate(d, TZ, 'yyyy-MM'));
}
function fecharMes(ym) {
  const pastaDaeMes = subpasta(DriveApp.getFolderById(PASTA_DAE), ym);
  const tol = +cfg('tolerancia') || 0;
  const anexos = [], linhasEmail = [];
  rows('Funcionarias').filter(f => isTrue(f.ativo)).forEach(f => {
    const func = Object.assign(pub(f));
    const E = espelho(func, batidasDe(f.id, ym), ym, tol);
    const nome = 'Espelho ' + ym + ' — ' + f.nome;
    const pasta = subpasta(DriveApp.getFolderById(PASTA_ESPELHOS), f.nome);
    subpasta(pastaDaeMes, f.nome); // DAE › AAAA-MM › funcionária (recibos/holerites do mês)
    // PDF
    const pdf = Utilities.newBlob(htmlEspelho(func, E, ym), 'text/html', nome + '.html').getAs('application/pdf').setName(nome + '.pdf');
    pasta.createFile(pdf); anexos.push(pdf);
    // Excel
    const ss = SpreadsheetApp.create(nome);
    const sh = ss.getSheets()[0]; const L = linhasPlanilha(E);
    sh.getRange(1, 1, L.length, L[0].length).setValues(L.map(r => r.concat(Array(L[0].length - r.length).fill(''))));
    SpreadsheetApp.flush();
    const xlsx = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(nome + '.xlsx');
    pasta.createFile(xlsx); anexos.push(xlsx);
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    linhasEmail.push(f.nome + ': trabalhou ' + mm(E.tot.trab) + ' · extras ' + mm(E.tot.extra) + ' · débito ' + mm(E.tot.deb) + ' · faltas ' + E.tot.faltas + ' · noturno ' + mm(E.tot.not) + (E.tot.alertas ? ' · ' + E.tot.alertas + ' dia(s) com alerta' : ''));
  });
  rows('Funcionarias').filter(f => isTrue(f.ativo)).forEach(f => push([f.id], 'Ponto — confira seu mês', 'Seu espelho de ' + ym + ' está pronto. Abra o app em "Meu mês" e confirme.'));
  push(['empregador'], 'Ponto — fechamento de ' + ym, linhasEmail.join(' | ').slice(0, 380));
  MailApp.sendEmail({ to: String(cfg('email') || '') || Session.getEffectiveUser().getEmail(), subject: 'Ponto — fechamento de ' + ym,
    body: 'Resumo do mês ' + ym + ':\n\n' + linhasEmail.join('\n') + '\n\nEspelhos em PDF e Excel anexos e salvos em eSocial › Espelhos de ponto, na pasta de cada funcionária. A pasta do DAE do mês já foi criada em eSocial › DAE › ' + ym + '.\nUse esses números para lançar horas extras, adicional noturno e faltas na folha do eSocial.',
    attachments: anexos });
}
function linhasPlanilha(E) {
  const L = [['Data','Dia','Marcações','Trabalhado','Previsto','Extra','Débito','Noturno','Obs.']];
  E.dias.forEach(d => L.push([d.key.split('-').reverse().join('/'), DIAS[d.dow], d.bs.map(b => b.hora).join(' '), mm(d.trab), mm(d.esp), mm(d.extra), mm(d.deb), mm(d.not), d.obs.join(', ')]));
  L.push([]); L.push(['Totais','','',mm(E.tot.trab),mm(E.tot.esp),mm(E.tot.extra),mm(E.tot.deb),mm(E.tot.not),E.tot.faltas + ' falta(s)']);
  return L;
}
function htmlEspelho(f, E, ym) {
  const conf = rows('Confirmacoes').filter(c => c.funcId === f.id && c.ym === ym && c.acao === 'confirmado').pop();
  const tr = E.dias.map(d => '<tr><td>' + d.key.split('-').reverse().join('/') + ' ' + DIAS[d.dow] + '</td><td>' + (d.bs.map(b => b.hora).join(' · ') || (d.j ? '—' : 'folga')) + '</td><td class=n>' + (d.trab ? mm(d.trab) : '') + '</td><td class=n>' + (d.extra ? mm(d.extra) : '') + '</td><td class=n>' + (d.deb ? mm(d.deb) : '') + '</td><td>' + d.obs.join(', ') + '</td></tr>').join('');
  return '<html><head><meta charset="utf-8"><style>body{font-family:Arial;font-size:10px}h1{font-size:16px}table{width:100%;border-collapse:collapse}td,th{border-bottom:1px solid #ccc;padding:3px;text-align:left}.n{text-align:right}</style></head><body>' +
    '<h1>Espelho de ponto — ' + f.nome + '</h1><p>Empregador: ' + (cfg('empregador') || '') + ' · CPF ' + (cfg('cpfEmpregador') || '') + '<br>Funcionária: ' + f.nome + (f.cpf ? ' · CPF ' + f.cpf : '') + ' · ' + (f.cargo || '') + '<br>Período: ' + ym + '<br>Código do espelho: ' + hashEspelho(E) + '</p>' +
    '<p><b>Trabalhado:</b> ' + mm(E.tot.trab) + ' · <b>Extras:</b> ' + mm(E.tot.extra) + ' · <b>Atrasos/saídas:</b> ' + mm(E.tot.deb) + ' · <b>Faltas:</b> ' + E.tot.faltas + ' · <b>Noturno:</b> ' + mm(E.tot.not) + ' (reduzida ' + mm(E.tot.notReduz) + ')</p>' +
    '<table><tr><th>Dia</th><th>Marcações</th><th class=n>Trab.</th><th class=n>Extra</th><th class=n>Débito</th><th>Obs.</th></tr>' + tr + '</table>' +
    '<p>' + (conf ? 'Confirmado pela funcionária no app em ' + Utilities.formatDate(new Date(+conf.ts), TZ, 'dd/MM/yyyy HH:mm') + ' · código ' + String(conf.hash).slice(0, 16) : 'Aguardando confirmação da funcionária.') + '</p>' +
    '<br><br><p>_______________________________ &nbsp;&nbsp;&nbsp; _______________________________<br>Empregador &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Funcionária</p></body></html>';
}

function subpasta(pai, nome) { const it = pai.getFoldersByName(nome); return it.hasNext() ? it.next() : pai.createFolder(nome); }

/* ===================== CÁLCULO (igual ao app) ===================== */
const DIAS = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
function toMin(s) { if (!s) return null; const a = String(s).split(':'); return (+a[0]) * 60 + (+a[1]); }
function mm(m) { const s = m < 0 ? '-' : ''; m = Math.abs(Math.round(m)); return s + Math.floor(m / 60) + ':' + ('0' + (m % 60)).slice(-2); }
function esperado(j) { if (!j) return 0; const e = toMin(j[0]), si = toMin(j[1]), vi = toMin(j[2]), s = toMin(j[3]); if (si != null && vi != null && s != null) return (si - e) + (s - vi); if (si != null && s == null) return si - e; return s - e; }
function noturno(a, b) { let t = 0; for (let d = -1440; d <= 1440; d += 1440) t += Math.max(0, Math.min(b, d + 1740) - Math.max(a, d + 1320)); return t; }
function espelho(func, batidas, ym, tol) {
  const Y = +ym.slice(0, 4), M = +ym.slice(5, 7); const ult = new Date(Y, M, 0).getDate();
  const hoje = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const por = {}; batidas.filter(b => b.status !== 'recusada').forEach(b => (por[b.data] = por[b.data] || []).push(b));
  const dias = [], tot = { trab: 0, esp: 0, extra: 0, deb: 0, not: 0, faltas: 0, alertas: 0 };
  for (let d = 1; d <= ult; d++) {
    const key = ym + '-' + ('0' + d).slice(-2); const dow = new Date(Y, M - 1, d).getDay();
    const jj = func.jornada && func.jornada[dow]; const j = jj && jj[0] ? jj : null; const esp = esperado(j);
    const bs = (por[key] || []).slice().sort((a, b) => a.ts - b.ts); const mins = bs.map(b => toMin(b.hora));
    let trab = 0, not = 0, interv = null; const obs = [];
    for (let i = 0; i + 1 < bs.length; i += 2) { let a = mins[i], b = mins[i + 1]; if (b < a) b += 1440; trab += b - a; not += noturno(a, b); }
    if (bs.length >= 3) interv = mins[2] - mins[1];
    const futuro = key > hoje;
    if (bs.length % 2 === 1 && key < hoje) obs.push('marcação ímpar');
    if (trab > 360 && (interv == null || interv < 60)) obs.push('intervalo < 1h');
    if (interv != null && interv > 120) obs.push('intervalo > 2h');
    let saldo = trab - esp; if (Math.abs(saldo) <= tol) saldo = 0;
    let extra = 0, deb = 0, falta = false;
    if (!futuro && key !== hoje) { if (esp > 0 && !bs.length) { falta = true; deb = esp; obs.push('falta'); } else { extra = Math.max(0, saldo); deb = Math.max(0, -saldo); } }
    if (!j && bs.length) obs.push('dia sem jornada prevista');
    if (bs.some(b => b.status === 'pendente')) obs.push('localização a aprovar');
    if (obs.length) tot.alertas++;
    tot.trab += trab; if (!futuro) tot.esp += esp; tot.extra += extra; tot.deb += deb; tot.not += not; if (falta) tot.faltas++;
    dias.push({ key: key, d: d, dow: dow, j: j, esp: esp, bs: bs, trab: trab, extra: extra, deb: deb, not: not, obs: obs, futuro: futuro });
  }
  tot.notReduz = Math.round(tot.not * 60 / 52.5);
  return { dias: dias, tot: tot };
}

/* ===================== PLANILHA ===================== */
// Planilha aberta e lida uma vez por requisição (antes relia a cada consulta)
let SS_ = null; const MEMO_ = {};
function ss() { return SS_ || (SS_ = SpreadsheetApp.openById(P.getProperty('SHEET_ID'))); }
function rows(nome) {
  if (MEMO_[nome]) return MEMO_[nome];
  const v = ss().getSheetByName(nome).getDataRange().getValues(); const h = v.shift();
  return MEMO_[nome] = v.filter(r => r.join('') !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], TZ, 'yyyy-MM-dd') : r[i]); return o; });
}
function append(nome, o) { delete MEMO_[nome]; ss().getSheetByName(nome).appendRow(COLS[nome].map(k => o[k] === undefined ? '' : o[k])); }
function updateById(nome, id, o) {
  delete MEMO_[nome];
  const sh = ss().getSheetByName(nome); const v = sh.getDataRange().getValues(); const h = v[0];
  for (let r = 1; r < v.length; r++) if (v[r][0] === id) { Object.keys(o).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(r + 1, c + 1).setValue(o[k]); }); return; }
  throw new Error('Registro não encontrado.');
}
function cfg(k) { const r = rows('Config').find(x => x.chave === k); return r ? r.valor : ''; }
function setCfg(k, v) { delete MEMO_['Config']; const sh = ss().getSheetByName('Config'); const vals = sh.getDataRange().getValues();
  for (let r = 1; r < vals.length; r++) if (vals[r][0] === k) { sh.getRange(r + 1, 2).setValue(v); return; } sh.appendRow([k, v]); }
function batidasDe(funcId, ym) { return rows('Batidas').filter(b => b.funcId === funcId && (!ym || String(b.data).indexOf(ym) === 0)).map(front); }
function front(b) { return { id: b.id, funcId: b.funcId, nome: b.nome, data: String(b.data), hora: String(b.hora).length > 5 ? Utilities.formatDate(new Date(b.hora), TZ, 'HH:mm') : String(b.hora), ts: +b.ts, tipo: b.tipo, localId: b.localId, lat: +b.lat, lng: +b.lng, acc: +b.precisao, dist: +b.distancia, selfieId: b.selfieId, hash: b.hash, status: b.status || 'ok' }; }
function pub(f) { let j = {}; try { j = JSON.parse(f.jornada || '{}'); } catch (e) {} return { id: f.id, nome: f.nome, cpf: String(f.cpf || ''), cargo: f.cargo, localId: f.localId, ativo: isTrue(f.ativo), jornada: j }; }

/* ===================== UTIL ===================== */
function isTrue(v) { return v === true || v === 'TRUE' || v === 'true'; }
function rnd(n) { let s = ''; while (s.length < n) s += Utilities.getUuid().replace(/-/g, ''); return s.slice(0, n); }
function sha256(t) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8).map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join(''); }
function hashPin(pin) { return sha256(P.getProperty('SALT') + '|' + String(pin || '')); }
function distM(a, b, c, d) { const R = 6371000, r = x => x * Math.PI / 180; const h = Math.pow(Math.sin(r(c - a) / 2), 2) + Math.cos(r(a)) * Math.cos(r(c)) * Math.pow(Math.sin(r(d - b) / 2), 2); return 2 * R * Math.asin(Math.sqrt(h)); }
