/**
 * PONTO DA CASA — carregador do servidor.
 * Este arquivo quase nunca muda: ele busca o código de verdade no GitHub
 * (btmarques/ponto-casa › server/Code.js) e executa. Assim as atualizações
 * chegam sozinhas, sem colar código aqui de novo.
 */
const FONTE_CODIGO = 'https://raw.githubusercontent.com/btmarques/ponto-casa/main/server/Code.js';
const NOMES = ['doGet', 'doPost', 'setup', 'fecharMesAnterior', 'fecharMes', 'alertaSemanal', 'lembretes', 'instalarAlertaSemanal', 'configurarNotificacoes', 'migrar'];
let SRV_ = null;

function servidor_() {
  if (SRV_) return SRV_;
  const cache = CacheService.getScriptCache();
  let codigo = cache.get('codigo');
  if (!codigo) {
    const r = UrlFetchApp.fetch(FONTE_CODIGO + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Não consegui baixar o código do servidor (' + r.getResponseCode() + ').');
    codigo = r.getContentText();
    try { cache.put('codigo', codigo, 300); } catch (e) {}
  }
  const exporta = NOMES.map(n => n + ': typeof ' + n + " === 'function' ? " + n + ' : undefined').join(', ');
  SRV_ = new Function(codigo + '\nreturn {' + exporta + '};')();
  return SRV_;
}

function doGet(e) { return servidor_().doGet(e); }
function doPost(e) { return servidor_().doPost(e); }
function setup() { return servidor_().setup(); }
function fecharMesAnterior() { return servidor_().fecharMesAnterior(); }
function alertaSemanal() { return servidor_().alertaSemanal(); }
function lembretes() { return servidor_().lembretes(); }
function migrar() { const f = servidor_().migrar; return f ? f() : 'nada a migrar'; }
function atualizarAgora() { CacheService.getScriptCache().remove('codigo'); SRV_ = null; servidor_(); Logger.log('Código do servidor recarregado do GitHub.'); }

// Nunca é chamada: só lista os serviços do Google que o código usa, para o Google pedir as permissões certas.
function permissoes_() {
  DriveApp.getRootFolder(); SpreadsheetApp.create('x'); MailApp.sendEmail('', '', '');
  UrlFetchApp.fetch(''); ScriptApp.newTrigger('x'); LockService.getScriptLock();
  Session.getEffectiveUser(); PropertiesService.getScriptProperties(); CacheService.getScriptCache();
}
