const SHEET_ID = '14PWwSw-_vPt1-o--vevqQwZ1IbGb7MUPR4XV__veFkY';

function doGet(e) {
  const tipo = e.parameter.tipo || '';
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Gravar confirmação de presença ──
  if (tipo === 'rsvp') {
    let sheet = ss.getSheetByName('Confirmações');
    if (!sheet) {
      sheet = ss.insertSheet('Confirmações');
      sheet.appendRow(['Data/Hora', 'Nome', 'Confirmação']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    sheet.appendRow([new Date(), e.parameter.nome || '', e.parameter.confirmacao || '']);
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  // ── Gravar mensagem do mural ──
  if (tipo === 'mural_enviar') {
    let sheet = ss.getSheetByName('Mural');
    if (!sheet) {
      sheet = ss.insertSheet('Mural');
      sheet.appendRow(['Data/Hora', 'Nome', 'Mensagem', 'Aprovado']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }
    sheet.appendRow([new Date(), e.parameter.nome || '', e.parameter.mensagem || '', false]);
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  // ── Ler mensagens aprovadas do mural (JSONP) ──
  if (tipo === 'mural') {
    let sheet = ss.getSheetByName('Mural');
    const msgs = [];
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][3] === true) {
          msgs.push({
            data: Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'dd/MM/yyyy'),
            nome: String(data[i][1]),
            mensagem: String(data[i][2])
          });
        }
      }
    }
    const cb = e.parameter.callback || 'muralCallback';
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(msgs) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
