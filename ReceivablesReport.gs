/*************************************************
 * ReceivablesReport.gs  (merged script)
 * Timezone: Asia/Dubai
 * Trigger: Tuesdays ~07:05  (runs runReceivablesReport)
 * Menu: Automations → Receivables Report → Run now (Refresh & Export PDF)
 *************************************************/

const TIMEZONE = 'Asia/Dubai';

/** Presentations **/
const PRESO_MAIN_ID   = '1Uy2wFrhmZ-3lZSrpQXC4OqxIr-qQW3Z3X46IVg1-A70'; // slide 1 to export
const PRESO_OTHER_ID  = '1o0I3KulEQ29rYlKZaOt1VW20DuGZOebo2rMcZBaIn2g'; // refresh slide 2

/** Drive folder for output **/
const EXPORT_FOLDER_ID = '1PG2HNHBdrSZ4Rjdt4CPJSgOFTN7HvajG';

/** MENU **/
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    const receivablesSubmenu = ui.createMenu('Receivables Report')
      .addItem('Run now (Refresh & Export PDF)', 'runReceivablesReport');

    ui.createMenu('Automations')
      .addSubMenu(receivablesSubmenu)
      .addToUi();

    ensureTuesdayTriggerInstalled_();
  } catch (err) {
    Logger.log('onOpen error: ' + err);
  }
}

/** Ensure only the Tuesday 07:05 trigger exists for the merged script **/
function ensureTuesdayTriggerInstalled_() {
  const handler = 'runReceivablesReport';
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler);
  if (existing.length === 0) {
    ScriptApp.newTrigger(handler)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.TUESDAY)
      .atHour(7)
      .nearMinute(5)  // ~07:05 best-effort
      .inTimezone(TIMEZONE)
      .create();
  } else {
    // (optional) ensure timezone alignment
    // No-op: Google doesn’t expose trigger mutation; remove/recreate if you ever need to change it.
  }
}

/** ================== MERGED SCRIPT (Script2) ==================
 * Order:
 *  1) Refresh linked objects on slide 2 of PRESO_OTHER_ID  (charts only)
 *  2) Refresh linked objects on slide 1 of PRESO_MAIN_ID   (charts only)
 *  3) Export slide 1 of PRESO_MAIN_ID as high-res PDF to folder with DD-MMM-YY.pdf
 *  4) (ADDED) Upload the exported PDF to Slack #2-receivables
 */
function runReceivablesReport() {
  // 1) Refresh charts on slide 2 of the other deck
  refreshChartsOnSlideIndex_(PRESO_OTHER_ID, 1); // slide index 1 = second slide

  // 2) Refresh charts on slide 1 of the main deck (the one we export)
  refreshChartsOnSlideIndex_(PRESO_MAIN_ID, 0);  // slide index 0 = first slide

  // 3) Export slide 1 of main deck as vector-sharp PDF → save to Drive
  //    (CHANGED) export function now returns the created fileId
  const fileId = exportFirstSlideToPdf_(PRESO_MAIN_ID, EXPORT_FOLDER_ID);

  // 4) (ADDED) Send the exported PDF to Slack via App method
  try {
    sendPdfToSlackByFileId_(fileId);
  } catch (e) {
    Logger.log('Slack delivery failed: ' + e);
  }
}

/**
 * Refresh all Sheets-linked CHARTS on a given slide (tables cannot be auto-refreshed via API).
 * Requires: Advanced Google services → Slides API enabled.
 */
function refreshChartsOnSlideIndex_(presentationId, slideIndex) {
  const pres = Slides.Presentations.get(presentationId);
  if (!pres.slides || pres.slides.length <= slideIndex) {
    throw new Error(`Slide index ${slideIndex} not found in presentation ${presentationId}`);
  }
  const slide = pres.slides[slideIndex];

  const chartIds = [];
  (slide.pageElements || []).forEach(pe => {
    if (pe.sheetsChart && pe.objectId) chartIds.push(pe.objectId);
  });

  if (chartIds.length === 0) return; // nothing to refresh

  const requests = chartIds.map(id => ({ refreshSheetsChart: { objectId: id } }));
  Slides.Presentations.batchUpdate({ requests }, presentationId);
}

/** Export slide 1 as single-page PDF named DD-MMM-YY.pdf into target folder (vector)
 *  (CHANGED) Returns the created Drive fileId.
 */
function exportFirstSlideToPdf_(presentationId, folderId) {
  const pres = Slides.Presentations.get(presentationId);
  if (!pres.slides || pres.slides.length === 0) throw new Error('No slides to export.');
  const firstSlideId = pres.slides[0].objectId;

  const url = 'https://docs.google.com/presentation/d/' +
              presentationId +
              '/export/pdf?pageid=' + encodeURIComponent(firstSlideId);

  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('PDF export failed: HTTP ' + resp.getResponseCode() + ' — ' + resp.getContentText().slice(0, 200));
  }

  const filename = Utilities.formatDate(new Date(), TIMEZONE, 'dd-MMM-yy') + '.pdf';
  const pdfBlob = resp.getBlob().setName(filename); // application/pdf
  const file = DriveApp.getFolderById(folderId).createFile(pdfBlob);
  return file.getId(); // <<< return new file's ID
}

/** ================== ADDED: Slack upload (App method) ==================
 * Requires Script properties:
 *   SLACK_BOT_TOKEN  = xoxb-...
 *   SLACK_CHANNEL_ID = CXXXXXXXX  (for #2-receivables)
 * Posts with initial comment: "Weekly Receivables Report"
 * NOTE: Set the bot’s display name in Slack to "Uqudo Finance" for the sender name.
 */
function sendPdfToSlackByFileId_(fileId) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const channelId = props.getProperty('SLACK_CHANNEL_ID');
  if (!token || !channelId) throw new Error('Set SLACK_BOT_TOKEN and SLACK_CHANNEL_ID in Script Properties.');

  const file = DriveApp.getFileById(fileId);
  const filename = file.getName();
  const length = file.getSize();
  const blob = file.getBlob().setContentType('application/pdf');

  // 1) Get an upload URL + file_id
  const getUrlResp = UrlFetchApp.fetch('https://slack.com/api/files.getUploadURLExternal', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token },
    contentType: 'application/x-www-form-urlencoded; charset=utf-8',
    payload: { filename: filename, length: String(length) },
    muteHttpExceptions: true
  });
  const getUrlJson = JSON.parse(getUrlResp.getContentText());
  if (!getUrlJson.ok) throw new Error('Slack getUploadURLExternal failed: ' + getUrlResp.getContentText());
  const uploadUrl = getUrlJson.upload_url;
  const slackFileId = getUrlJson.file_id;

  // 2) Upload raw bytes to the returned URL
  const uploadResp = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    contentType: 'application/octet-stream',
    payload: blob.getBytes(),
    muteHttpExceptions: true
  });
  if (uploadResp.getResponseCode() !== 200) {
    throw new Error('Raw upload to Slack failed. HTTP ' + uploadResp.getResponseCode() + ': ' + uploadResp.getContentText());
  }

  // 3) Complete & share to channel with initial comment
  const completeBody = {
    channel_id: channelId,
    initial_comment: 'Weekly Receivables Report',
    files: [{ id: slackFileId, title: filename }]
  };
  const completeResp = UrlFetchApp.fetch('https://slack.com/api/files.completeUploadExternal', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token },
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(completeBody),
    muteHttpExceptions: true
  });
  const completeJson = JSON.parse(completeResp.getContentText());
  if (!completeJson.ok) throw new Error('Slack completeUploadExternal failed: ' + completeResp.getContentText());
}
