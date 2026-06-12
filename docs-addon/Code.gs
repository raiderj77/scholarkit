/**
 * ScholarKit for Google Docs. Personal-use build.
 * Paste this whole file into the Apps Script editor as Code.gs.
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('ScholarKit')
    .addItem('Open ScholarKit', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('ScholarKit');
  DocumentApp.getUi().showSidebar(html);
}

/* ---------- document access ---------- */

function getSelectedText() {
  var sel = DocumentApp.getActiveDocument().getSelection();
  if (!sel) return '';
  var out = [];
  sel.getRangeElements().forEach(function (re) {
    var el = re.getElement();
    if (!el.editAsText) return;
    var t = el.asText().getText();
    if (re.isPartial()) {
      out.push(t.substring(re.getStartOffset(), re.getEndOffsetInclusive() + 1));
    } else if (t) {
      out.push(t);
    }
  });
  return out.join('\n');
}

function insertAtCursor(text) {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  if (cursor) {
    var el = cursor.insertText(text);
    if (!el) throw new Error('Could not insert at the cursor. Click into the document body first.');
    return true;
  }
  // No cursor and no selection: append to end of body.
  doc.getBody().appendParagraph(text);
  return true;
}

function replaceSelection(text) {
  var doc = DocumentApp.getActiveDocument();
  var sel = doc.getSelection();
  if (!sel) return insertAtCursor(text);

  var ranges = sel.getRangeElements().filter(function (re) {
    return re.getElement().editAsText;
  });
  if (!ranges.length) return insertAtCursor(text);

  // Delete selected text from last range to first, then insert into the first.
  for (var i = ranges.length - 1; i >= 0; i--) {
    var re = ranges[i];
    var t = re.getElement().asText();
    if (re.isPartial()) {
      t.deleteText(re.getStartOffset(), re.getEndOffsetInclusive());
    } else {
      var full = t.getText();
      if (full.length) t.deleteText(0, full.length - 1);
    }
  }
  var first = ranges[0];
  var startOffset = first.isPartial() ? first.getStartOffset() : 0;
  first.getElement().asText().insertText(startOffset, text);
  return true;
}

function getDocText() {
  return DocumentApp.getActiveDocument().getBody().getText();
}

/** Pull plain text from another Google Doc you own, by URL or ID. */
function getDocTextById(idOrUrl) {
  var m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('That does not look like a Google Doc link or ID.');
  var doc = DocumentApp.openById(m[0]);
  return { name: doc.getName(), text: doc.getBody().getText() };
}

/* ---------- settings ---------- */

function saveSettings(key, model) {
  var p = PropertiesService.getUserProperties();
  p.setProperty('SK_KEY', key || '');
  p.setProperty('SK_MODEL', model || 'gemini-2.5-flash');
  return true;
}

function getSettings() {
  var p = PropertiesService.getUserProperties();
  return {
    hasKey: !!p.getProperty('SK_KEY'),
    model: p.getProperty('SK_MODEL') || 'gemini-2.5-flash'
  };
}

/* ---------- Gemini ---------- */

function callGemini(prompt) {
  var p = PropertiesService.getUserProperties();
  var key = p.getProperty('SK_KEY');
  var model = p.getProperty('SK_MODEL') || 'gemini-2.5-flash';
  if (!key) throw new Error('No API key saved. Open the Settings tab and save your Gemini key.');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, responseMimeType: 'application/json' }
    })
  });

  var code = res.getResponseCode();
  var data = JSON.parse(res.getContentText());
  if (code >= 300) {
    throw new Error((data.error && data.error.message) || ('HTTP ' + code));
  }
  var txt;
  try { txt = data.candidates[0].content.parts[0].text; }
  catch (e) { throw new Error('Empty response from model.'); }
  return txt.replace(/^```(json)?/i, '').replace(/```$/m, '').trim();
}
