// ════════════════════════════════════════════════════════════
//  SAT Doctor — Private Sheet Form Trigger
//  Paste this into the Apps Script bound to your PRIVATE sheet
//  (open private sheet → Extensions → Apps Script).
//
//  Then add a trigger:
//    Triggers → + Add Trigger
//    Function:     onFormSubmit
//    Event source: From spreadsheet
//    Event type:   On form submit
// ════════════════════════════════════════════════════════════

// ← Paste your main Apps Script /exec URL here
var MAIN_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyTFXAiC81l9POiK7sKQ0qHU__ErfMF9XHx7iDJhtfEmklfklf42dayO9eIOfnSR1E-/exec';

function onFormSubmit(e) {
  try {
    var r = e.namedValues || {};

    var params = {
      action:    'book',
      name:      (r['Name']    || r['name']    || [''])[0].trim(),
      email:     (r['Email']   || r['email']   || [''])[0].trim(),
      phone:     (r['Phone']   || r['phone']   || [''])[0].trim(),
      course:    (r['Course']  || r['course']  || [''])[0].trim(),
      slotId:    (r['Slot ID'] || r['slot_id'] || r['SlotID'] || [''])[0].trim(),
      requestId: 'form-' + Date.now()
    };

    // Log what we received so you can debug from the private sheet's script
    Logger.log('onFormSubmit received: ' + JSON.stringify(params));

    if (!params.slotId || !params.name || !params.email || !params.phone) {
      Logger.log('Missing required fields. Keys received: ' + Object.keys(r).join(', '));
      return;
    }

    var qs = Object.keys(params).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }).join('&');

    var url = MAIN_APPS_SCRIPT_URL + '?' + qs;
    Logger.log('Calling main script: ' + url);

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Logger.log('Response: ' + response.getContentText());

  } catch(err) {
    Logger.log('onFormSubmit error: ' + String(err.message || err));
  }
}