

// All the columns are counting start from 0 instead of 1 
var CONFIGURATION = {
  SPREAD_SHEET_URL: "URL to Google SpreadSheet",
  SCHEDULE_SHEET: {
    RANGE: "!A:A",         // select range of Schedule Sheet
    DATE_COLUMN: 0,        // which column is for Date
    SERMON_COLUMN: 0,      // which column is for Sermon Person
    WORSHIP_COLUMN: 0      // which column is for Worship Person
  },
  CONTACT_SHEET: {
    RANGE: "!A:A",         // select range of Contact Sheet
    NAME_COLUMN: 0,        // which column is for name
    REFER_NAME_COLUMN: 0,  // which column is for refer_name
    EMAIL_COLUMN: 0,       // which column is for email
    IS_ADMIN_COLUMN: 0     // which column is for is_admin
  },
  EMAIL_SUBJECT: "Email Subject for all emails"
}

function autonoti() {

  // init check
  for (key in CONFIGURATION) {
    if (typeof CONFIGURATION[key] == 'string' && CONFIGURATION[key].length <= 0)
      throw Error("Missing configuration constants.");
  }

  var today = new Date();
  Logger.log(`Start Apps Script At: [${getDateString(today)}]`)

  // init SpreadsheetApp
  var spapp = SpreadsheetApp.openByUrl(CONFIGURATION['SPREAD_SHEET_URL']);
  if (!spapp)
    throw Error("Cannot open spreadsheet url.");

  // get admin info
  var admins = getAdminInfo(spapp);
  CONFIGURATION['ADMINS'] = admins;
  var tmpstr = "Admins: [";
  for (a of admins)
    tmpstr += a[CONFIGURATION['CONTACT_SHEET']['NAME_COLUMN']] + ", ";
  tmpstr = tmpstr.substr(0, tmpstr.length-2);
  tmpstr += "]";
  Logger.log(tmpstr);

  // get current sermon_info & worship_info
  var cur_ppl = getWklyPeople(spapp);
  if (!cur_ppl)
    throw Error("Cannot find weekly people.");
  var sermon_info = getContactInfo(spapp, cur_ppl[1]);
  var worship_info = getContactInfo(spapp, cur_ppl[2]);
  Logger.log(`sermon_info = ${sermon_info}`);
  Logger.log(`worship_info = ${worship_info}`);

  // generate sermon & worship msg
  var msg_tmplt_url = spapp.getSheetByName("MessagesTemplate").getRange("!A1:B2").getValues();
  var sermon_msg = null;
  var worship_msg = null;
  if (sermon_info) {
    var sermon_url = msg_tmplt_url[0][1];
    var sermon_id = DocumentApp.openByUrl(sermon_url).getId();
    sermon_msg = parseMessage(getHtmlByDocId(sermon_id), sermon_info[1], worship_info[1]);
  }
  if (worship_info) {
    var worship_url = msg_tmplt_url[1][1];
    var worship_id = DocumentApp.openByUrl(worship_url).getId();
    worship_msg = parseMessage(getHtmlByDocId(worship_id), sermon_info[1], worship_info[1]);
  }

  // send email to all

  // sermon
  if (sermon_info && sermon_msg) {
    sendEmail(
      sermon_info[2],
      CONFIGURATION['EMAIL_SUBJECT'],
      sermon_msg
    );
  }
  
  // worship
  if (worship_info && worship_msg) {
    sendEmail(
      worship_info[2],
      CONFIGURATION['EMAIL_SUBJECT'],
      worship_msg
    );
  }

  // admin
  for (a of admins) {
    sendEmail(
      a[2],
      CONFIGURATION['EMAIL_SUBJECT'],
      "sermon msg:<br>" + sermon_msg + "<br>worship msg<br>" + worship_msg
    );
  }

}

function getDateString(date) {
    return date.toLocaleDateString(
      'en-US',
      {
        year:'numeric',
        month:'numeric',
        day:'numeric',
        hour:'numeric',
        minute:'numeric',
        second:'numeric',
        timeZone:'America/Los_Angeles',
        timeZoneName :'short'
      }
  );
}

function getHtmlByDocId(id) {
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
  var param = 
  {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
  };
  var html = UrlFetchApp.fetch(url,param).getContentText();
  return html;
}

function parseMessage(msg, sermon_name, worship_name) {
  msg_keywords = {
    'sermon_name': sermon_name,
    'worship_name': worship_name,
    'admin_name': CONFIGURATION['ADMINS'][0][1],
    'admin_email': CONFIGURATION['ADMINS'][0][2]
  };

  match = [...msg.matchAll(/\${([a-zA-Z_]+)}/g)];
  for (m of match) {
    replacement = msg_keywords[m[1]];
    msg = msg.replace(m[0], replacement);
  }
  return msg;
}

function getWklyPeople(spapp) {
  var today = new Date();
  var sheet = spapp.getSheetByName("Schedule");
  var values = sheet.getRange(CONFIGURATION['SCHEDULE_SHEET']['RANGE']).getValues();
  var date = new Date(0);
  var date_col = CONFIGURATION['SCHEDULE_SHEET']['DATE_COLUMN'];
  var serm_col = CONFIGURATION['SCHEDULE_SHEET']['SERMON_COLUMN'];
  var wors_col = CONFIGURATION['SCHEDULE_SHEET']['WORSHIP_COLUMN'];
  for (v of values) {
    if (typeof v[0] == 'string' && v[0].includes("year=")) {
      date.setYear(v[0].substr(5, 4))
      date.setMonth(1);
      date.setDate(1);
    }
    else if (v[date_col] instanceof Date) {
      date.setMonth(v[date_col].getMonth());
      date.setDate(v[date_col].getDate());
      date.setHours(0);
      date.setMinutes(0);
      date.setSeconds(0);
    }
    if (date >= today) {
      return [date, v[serm_col], v[wors_col]];
    }
  }
  return null;
}

function getContactInfo(spapp, name) {
  var sheet = spapp.getSheetByName("Contact");
  var values = sheet.getRange(CONFIGURATION['CONTACT_SHEET']['RANGE']).getValues();
  for (v of values) {
    if (v[CONFIGURATION['CONTACT_SHEET']['NAME_COLUMN']] == name)
      return v;
  }
  return null;
}

function getAdminInfo(spapp) {
  var sheet = spapp.getSheetByName("Contact");
  var values = sheet.getRange(CONFIGURATION['CONTACT_SHEET']['RANGE']).getValues();
  var output = [];
  for (v of values) {
    if (v[CONFIGURATION['CONTACT_SHEET']['IS_ADMIN_COLUMN']] > 0)
      output.push(v);
  }
  return output;
}

function sendEmail(to_email, email_subject, html_message) {
  Logger.log(`Send Email To ${to_email}, At [${getDateString(new Date())}]`);
  var ma = MailApp;
  ma.sendEmail({
    to: to_email,
    subject: email_subject,
    htmlBody: html_message
  });
  return ma.getRemainingDailyQuota();
}

