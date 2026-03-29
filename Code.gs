// ── CONFIG ────────────────────────────────────────────────
const SHEET_NAME  = 'Sheet1';
const SIGNUP_URL  = 'https://dmichals.github.io/CrewSignup/signup.html';
const REPLY_EMAIL = 'djmichals@gmail.com';
const SITE_NAME   = 'Sailing with Attitude';
// ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const rows  = sheet.getDataRange().getValues();
    const email = data.email.trim().toLowerCase();
    const now   = new Date();
    let editID='', rowIndex=-1;

    for (let i=1; i<rows.length; i++) {
      if (rows[i][2].toString().trim().toLowerCase()===email) {
        rowIndex=i+1; editID=rows[i][7]; break;
      }
    }
    if (!editID) editID=Utilities.getUuid();

    const editLink = SIGNUP_URL
      +'?name='  +encodeURIComponent(data.name)
      +'&email=' +encodeURIComponent(data.email)
      +'&days='  +encodeURIComponent(data.days)
      +'&jobs='  +encodeURIComponent(data.jobs)
      +'&beer='  +encodeURIComponent(data.beer)
      +'&notes=' +encodeURIComponent(data.notes||'')
      +'&editid='+editID;

    const rowData=[now,data.name,data.email,data.days,data.jobs,data.beer,data.notes||'',editID];
    if (rowIndex>0) sheet.getRange(rowIndex,1,1,8).setValues([rowData]);
    else sheet.appendRow(rowData);

    MailApp.sendEmail({
      to: data.email,
      subject: SITE_NAME+' — you\'re signed up for 2025!',
      body: 'Hi '+data.name+',\n\n'
        +'You\'re on the schedule!\n\n'
        +'Days: '+data.days+'\n'
        +'Jobs: '+data.jobs+'\n'
        +(data.beer==='Yes'?'Beer duty: Yes!\n':'')+'\n'
        +'Need to make changes? Use your personal link anytime this season:\n\n'
        +editLink
        +'\n\nSee you on the water!\nDave'
        +'\n\n---\nQuestions? Email Dave at '+REPLY_EMAIL,
      replyTo: REPLY_EMAIL,
    });

    MailApp.sendEmail({
      to: REPLY_EMAIL,
      subject: SITE_NAME+' — new signup: '+data.name,
      body: data.name+' ('+data.email+') just signed up.\n\n'
        +'Days: '+data.days+'\nJobs: '+data.jobs+'\nBeer: '+data.beer+'\nNotes: '+(data.notes||'none'),
    });

    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',msg:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
}
