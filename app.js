function doPost(e) {
  if (typeof e !== 'undefined')
  let role = e.parameter.role;
  let workplace = e.parameter.workplace;
  let team = e.parameter.team;
  let Fname = e.parameter.Fname_001.replace(/^\s+|\s+$/g, '');
  let Lname = e.parameter.Lname_001.replace(/^\s+|\s+$/g, '');
  let email = e.parameter.email_001.replace(/^\s+|\s+$/g, '');
  let phone = e.parameter.phone_001.replace(/^\s+|\s+$/g, '');
  let leavingdate = e.parameter.leavingdate;
  let startingdate = e.parameter.startingdate;
  let addrline1 = e.parameter.addrline1_001.replace(/^\s+|\s+$/g, '');
  let addrline2 = e.parameter.addrline2_001.replace(/^\s+|\s+$/g, '');
  let town = e.parameter.town_001.replace(/^\s+|\s+$/g, '');
  let county = e.parameter.county_001.replace(/^\s+|\s+$/g, '');
  let postcode = e.parameter.postcode_001.replace(/^\s+|\s+$/g, '');
  let reasonforleaving = e.parameter.reasonforleaving;
  let outstandingexpenses = e.parameter.outstandingexpenses;
  let charityequipment = e.parameter.charityequipment;
  let days_in_post = e.parameter.days_in_post;

  sendLetter(role, workplace, team, Fname, Lname, email, phone, startingdate, leavingdate, addrline1, addrline2, town, county, postcode, reasonforleaving, outstandingexpenses, charityequipment, days_in_post);
  let html = HtmlService.createTemplateFromFile('Thankyou');
  return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/*  Get the URL for the Google Apps Script running as a WebApp. */
function getScriptUrl() {
  let url = ScriptApp.getService().getUrl();
  return url;
}

function doGet(e) {
  let currentuser = Session.getActiveUser().getEmail();
  let authusers = ["sboocock@calancs.org.uk", "aoshea@calancs.org.uk", "gsimpson@calancs.org.uk", "mdeslandes@calancs.org.uk", "myates@calancs.org.uk",
    "jleadbetter@calancs.org.uk", "icook@calancs.org.u", "ccatterall@calancs.org.uk", "mdeslandes@calancs.org.uk",
    "sbennett@calancs.org.uk", "admin@calancs.org.uk", "gsimpson@calancs.org.uk", "mdeslandes@calancs.org.uk",
    "pfurner@calancs.org.uk", "amilligan@calancs.org.uk", "nhigham@calancs.org.uk", "swalsh@calancs.org.uk", "jasbury@calancs.org.uk", "pirlam@calancs.org.uk"
  ];
  let isauthorised = authusers.indexOf(currentuser);
  if (isauthorised >= 0) {
    return HtmlService.createHtmlOutputFromFile('Form');
  } else {
    let html = HtmlService.createTemplateFromFile('Error');
    html.currentuser = currentuser;
    return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    //return HtmlService.createHtmlOutputFromFile('Error');
  }
}

function ordinal_suffix_of(i) {
  let j = i % 10,
    k = i % 100;
  if (j == 1 && k != 11) {
    return "st";
  }
  if (j == 2 && k != 12) {
    return "nd";
  }
  if (j == 3 && k != 13) {
    return "rd";
  }
  return "th";
}

function sendLetter(role, workplace, team, Fname, Lname, email, phone, startingdate, leavingdate, addrline1, addrline2, town, county, postcode, reasonforleaving, outstandingexpenses, charityequipment, days_in_post) {
  let roletype;
  switch (role) {
    case "Volunteer":
      roletype = "volunteering role";
      break;
    case "Staff":
      roletype = "employment";
      break;
  }
  //Set up destination folder
  let dstFolderId = DriveApp.getFolderById("1lVsjyYWp2VdvSqEF4lgfgMXF6kEK-FcN");
  //Convert date to ISO format (yyyy/mm/dd)
  let startdateArray = startingdate.split("/");
  let leavingdateArray = leavingdate.split("/");
  let startdate = new Date(startdateArray[2] + "/" + startdateArray[1] + "/" + startdateArray[0]);
  let leavingdate = new Date(leavingdateArray[2] + "/" + leavingdateArray[1] + "/" + leavingdateArray[0]);
  let ordinal;
  let day;
  if (leavingdateArray[0].substring(0, 1) == "0") {
    ordinal = ordinal_suffix_of(leavingdateArray[0].substring(1))
    day = leavingdateArray[0].substring(1);
  } else {
    ordinal = ordinal_suffix_of(leavingdateArray[0])
    day = leavingdateArray[0];
  }
  let monthname;
  switch (leavingdateArray[1]) {
    case "01":
      monthname = "January";
      break;
    case "02":
      monthname = "February";
      break;
    case "03":
      monthname = "March";
      break;
    case "04":
      monthname = "April";
      break;
    case "05":
      monthname = "May";
      break;
    case "06":
      monthname = "June";
      break;
    case "07":
      monthname = "July";
      break;
    case "08":
      monthname = "August";
      break;
    case "09":
      monthname = "September";
      break;
    case "10":
      monthname = "October";
      break;
    case "11":
      monthname = "November";
      break;
    case "12":
      monthname = "December";
      break;
  }
  let leavingdateLongFormat = day + ordinal + " " + monthname + " " + leavingdateArray[2];
  //Set up address
  let address;
  if (addrline2.length > 0) {
    if (county.length > 1) {
      address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + county + '\r' + postcode;
    } else {
      address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + postcode;
    }
  } else {
    if (county.length > 1) {
      address = addrline1 + '\r' + town + '\r' + county + '\r' + postcode;
    } else {
      address = addrline1 + "\r" + town + "\r" + postcode;
    }
  }
  //Leavers Letter
  if (days_in_post >= 180) {
    let docid = DriveApp.getFileById("1n3zWEttQA8LiwADEBFH-a5oHzU6ycGBhz0KM3l_d2fs").makeCopy("Leavers_Letter_" + Utilities.formatDate(new Date(), "GMT+1", "dd-MMM-yyyy") + "_" + name, dstFolderId).getId()
    let doc = DocumentApp.openById(docid);
    let body = doc.getActiveSection();
    body.replaceText("%fullname%", Fname + " " + Lname);
    body.replaceText("%addrline1%", address);
    body.replaceText("%town%", town);
    body.replaceText("%county%", county);
    body.replaceText("%postcode%", postcode);
    body.replaceText("%name%", Fname);
    body.replaceText("%leavingdate%", leavingdateLongFormat);
    body.replaceText("%today%", Utilities.formatDate(new Date(), "GMT+1", "dd-MMM-yyyy"));
    body.replaceText("%team%", team);
    body.replaceText("%area%", workplace);
    body.replaceText("%roletype%", roletype);
    doc.saveAndClose();
    //GmailApp.sendEmail('dgradwell@calancs.org.uk', 'Leaver Notification (LETTER)', 'Please see the attached letter.', {bcc:'gsimpson@calancs.org.uk,admin@calancs.org.uk',attachments: [doc.getAs(MimeType.PDF)],name: 'Leaver Notification Letter' });
    GmailApp.sendEmail('jasbury@calancs.org.uk', 'Leaver Notification Letter', 'Please see the attached letter.', {
      attachments: [doc.getAs(MimeType.PDF)],
      name: 'Leaver Notification'
    });
  }
  //Leavers Report
  let docid2 = DriveApp.getFileById("1dm7XshpRE1jsHhUkV45X61-eYA1IjW6ut2H6g1mxhzE").makeCopy("Leavers_Report_" + Utilities.formatDate(new Date(), "GMT+1", "dd-MMM-yyyy") + "_" + name, dstFolderId).getId();
  let doc2 = DocumentApp.openById(docid2);
  if (phone.length < 1) {
    phone = "Not recorded";
  }
  if (reasonforleaving.length < 1) {
    reasonforleaving = "Not recorded";
  }
  if (outstandingexpenses.length < 1) {
    outstandingexpenses = "None stated";
  }
  if (charityequipment.length < 1) {
    charityequipment = "None stated";
  }
  let body2 = doc2.getActiveSection();
  body2.replaceText("%fullname%", Fname + " " + Lname);
  body2.replaceText("%addrline1%", address)
  body2.replaceText("%town%", town);
  body2.replaceText("%county%", county);
  body2.replaceText("%postcode%", postcode);
  body2.replaceText("%name%", Fname);
  body2.replaceText("%startdate%", Utilities.formatDate(startdate, "GMT+1", "dd-MMM-yyyy"));
  body2.replaceText("%leavingdate%", Utilities.formatDate(leavingdate, "GMT+1", "dd-MMM-yyyy"));
  body2.replaceText("%team%", team);
  body2.replaceText("%email%", email);
  body2.replaceText("%phone%", phone);
  body2.replaceText("%role%", role);
  body2.replaceText("%workplace%", workplace);
  body2.replaceText("%reasonforleaving%", reasonforleaving);
  body2.replaceText("%outstandingexpenses%", outstandingexpenses);
  body2.replaceText("%charityequipment%", charityequipment);
  doc2.saveAndClose();
  switch (workplace) {
    case "Leyland":
      GmailApp.sendEmail('admin@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
        cc: 'sboocock@calancs.org.uk,\
            aoshea@calancs.org.uk,gsimpson@calancs.org.uk,mdeslandes@calancs.org.uk,jleadbetter@calancs.org.uk,icook@calancs.org.uk,ccatterall@calancs.org.uk,\
            sbennett@calancs.org.uk',
        attachments: [doc2.getAs(MimeType.PDF)],
        name: 'Leaver Notification Report'
      });
      break;
    case "Chorley":
      GmailApp.sendEmail('admin@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
        cc: 'sboocock@calancs.org.uk,\
            aoshea@calancs.org.uk,gsimpson@calancs.org.uk,mdeslandes@calancs.org.uk,nhigham@calancs.org.uk',
        attachments: [doc2.getAs(MimeType.PDF)],
        name: 'Leaver Notification Report'
      });
      break;
    case "West Lancashire":
      GmailApp.sendEmail('admin@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
        cc: 'sboocock@calancs.org.uk,\
            aoshea@calancs.org.uk,gsimpson@calancs.org.uk,mdeslandes@calancs.org.uk,myates@calancs.org.uk,swalsh@calancs.org.uk',
        attachments: [doc2.getAs(MimeType.PDF)],
        name: 'Leaver Notification Report'
      });
      break;
    case "Wyre":
      //GmailApp.sendEmail('admin@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {cc:'sboocock@calancs.org.uk,\
      //aoshea@calancs.org.uk,gsimpson@calancs.org.uk,mdeslandes@calancs.org.uk,pfurner@calancs.org.uk', attachments: [doc2.getAs(MimeType.PDF)],name: 'Leaver Notification Report' });
      GmailApp.sendEmail('jasbury@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
        attachments: [doc2.getAs(MimeType.PDF)],
        name: 'Leaver Notification Report'
      });
      break;
    case "Blackburn with Darwen":
      GmailApp.sendEmail('admin@calancs.org.uk', 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
        cc: 'sboocock@calancs.org.uk,\
           aoshea@calancs.org.uk,gsimpson@calancs.org.uk,mdeslandes@calancs.org.uk,amilligan@calancs.org.uk',
        attachments: [doc2.getAs(MimeType.PDF)],
        name: 'Leaver Notification Report'
      });
      break;
  }
}
