function doPost(e) {
  if (typeof e !== 'undefined')
    var role = e.parameter.role;
  var workplace = e.parameter.workplace;
  var team = e.parameter.team;
  var Fname = e.parameter.Fname_001.replace(/^\s+|\s+$/g, '');
  var Lname = e.parameter.Lname_001.replace(/^\s+|\s+$/g, '');
  var email = e.parameter.email_001.replace(/^\s+|\s+$/g, '');
  var phone = e.parameter.phone_001.replace(/^\s+|\s+$/g, '');
  var leavingdate = e.parameter.leavingdate;
  var startingdate = e.parameter.startingdate;
  var addrline1 = e.parameter.addrline1_001.replace(/^\s+|\s+$/g, '');
  var addrline2 = e.parameter.addrline2_001.replace(/^\s+|\s+$/g, '');
  var town = e.parameter.town_001.replace(/^\s+|\s+$/g, '');
  var county = e.parameter.county_001.replace(/^\s+|\s+$/g, '');
  var postcode = e.parameter.postcode_001.replace(/^\s+|\s+$/g, '');
  var reasonforleaving = e.parameter.reasonforleaving;
  var outstandingexpenses = e.parameter.outstandingexpenses;
  var charityequipment = e.parameter.charityequipment;
  var days_in_post = e.parameter.days_in_post;

  sendLetter(role, workplace, team, Fname, Lname, email, phone, startingdate, leavingdate, addrline1, addrline2, town, county, postcode, reasonforleaving, outstandingexpenses, charityequipment, days_in_post);
  var html = HtmlService.createTemplateFromFile('Thankyou');
  return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/*  Get the URL for the Google Apps Script running as a WebApp. */
function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  return url;
}

function doGet(e) {
  var currentuser = Session.getActiveUser().getEmail();
  var authusers = ["sboocock@calancs.org.uk", "aoshea@calancs.org.uk", "gsimpson@calancs.org.uk", "mdeslandes@calancs.org.uk", "myates@calancs.org.uk",
    "jleadbetter@calancs.org.uk", "icook@calancs.org.u", "ccatterall@calancs.org.uk", "mdeslandes@calancs.org.uk",
    "sbennett@calancs.org.uk", "admin@calancs.org.uk", "gsimpson@calancs.org.uk", "mdeslandes@calancs.org.uk",
    "pfurner@calancs.org.uk", "amilligan@calancs.org.uk", "nhigham@calancs.org.uk", "swalsh@calancs.org.uk", "jasbury@calancs.org.uk", "pirlam@calancs.org.uk"
  ];
  var isauthorised = authusers.indexOf(currentuser);
  if (isauthorised >= 0) {
    return HtmlService.createHtmlOutputFromFile('Form');
  } else {
    var html = HtmlService.createTemplateFromFile('Error');
    html.currentuser = currentuser;
    return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    //return HtmlService.createHtmlOutputFromFile('Error');
  }
}