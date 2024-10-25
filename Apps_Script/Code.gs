function sendEmail() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses 1");
  const lastRow = spreadSheet.getLastRow();
  const date = new Date();

  // PERSONAL DETAILS
  // REPLACE WITH YOUR OWN
  const my_name = "Lu Hou Yang"; // replace with your name
  const my_phone = "+6010-258 0630"; // +6010-000 0000 replace with your phone number

  // EMAIL TEMPLATE DOCS
  // REPLACE WITH OWN DETAILS
  const docId = '1dAga7XKv4ZfhHakcWnKhM3JHtUj-6KLPNuvegT93WC8';
  // Important Note – To find your document’s ID, open your Google Doc and look at the URL. The ID is the long string of characters between https://docs.google.com/document/d/ and /edit.
  const url = `https://docs.google.com/document/d/${docId}/export?format=html`;
  var content = UrlFetchApp.fetch(url).getContentText();

  // cc & bcc
  // var cc = "luhouyang@gmail.com,luyangbang@gmail.com"; // add people to cc here
  // var bcc = "luhouyang@gmail.com,luyangbang@gmail.com"; // add people to bcc here

  // add row index of people/company you want to exclude
  // var exclude = []

  for (var i = 2; i < lastRow + 1; i++) {
    // if (exclude.indexOf(i) == -1){ }
    var alumni_name = spreadSheet.getRange(i, 2).getValue();
    var email = spreadSheet.getRange(i, 3).getValue();
    var graduation_year = spreadSheet.getRange(i, 4).getValue();
    var shipping_address = spreadSheet.getRange(i, 6).getValue();
    var postcode = spreadSheet.getRange(i, 7).getValue();

    var subject = "GDSC-UTP Swag Pre-Order";

    var message = HtmlService.createHtmlOutput(content).getContent();
    message = message.replace('<style type="text/css">', '<link href="https://fonts.cdnfonts.com/css/google-sans" rel="stylesheet"><style type="text/css">');

    while(message.includes('"Google Sans"')){
      message = message.replace('"Google Sans"', '"Product Sans"')
    }
    
    while (message.includes('"Arial"')){
      message = message.replace('"Arial"', '"Product Sans"')
    }

    // var message = HtmlService.createHtmlOutputFromFile('test.html').getContent();
    // console.log(message);

    message = message.replace('{{ALUMNI_NAME}}', alumni_name);
    message = message.replace('{{SHIPPING_ADDRESS}}', shipping_address);
    message = message.replace('{{POSTCODE}}', postcode);

    try {
      MailApp.sendEmail({
        to: email,
        // cc: cc // UNCOMMENT TO CC
        // bcc: bcc, // UNCOMMENT TO BCC
        subject: subject,
        htmlBody: message
      });
    } catch (error) {
      Logger.log(error);
    }
  }
}