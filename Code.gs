function CreateVA(e) {

  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();
  var data = range.getValues();
  var sheetname = range.getSheet().getName();
  if (column == 5 && sheetname == 'Create VA'){
  var values = SpreadsheetApp.getActiveSheet().getRange(row, 1, 1, 6).getValues()[0];
  var external_id = values[0];
  var VANumber = values[1];
  var invoice = values[2]
  var expected_amount = invoice + 3650;
  var name = values[3];
  var description = values[4];
  var expiration_date = values[5];
  var apiKey = '';
  var Basic = Utilities.base64Encode(apiKey + ':0');
  var header = {
   "Authorization" : "Basic " + Basic
  };
  var formDataMandiri = {
    "external_id": external_id,
    "expiration_date": expiration_date,
    "bank_code": "MANDIRI",
    "name": name,
    "virtual_account_number": "5301" +VANumber+ "",
    "is_closed": true,
    "expected_amount": expected_amount
    
  }
  var optionsMandiri = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataMandiri)
  };
  //var formDataBCA = {
  //  "external_id": external_id,
  //  "bank_code": "BCA",
  //  "name": name,
  //  "virtual_account_number": "9999" +VANumber+ "",
  //  "is_closed": true,
  //  "expected_amount": expected_amount
    
  //}
  //var optionsBCA = {
  //'method' : 'post',
  //'headers' :header,
  //'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  //'payload' : JSON.stringify(formDataBCA)
  //};
  var formDataBRI = {
    "external_id": external_id,
    "expiration_date": expiration_date,
    "bank_code": "BRI",
    "name": name,
    "virtual_account_number": "5301" +VANumber+ "",
    "is_closed": true,
    "expected_amount": expected_amount,
    "description": description
    
  }
  var optionsBRI = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataBRI)
  };
  var formDataBNI = {
    "external_id": external_id,
    "expiration_date": expiration_date,
    "bank_code": "BNI",
    "name": name,
    "virtual_account_number": "530100" +VANumber+ "",
    "is_closed": true,
    "expected_amount": expected_amount
    
  }
  var optionsBNI = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataBNI)
  };
  var formDataPERMATA = {
    "external_id": external_id,
    "expiration_date": expiration_date,
    "bank_code": "PERMATA",
    "name": name,
    "virtual_account_number": "530100" +VANumber+ "",
    "is_closed": true,
    "expected_amount": expected_amount
    
  }
  var optionsPERMATA = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataPERMATA)
  };
  var CreateVAMandiri = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts", optionsMandiri);
  //var CreateVABCA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts", optionsBCA);
  var CreateVABRI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts", optionsBRI);
  var CreateVABNI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts", optionsBNI);
  var CreateVAPERMATA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts", optionsPERMATA);
  if ( CreateVAMandiri.getResponseCode() == 200 && CreateVABRI.getResponseCode() == 200 && CreateVABNI.getResponseCode() == 200 && CreateVAPERMATA.getResponseCode() == 200 ) {
    var jsonMandiri = CreateVAMandiri.getContentText();
  //  var jsonBCA = CreateVABCA.getContentText();
    var jsonBRI = CreateVABRI.getContentText();
    var jsonBNI = CreateVABNI.getContentText();
    var jsonPERMATA = CreateVAPERMATA.getContentText();
    var dataMandiri = JSON.parse(jsonMandiri);
  //  var dataBCA = JSON.parse(jsonBCA);
    var dataBRI = JSON.parse(jsonBRI);
    var dataBNI = JSON.parse(jsonBNI);
    var dataPERMATA = JSON.parse(jsonPERMATA);
      var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
      var sheet = ss.getSheetByName("Edit Invoice");
      var lc = sheet.getLastColumn();
      var lr = sheet.getLastRow();
    var insertExternal_id = sheet.getRange(lr + 1, 1, 1, 1).setValue(external_id);
    var insertName = sheet.getRange(lr + 1, 2, 1, 1).setValue(name);
    var insertVAIDMandiri = sheet.getRange(lr + 1, 5, 1, 1).setValue(dataMandiri["id"]);
  //  var insertVAIDBCA = sheet.getRange(lr + 1, 5, 1, 1).setValue(dataBCA["id"]);
    var insertVAIDBRI = sheet.getRange(lr + 1, 7, 1, 1).setValue(dataBRI["id"]);
    var insertVAIDBNI = sheet.getRange(lr + 1, 8, 1, 1).setValue(dataBNI["id"]);
    var insertVAIDPERMATA = sheet.getRange(lr + 1, 9, 1, 1).setValue(dataPERMATA["id"]);
    var insertVANumberMandiri = sheet.getRange(lr + 1, 10, 1, 1).setValue(dataMandiri["account_number"]);
 //  var insertVANumberBCA = sheet.getRange(lr + 1, 10, 1, 1).setValue(dataBCA["account_number"]);
    var insertVANumberBRI = sheet.getRange(lr + 1, 12, 1, 1).setValue(dataBRI["account_number"]);
    var insertVANumberBNI = sheet.getRange(lr + 1, 13, 1, 1).setValue(dataBNI["account_number"]);
    var insertVANumberPERMATA = sheet.getRange(lr + 1, 14, 1, 1).setValue(dataPERMATA["account_number"]);
    var insertAmount = sheet.getRange(lr + 1, 22, 1, 1).setValue(invoice);
    var insertDescription = sheet.getRange(lr + 1, 23, 1, 1).setValue(description);
    var insertExpiration_date = sheet.getRange(lr + 1, 24, 1, 1).setValue(expiration_date);
  }

  }

}
function UpdateVA(e) {

  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();
  var data = range.getValues();
  var sheetname = range.getSheet().getName();
  if (column == 23 && sheetname == 'Edit Invoice'){
  var values = SpreadsheetApp.getActiveSheet().getRange(row, 1, 1, 24).getValues()[0];
  var expected_amount = values[20] + 3650;
  var Mandiri = values[4];
  //var BCA = values[4];
  var BRI = values[6];
  var BNI = values[7];
  var Permata = values[8];
  var description = values[22];
  var expiration_date = values[23];
  var apiKey = '';
  var Basic = Utilities.base64Encode(apiKey + ':0');
  var header = {
   "Authorization" : "Basic " + Basic
  };
  var formData = {
    "expiration_date": expiration_date,
      "expected_amount": expected_amount
  }
  var formDataBRI = {
    "expiration_date": expiration_date,
    "expected_amount": expected_amount,
    "description": description
  }
  var options = {
  'method' : 'patch',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formData)
  };
  var optionsBRI = {
  'method' : 'patch',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataBRI)
  };
  var UpdateVAMandiri = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Mandiri + "", options);
  //var UpdateVABCA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BCA + "", options);
  var UpdateVABRI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BRI + "", optionsBRI);
  var UpdateVABNI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BNI + "", options);
  var UpdateVAPermata = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Permata + "", options);
  if ( UpdateVAMandiri.getResponseCode() == 200 ) {
    var status = SpreadsheetApp.getActiveSheet().getRange(row, 25, 1, 1).setValue('Unpaid');
    var resetcolor = SpreadsheetApp.getActiveSheet().getRange(row, 1, 1, 27).setBackground("white");
    var resetlasttimepayment = SpreadsheetApp.getActiveSheet().getRange(row, 26, 1, 1).setValue('');
  }

  }

}
function SendInvoiceSecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "") {
        i=i+1;
      var VAnis = sheet.getRange(i,1).getValue().slice(-6);
      var VABRI = sheet.getRange(i,12).getValue();
      var VAPERMATA = sheet.getRange(i,14).getValue();
      var studentname = sheet.getRange(i,2).getValue();
      var grade = sheet.getRange(i,3).getValue();
      var parentsemail = sheet.getRange(i,4).getValue();
      var sd = sheet.getRange(i,15).getValue();
      var spp = sheet.getRange(i,16).getValue();
      var ekskul = sheet.getRange(i,17).getValue();
      var upd = sheet.getRange(i,18).getValue();
      var ittihada = sheet.getRange(i,19).getValue();
      var buku = sheet.getRange(i,20).getValue();
      var lainlain = sheet.getRange(i,21).getValue();
      var expected_amountt = sheet.getRange(i, 22).getValue();
      var expected_amount = expected_amountt + 3650;
      var Mandiri = sheet.getRange(i, 5).getValue();
      //var BCA = row[4];
      var BRI = sheet.getRange(i, 7).getValue();
      var BNI = sheet.getRange(i, 8).getValue();
      var Permata = sheet.getRange(i, 9).getValue();
      var description = sheet.getRange(i, 23).getValue();
      var expiration_date = sheet.getRange(i, 24).getValue();
      var information = sheet.getRange(i, 27).getValue();
      var apiKey = '';
      var Basic = Utilities.base64Encode(apiKey + ':0');
      var header = {
        "Authorization" : "Basic " + Basic
      };
      var formData = {
        "expiration_date": expiration_date,
        "expected_amount": expected_amount
      }
      var formDataBRI = {
        "expiration_date": expiration_date,
        "expected_amount": expected_amount,
        "description": description
      }
      var options = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formData)
      };
      var optionsBRI = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formDataBRI)
      };
      var UpdateVAMandiri = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Mandiri + "", options);
      //var UpdateVABCA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BCA + "", options);
      var UpdateVABRI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BRI + "", optionsBRI);
      var UpdateVABNI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BNI + "", options);
      var UpdateVAPermata = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Permata + "", options);
      if ( UpdateVAMandiri.getResponseCode() == 200 ) {
        var status = sheet.getRange(i, 25, 1, 1).setValue('Unpaid');
        var resetcolor = sheet.getRange(i, 1, 1, 27).setBackground("white");
        var resetlasttimepayment = sheet.getRange(i, 26, 1, 1).setValue('');
        var date = new Date();
        var mt = date.getMonth();
        var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",  "November", "December"];
        var currentMonth = months[mt];
        var BodyTemplate = HtmlService.createTemplateFromFile("body");
        BodyTemplate.studentname = studentname;
        BodyTemplate.vanumber = VAnis;
        BodyTemplate.vabri = VABRI;
        BodyTemplate.vapermata = VAPERMATA;
        BodyTemplate.grade = grade;
        BodyTemplate.sd = sd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.spp = spp.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ekskul = ekskul.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.upd = upd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ittihada = ittihada.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.buku = buku.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.lainlain = lainlain.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.total = expected_amount.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.month = currentMonth;
        BodyTemplate.information = information;
        var HtmlforBody = BodyTemplate.evaluate().getContent();
        var subject = "School Payment Ananda " +studentname+ " - " +currentMonth+ " 2021"
        var options = {};
        options.htmlBody = HtmlforBody;
        GmailApp.sendEmail("" +parentsemail+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );
      } 
    }
  }
  }
  function SendRevisionSecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "revisi") {
        i=i+1;
      var VAnis = sheet.getRange(i,1).getValue().slice(-6);
      var VABRI = sheet.getRange(i,12).getValue();
      var VAPERMATA = sheet.getRange(i,14).getValue();
      var studentname = sheet.getRange(i,2).getValue();
      var grade = sheet.getRange(i,3).getValue();
      var parentsemail = sheet.getRange(i,4).getValue();
      var sd = sheet.getRange(i,15).getValue();
      var spp = sheet.getRange(i,16).getValue();
      var ekskul = sheet.getRange(i,17).getValue();
      var upd = sheet.getRange(i,18).getValue();
      var ittihada = sheet.getRange(i,19).getValue();
      var buku = sheet.getRange(i,20).getValue();
      var lainlain = sheet.getRange(i,21).getValue();
      var expected_amountt = sheet.getRange(i, 22).getValue();
      var expected_amount = expected_amountt + 3650;
      var Mandiri = sheet.getRange(i, 5).getValue();
      //var BCA = row[4];
      var BRI = sheet.getRange(i, 7).getValue();
      var BNI = sheet.getRange(i, 8).getValue();
      var Permata = sheet.getRange(i, 9).getValue();
      var description = sheet.getRange(i, 23).getValue();
      var expiration_date = sheet.getRange(i, 24).getValue();
      var information = sheet.getRange(i, 27).getValue();
      var apiKey = '';
      var Basic = Utilities.base64Encode(apiKey + ':0');
      var header = {
        "Authorization" : "Basic " + Basic
      };
      var formData = {
        "expiration_date": expiration_date,
        "expected_amount": expected_amount
      }
      var formDataBRI = {
        "expiration_date": expiration_date,
        "expected_amount": expected_amount,
        "description": description
      }
      var options = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formData)
      };
      var optionsBRI = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formDataBRI)
      };
      var UpdateVAMandiri = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Mandiri + "", options);
      //var UpdateVABCA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BCA + "", options);
      var UpdateVABRI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BRI + "", optionsBRI);
      var UpdateVABNI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BNI + "", options);
      var UpdateVAPermata = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Permata + "", options);
      if ( UpdateVAMandiri.getResponseCode() == 200 ) {
        var status = sheet.getRange(i, 25, 1, 1).setValue('Unpaid');
        var resetcolor = sheet.getRange(i, 1, 1, 27).setBackground("white");
        var resetlasttimepayment = sheet.getRange(i, 26, 1, 1).setValue('');
        var date = new Date();
        var mt = date.getMonth();
        var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",  "November", "December"];
        var currentMonth = months[mt];
        var BodyTemplate = HtmlService.createTemplateFromFile("body");
        BodyTemplate.studentname = studentname;
        BodyTemplate.vanumber = VAnis;
        BodyTemplate.vabri = VABRI;
        BodyTemplate.vapermata = VAPERMATA;
        BodyTemplate.grade = grade;
        BodyTemplate.sd = sd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.spp = spp.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ekskul = ekskul.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.upd = upd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ittihada = ittihada.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.buku = buku.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.lainlain = lainlain.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.total = expected_amount.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.month = currentMonth;
        BodyTemplate.information = information;
        var HtmlforBody = BodyTemplate.evaluate().getContent();
        var subject = "School Payment Ananda " +studentname+ " - " +currentMonth+ " 2021"
        var options = {};
        options.htmlBody = HtmlforBody;
        GmailApp.sendEmail("" +parentsemail+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );
      } 
    }
  }
  }
  function SendReminderSecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "Unpaid") {
        i=i+1;
      var VAnis = sheet.getRange(i,1).getValue().slice(-6);
      var studentname = sheet.getRange(i,2).getValue();
      var grade = sheet.getRange(i,3).getValue();
      var parentsemail = sheet.getRange(i,4).getValue();
      var sd = sheet.getRange(i,15).getValue();
      var spp = sheet.getRange(i,16).getValue();
      var ekskul = sheet.getRange(i,17).getValue();
      var upd = sheet.getRange(i,18).getValue();
      var ittihada = sheet.getRange(i,19).getValue();
      var buku = sheet.getRange(i,20).getValue();
      var lainlain = sheet.getRange(i,21).getValue();
      var expected_amountt = sheet.getRange(i, 22).getValue();
      var expected_amount = expected_amountt + 3650;
      var information = sheet.getRange(i, 27).getValue();
        var date = new Date();
        var mt = date.getMonth();
        var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",  "November", "December"];
        var currentMonth = months[mt];
        var BodyTemplate = HtmlService.createTemplateFromFile("reminder");
        BodyTemplate.studentname = studentname;
        BodyTemplate.vanumber = VAnis;
        BodyTemplate.grade = grade;
        BodyTemplate.sd = sd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.spp = spp.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ekskul = ekskul.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.upd = upd.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.ittihada = ittihada.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.buku = buku.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.lainlain = lainlain.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.total = expected_amount.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
        BodyTemplate.month = currentMonth;
        BodyTemplate.information = information;
        var HtmlforBody = BodyTemplate.evaluate().getContent();
        var subject = "Reminder for " +currentMonth+ " 2021 payment - Ananda " +studentname+ ""
        var options = {};
        options.htmlBody = HtmlforBody;
        GmailApp.sendEmail("" +parentsemail+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );
      var status = sheet.getRange(i, 25, 1, 1).setValue('Still Unpaid'); 
    }
  }
  }

  function SendNotificationSecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "") {
        i=i+1;
      var VAnis = sheet.getRange(i,1).getValue().slice(-6);
      var studentname = sheet.getRange(i,2).getValue();
      var grade = sheet.getRange(i,3).getValue();
      var parentsemail = sheet.getRange(i,4).getValue();
        var BodyTemplate = HtmlService.createTemplateFromFile("notification");
        BodyTemplate.studentname = studentname;
        BodyTemplate.vanumber = VAnis;
        BodyTemplate.grade = grade;
        var HtmlforBody = BodyTemplate.evaluate().getContent();
        var subject = "" +studentname+ " - Mutiara Harapan Islamic School New Payment System"
        var options = {};
        options.htmlBody = HtmlforBody;
        options.cc = "secondary@mutiaraharapan.sch.id"
        GmailApp.sendEmail("" +parentsemail+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );
      var status = sheet.getRange(i, 25, 1, 1).setValue('SENT EMAIL');
      
    }
  }
  }
  function CloseVASecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "") {
        i=i+1;
      var VAnis = sheet.getRange(i,1).getValue().slice(-6);
      var studentname = sheet.getRange(i,2).getValue();
      var grade = sheet.getRange(i,3).getValue();
      var parentsemail = sheet.getRange(i,4).getValue();
      var sd = sheet.getRange(i,15).getValue();
      var spp = sheet.getRange(i,16).getValue();
      var ekskul = sheet.getRange(i,17).getValue();
      var upd = sheet.getRange(i,18).getValue();
      var ittihada = sheet.getRange(i,19).getValue();
      var buku = sheet.getRange(i,20).getValue();
      var lainlain = sheet.getRange(i,21).getValue();
      var expected_amountt = sheet.getRange(i, 22).getValue();
      var expected_amount = expected_amountt + 3650;
      var Mandiri = sheet.getRange(i, 5).getValue();
      //var BCA = row[4];
      var BRI = sheet.getRange(i, 7).getValue();
      var BNI = sheet.getRange(i, 8).getValue();
      var Permata = sheet.getRange(i, 9).getValue();
      var description = sheet.getRange(i, 23).getValue();
      var expiration_date = sheet.getRange(i, 24).getValue();
      var information = sheet.getRange(i, 27).getValue();
      var apiKey = '';
      var Basic = Utilities.base64Encode(apiKey + ':0');
      var header = {
        "Authorization" : "Basic " + Basic
      };
      var formData = {
        "expiration_date": expiration_date
      }
      var formDataBRI = {
        "expiration_date": expiration_date
      }
      var options = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formData)
      };
      var optionsBRI = {
        'method' : 'patch',
        'headers' :header,
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload' : JSON.stringify(formDataBRI)
      };
      var UpdateVAMandiri = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Mandiri + "", options);
      //var UpdateVABCA = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BCA + "", options);
      var UpdateVABRI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BRI + "", optionsBRI);
      var UpdateVABNI = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + BNI + "", options);
      var UpdateVAPermata = UrlFetchApp.fetch("https://api.xendit.co/callback_virtual_accounts/" + Permata + "", options); 
      var status = sheet.getRange(i, 25, 1, 1).setValue('Close VA');
    }
  }
  }

function SuspendemailSecondary() {
  var ss= SpreadsheetApp.openById("1rYtqT8w5zcl3Zw-rE1DHwrzEawvpt8bGyIVS2bOEsQY");
  var sheet = ss.getSheetByName("Edit Invoice");
  var lc = sheet.getLastColumn();
  var lr = sheet.getLastRow();
  var lookupRange = sheet.getRange(1, 1, lr, lc).getValues();
     for(var i = 0;i<lookupRange.length; i++){
      if (lookupRange[i][24] == "Unpaid" && lookupRange[i][27] !== '' ) {
        i=i+1;
      var studentemail = sheet.getRange(i,28).getValue();
      var formData = {
        "suspended": true
      }
      var suspendemail = AdminDirectory.Users.patch(formData, studentemail);
      var status = sheet.getRange(i, 25, 1, 1).setValue('Suspend');
    }
  }
  }
