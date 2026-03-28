# Automated Email Delivery of Daily Flipkart Data Dumps.
-----------------------------------------------------------------------------------------------------------

Developed a JavaScript-based automation to generate and share daily Flipkart brand PO and lifestyle data dumps via email, reducing manual effort and improving reporting efficiency.

----------------------------------------------------------------------------------------------------------
function createAndSendVendorFiles() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheet = ss.getSheetByName("Base");
  var emailSheet = ss.getSheetByName("Vendor_Email_Master");
  var templateSheet = ss.getSheetByName("Mail_Template");

  var baseData = baseSheet.getDataRange().getValues();
  var emailData = emailSheet.getDataRange().getValues();


  var lastRow = templateSheet.getLastRow();
  var templateData = templateSheet.getRange("A1:A" + lastRow).getValues();
  var template = templateData.map(row => row[0]).join("<br>");

  var todayDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );

  var folderTime = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy HH-mm"
  );

  // 📁 Create main folder
  var mainFolder = DriveApp.createFolder("FKI Open PO Dump - " + folderTime);


  var emailMap = {};
  for (var i = 1; i < emailData.length; i++) {
    emailMap[String(emailData[i][0]).trim()] = {
      email: emailData[i][1],
      cc: emailData[i][2]
    };
  }

  var vendorMap = {};
  for (var i = 1; i < baseData.length; i++) {
    var vendor = String(baseData[i][3]).trim(); // ✅ Column index 3
    if (vendor) vendorMap[vendor] = true;
  }

  var vendors = Object.keys(vendorMap);

  for (var v = 0; v < vendors.length; v++) {

    var vendor = vendors[v];
    var vendorInfo = emailMap[vendor];

    if (!vendorInfo || !vendorInfo.email) continue;

    // 🔥 Create spreadsheet
    var tempFile = SpreadsheetApp.create(vendor + "_Working");
    var tempId = tempFile.getId();

    // Move to folder
    var driveFile = DriveApp.getFileById(tempId);
    mainFolder.addFile(driveFile);
    DriveApp.getRootFolder().removeFile(driveFile);

    var newBase = tempFile.getSheets()[0];
    newBase.setName("Base");

    var baseFiltered = [];
    baseFiltered.push(baseData[0]);

    for (var b = 1; b < baseData.length; b++) {
      if (String(baseData[b][3]).trim() === vendor) {
        baseFiltered.push(baseData[b]);
      }
    }

    if (baseFiltered.length <= 1) continue;

    newBase.getRange(1, 1, baseFiltered.length, baseFiltered[0].length)
      .setValues(baseFiltered);

    formatSheet(newBase, baseFiltered);

    SpreadsheetApp.flush();
    Utilities.sleep(500);

    var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempId + "/export?format=xlsx";

    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
    });

    var excelBlob = response.getBlob()
      .setName(vendor + "_" + todayDate + ".xlsx");

    var excelFile = mainFolder.createFile(excelBlob);

    try {
      GmailApp.sendEmail(
        vendorInfo.email,
        "FKI Open PO Dump || " + vendor + " || " + todayDate,
        "Please see attachment",
        {
          htmlBody: template.replace(/{{vendor}}/g, vendor),
          cc: vendorInfo.cc,
          attachments: [excelFile.getBlob()]
        }
      );
    } catch (e) {
      Logger.log("Email failed for: " + vendor);
    }

    // 🧹 Delete temp file
    DriveApp.getFileById(tempId).setTrashed(true);

    Utilities.sleep(1500);
  }
}


function formatSheet(sheet, data) {

  var range = sheet.getRange(1, 1, data.length, data[0].length);

  range
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold");
  sheet.autoResizeColumns(1, data[0].length);

}
