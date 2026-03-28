function createAndSendVendorFiles() {


var ss = SpreadsheetApp.getActiveSpreadsheet();
var baseSheet = ss.getSheetByName("Base");
var poSheet = ss.getSheetByName("PO_Level");
var emailSheet = ss.getSheetByName("Vendor_Email_Master");
var templateSheet = ss.getSheetByName("Mail_Template");


var baseData = baseSheet.getDataRange().getValues();
var poData = poSheet.getDataRange().getValues();
var emailData = emailSheet.getDataRange().getValues();


// -------- MAIL TEMPLATE --------
var lastRow = templateSheet.getLastRow();
var templateData = templateSheet.getRange("A1:A" + lastRow).getValues();
var template = templateData.map(function(row){
return row[0];
}).join("<br>");


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


var vendorMap = {};


// Collect vendors
for (var i = 1; i < poData.length; i++) {
var vendor = String(poData[i][2]).trim();
if (vendor) vendorMap[vendor] = true;
}


var vendors = Object.keys(vendorMap);
var processed = 0;


for (var v = 0; v < vendors.length; v++) {


var vendor = vendors[v];


var vendorEmail = "";
var ccEmail = "";


for (var j = 1; j < emailData.length; j++) {
if (String(emailData[j][0]).trim() === vendor) {
vendorEmail = emailData[j][1];
ccEmail = emailData[j][2];
break;
}
}


if (!vendorEmail) continue;


// 🔥 Create spreadsheet per vendor
var tempFile = SpreadsheetApp.create(vendor + "_Working");
var tempId = tempFile.getId();


// Move file to folder
var driveFile = DriveApp.getFileById(tempId);
mainFolder.addFile(driveFile);
DriveApp.getRootFolder().removeFile(driveFile);


// -------- BASE SHEET --------
var newBase = tempFile.getSheets()[0];
newBase.setName("Base");


var baseFiltered = [];
baseFiltered.push(baseData[0]);


for (var b = 1; b < baseData.length; b++) {
if (String(baseData[b][10]).trim() === vendor) {
baseFiltered.push(baseData[b]);
}
}


newBase.getRange(1,1,baseFiltered.length,baseFiltered[0].length)
.setValues(baseFiltered);


formatSheet(newBase, baseFiltered);


// -------- PO_LEVEL SHEET --------
var newPO = tempFile.insertSheet("PO_Level");


var poFiltered = [];
poFiltered.push(poData[0]);


for (var p = 1; p < poData.length; p++) {
if (String(poData[p][2]).trim() === vendor) {
poFiltered.push(poData[p]);
}
}


newPO.getRange(1,1,poFiltered.length,poFiltered[0].length)
.setValues(poFiltered);


formatSheet(newPO, poFiltered);


SpreadsheetApp.flush();
Utilities.sleep(500);


// -------- EXPORT TO EXCEL WITH RETRY --------
var exportUrl = "https://docs.google.com/spreadsheets/d/" + tempId + "/export?format=xlsx";


var response;
for (var attempt = 0; attempt < 3; attempt++) {
try {
response = UrlFetchApp.fetch(exportUrl, {
headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
});
break;
} catch (err) {
Utilities.sleep(2000); // retry delay
}
}


var excelBlob = response.getBlob()
.setName(vendor + "_" + todayDate + ".xlsx");


var excelFile = mainFolder.createFile(excelBlob);


// -------- SEND MAIL --------
GmailApp.sendEmail(
vendorEmail,
"FKI Open PO & FSN Dump || " + vendor + " || " + todayDate,
"Please see attachment",
{
htmlBody: template.replace("{{vendor}}", vendor),
cc: ccEmail,
attachments: [excelFile.getBlob()]
}
);


processed++;


// Prevent 429 error
Utilities.sleep(1500);


// Extra pause after every 8 vendors
if (processed % 8 == 0) {
Utilities.sleep(5000);
}


}


}


// -------- FORMAT FUNCTION --------
function formatSheet(sheet, data) {


var range = sheet.getRange(1,1,data.length,data[0].length);


range
.setHorizontalAlignment("center")
.setVerticalAlignment("middle")
.setBorder(true,true,true,true,true,true);


sheet.getRange(1,1,1,data[0].length).setFontWeight("bold");
sheet.autoResizeColumns(1, data[0].length);


}


