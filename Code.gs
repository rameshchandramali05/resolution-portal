// ===============================
// DO GET
// ===============================

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Tech Panel")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


// ===============================
// INCLUDE HTML
// ===============================

function include(file) {

  return HtmlService
    .createHtmlOutputFromFile(file)
    .getContent();
}


// ===============================
// SAVE NEW WORK ORDER
// ===============================

function saveData(obj){

  if(
    obj.details == "" ||
    obj.issue == ""
  ){
    return "blank";
  }

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Complaints");


  var lastRow = sheet.getLastRow();

  var wo = "WO-" +
    Utilities.formatString("%03d", lastRow);


  sheet.appendRow([

    // A
    wo,

    // B
    "Pending",

    // C
    new Date(),

    // D
    obj.details,

    // E
    obj.issue,

    // F
    "",

    // G
    "",

    // H
    "",

    // I
    "",

    // J
    "",

    // K
    "",

    // L
    ""

  ]);

  return wo;
}


// ===============================
// GET DATA
// ===============================

function getData(){

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Complaints");

  return sheet
    .getDataRange()
    .getDisplayValues();
}


// ===============================
// UPDATE WORK
// ===============================

function updateWork(obj){

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Complaints");

  var data = sheet
    .getDataRange()
    .getValues();


  for(var i=1; i<data.length; i++){

    if(data[i][0] == obj.workorder){

      // F = Resolve Date
var resolveDate = obj.resolveDate;


if (resolveDate) {

  var d = new Date(resolveDate);

  sheet.getRange(i+1,6)
    .setValue(d)
    .setNumberFormat("M/d/yyyy HH:mm:ss");

} else {

  sheet.getRange(i+1,6).setValue("");
}






      // G = Resolve Detail
      sheet.getRange(i+1,7)
        .setValue(obj.workDetails);

      // H = Receive Company
      sheet.getRange(i+1,8)
        .setValue(obj.receiveCompany);

      // I = Receive Tech
      sheet.getRange(i+1,9)
        .setValue(obj.receiveTech);

      // J = Company
      sheet.getRange(i+1,10)
        .setValue(obj.companyPayment);

      // K = Tech
      sheet.getRange(i+1,11)
        .setValue(obj.techPayment);

      // L = TO-V
      // ===============================
// TO-V RUNNING BALANCE
// ===============================

// PREVIOUS ROW TO-V
var previousToV = 0;

if(i > 1){

  previousToV =
    Number(
      sheet.getRange(i,12).getValue() || 0
    );
}


// CURRENT VALUES
var receiveCompany =
  Number(obj.receiveCompany || 0);

var companyPayment =
  Number(obj.companyPayment || 0);


// NEW RUNNING BALANCE
var newToV =
  previousToV
  + companyPayment
  - receiveCompany;


// SAVE TO-V
sheet.getRange(i+1,12)
  .setValue(newToV);

      // STATUS CLOSED
      sheet.getRange(i+1,2)
        .setValue("Closed");

      return "Updated";
    }
  }
}


// ===============================
// APPROVE WORK
// ===============================

function approveWork(wo){

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Complaints");

  var data = sheet
    .getDataRange()
    .getValues();


  for(var i=1; i<data.length; i++){

    if(data[i][0] == wo){

      sheet.getRange(i+1,2)
        .setValue("Approved");
      return "Done";
    }
  }
}



function getNumbers() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Settings");
  var data = sheet.getDataRange().getValues();

  var obj = {};

  for (var i = 1; i < data.length; i++) {
    obj[data[i][0]] = data[i][1];
  }

  return obj;
}



