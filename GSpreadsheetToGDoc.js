// Row number from where to fill in the data (starts as 1 = first row)
var CUSTOMER_ID = 2;
var ACTIVE_ROW = -1;

// Google Doc id from the document template
// (Get ids from the URL)
var SOURCE_TEMPLATE = "FILL_ME_IN";

// In which spreadsheet we have all the customer data
var CUSTOMER_SPREADSHEET = "FILL_ME_IN";

// In which Google Drive we toss the target documents
var TARGET_FOLDER = "FILL_ME_IN";

/**
 * Return spreadsheet row content as JS array.
 *
 * Note: We assume the row ends when we encounter
 * the first empty cell. This might not be 
 * sometimes the desired behavior.
 *
 * Rows start at 1, not zero based!!! 🙁
 *
 */
function getRowAsArray(sheet, row) {
  var dataRange = sheet.getRange(row, 1, 1, 99);
  var data = dataRange.getValues();
  var columns = [];

  for (i in data) {
    var row = data[i];

    Logger.log("Got row", row);

    for(var l=0; l<99; l++) {
        var col = row[l];
        // First empty column interrupts
        if(!col) {
            break;
        }

        columns.push(col);
    }
  }

  return columns;
}

/**
 * Duplicates a Google Apps doc
 *
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(sourceId, name) {
    var source = DriveApp.getFileById(sourceId);
    var newFile = source.makeCopy(name);

    var targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
    targetFolder.addFile(newFile);

    return DocumentApp.openById(newFile.getId());
}

/**
 * Search a paragraph in the document and replaces it with the generated text 
 */
function replaceParagraph(doc, keyword, newText) {
  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();

    if(text.indexOf(keyword) >= 0) {
      p.setText(newText);
      p.setBold(false);
    }
  } 
}

/**
 * Search a keyword in the document and replaces it with the text in the Spreadsheet 
 */
function replaceKeyWords(doc, keyword, newText) {
  doc.getBody().replaceText('KEY'+keyword, newText);
}

/**
 * Script entry point
 */
function generateCustomerInvoice() {
  var CUSTOMER_ID = Browser.inputBox("Enter customer number in the spreadsheet", Browser.Buttons.OK_CANCEL);
  var data = SpreadsheetApp.openById(CUSTOMER_SPREADSHEET);

  // Fetch variable names
  // they are column names in the spreadsheet
  var sheet = data.getSheets()[0];
  var columns = getRowAsArray(sheet, 1);

  Logger.log("Processing columns:" + columns);
  var customerData = getRowAsArray(sheet, CUSTOMER_ID);    
  Logger.log("Processing data:" + customerData);

  // Assume first column holds the name of the customer
  var customerName = customerData[0];

  var target = createDuplicateDocument(SOURCE_TEMPLATE, customerName + " invoice");
  Logger.log("Created new document:" + target.getId());

  for(var i=1; i<columns.length; i++) {
      var key = columns[i] + ":"; 
      var text = customerData[i] || ""; 
      var value = key + " " + text;
      replaceKeyWords(target, i, customerData[i]);
  }
}

function getCurrentRow() {
  return SpreadsheetApp.getActiveSheet().getActiveSelection().getRowIndex();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoicing')
      .addItem('Generate Invoice', 'generateCustomerInvoice')
      .addToUi();
}