/**
 * Generate Google Docs based on a template document and data incoming from a Google Spreadsheet
 *
 * License: MIT
 *
 * Copyright 2013 Mikko Ohtamaa, http://opensourcehacker.com
 */

// Assumes Company name in column to for filename
var COMPANY_COLUMN = 2; // <=============================== Adjust this

// Google Doc id from the document template (Get ids from the URL)
var SOURCE_TEMPLATE = ""; // <========================== Adjust this

// In which spreadsheet we have all the customer data
var SPREADSHEET = ""; // <============================== Adjust this

// In which Google Drive we toss the target documents
var TARGET_FOLDER = ""; // <============================ Adjust this

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
  var body = doc.getBody();
  body.replaceText(keyword, newText);
}

/**
 * Script entry point
 */
function fillData() {

  var data = SpreadsheetApp.openById(SPREADSHEET);

  // Fetch variable names
  // they are column names in the spreadsheet
  var sheet = data.getSheets()[0];
  var columns = getRowAsArray(sheet, 1);

  if(mode == 1) {
    var COMPANY_ROW = sheet.getLastRow();
  }
  
  else {
    var COMPANY_ROW = data.getActiveCell().getRow();
  }

  // XXX: Cannot be accessed when run in the script editor?
  // WHYYYYYYYYY? Asking one number, too complex?
  //var COMPANY_ROW = Browser.inputBox("Enter customer number in the spreadsheet", Browser.Buttons.OK_CANCEL);
  if (!COMPANY_ROW) {
      return; 
  }

  Logger.log("Processing columns:" + columns);

  var customerData = getRowAsArray(sheet, COMPANY_ROW);  
  Logger.log("Processing data:" + customerData);

  // Assume third column holds the company name
  var companyName = customerData[COMPANY_COLUMN];

  var target = createDuplicateDocument(SOURCE_TEMPLATE, companyName + " Cover Letter");

  Logger.log("Created new document:" + target.getId());

  for (var i = 0; i < columns.length; i++) {
    var key = "{{" + columns[i] + "}}"; 

    var text = customerData[i] || ""; // No Javascript undefined
      
    // Don't replace the whole text, but leave the template text as a label
    // var value = key + " " + text;
      
    // Replace the whole text
    var value = text;
      
    replaceParagraph(target, key, value);
  }

  showDialog(target.getUrl());
}

function fillDataLatest() {
  fillData(1);
}

function fillDataSelected() {
  fillData(0);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Generate Cover Letter')
      .addItem('Latest Row', 'fillDataLatest')
      .addItem('Selected Row', 'fillDataSelected')
      .addToUi();
}

function showDialog(url) {
  var html = HtmlService.createHtmlOutput('<html><body><a href="' + url + '" target="_blank" onclick="google.script.host.close()">Link to Cover Letter</a></body></html>')
      .setWidth(150)
      .setHeight(30);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Cover Letter Created');
}
