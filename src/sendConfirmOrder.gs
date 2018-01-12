// Version 1.4.1

var PRODUCTION_EMAIL = 'production@example.com';

/*
 * Sends a command to the production email.
 *
 * @param command: name of the command
 * @type  command: String
 * @param message: body of the command email
 * @type  message: String
 */
function sendCommand(command, message) {
  GmailApp.sendEmail(
    PRODUCTION_EMAIL, 
    Utilities.formatString('CMD - %s', command), 
    message);
}

/*
 * Gets range of cells by name.
 *
 * @param name: name of the range
 * @type  name: String
 * @return: the cell range
 * @rtype: Range
 */
function getRange(spreadsheet, ui, name) {
  var range = spreadsheet.getRangeByName(name);
  if (range == null) {
    ui.alert('Invalid order spreadsheet: missing named range "' + name + '".');
    return null;
  }

  return range;
}

/*
 * Gets the value of a single-cell range.
 *
 * @param name: name of the range
 * @type  name: String
 * @param invalid: an invalid value for the cell
 * @type  invalid: Object
 * @return: the cell value
 * @rtype: Object
 */
function getValue(spreadsheet, ui, name, invalid) {
  var range = getRange(spreadsheet, ui, name);
  var value = range.getValue();
  if (value === invalid || value == null || range.getCell(1, 1).isBlank()) {
    instructions = ''
    if (value === invalid) {
      instructions = ' Please change the default value.'
    }

    ui.alert(
      'Invalid order spreadsheet: bad value for "' + name + '".' + 
      instructions);

    return null;
  }
  
  return value;
}

/*
 * Generates a name based on information derived from the spreadsheet.
 *
 * @param orderNum: identification number associated with the order
 * @type  orderNum: Number
 * @param customerCode: character code associated with the customer
 * @type  customerCode: String
 * @param projectName: project name associated with the order
 * @type  projectName: String
 * @return: the generated name
 * @rtype: String
 */
function getName(orderNum, customerCode, projectName) {
  var originalName = Utilities.formatString(
    'PO %d %s %s', orderNum, customerCode, projectName);

  var name = originalName;
  var files = DriveApp.getFilesByName(name);
  
  var i = 0;
  while (files.hasNext()) {
    i += 1;
    name = Utilities.formatString('%s (%d)', originalName, i);
    files = DriveApp.getFilesByName(name);
  }
  
  return name;
}

/*
 * Publishes the spreadsheet by renaming it and making it sharable by link.  
 * Also creates an associated Google Drive folder and some pre-defined contents.
 */
function confirmOrder() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.setName('order1');
  if (spreadsheet == null) {
    ui.alert('Invalid order spreadsheet: no active spreadsheet.');
    return;
  }
  
  // get information from the spreadsheet
  var orderNum = getValue(spreadsheet, ui, 'orderNum', 'PO');
  if (orderNum == null) {
    return;
  }

  var customerCode = getValue(spreadsheet, ui, 'customerCode', 'cID');
  if (customerCode == null) {
    return;
  }
  
  var projectName = getValue(spreadsheet, ui, 'projectName', 'Keyword');
  if (projectName == null) {
    return;
  }

  var port = getValue(spreadsheet, ui, 'port', '');
  if (port == null) {
    return;
  }

  var quoteCols = getRange(spreadsheet, ui, 'quote');
  var orderValues = getRange(spreadsheet, ui, 'saveOnConfirm');
  var orderNumRange = getRange(spreadsheet, ui, 'orderNum');
  
  // confirm dialog
  var confirm = ui.alert(
    'Please confirm', 
    'Are you sure?  This action is not reversable.', 
    ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.NO) {
    return;
  }
  
  // modify spreadsheet name and sharing
  var name = getName(orderNum, customerCode, projectName);
  var file = DriveApp.getFileById(spreadsheet.getId());
  file.setName(name);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // get/create parent folder
  var parentName = '0 Current Order';
  var parents = DriveApp.getFoldersByName(parentName);
  var parent;
  if (parents.hasNext()) {
    parent = parents.next();
  }
  else {
    parent = DriveApp.createFolder(parentName);
  }

  // create order folder structure and move order file
  var fileParents = file.getParents();
  var folder = parent.createFolder(name);
  folder.addFile(file);
  while (fileParents.hasNext()) {
    fileParents.next().removeFile(file);
  }

  folder.createFolder(Utilities.formatString('%s - %s', 'TECH', name));
  folder.createFolder(Utilities.formatString('%s - %s', 'SHIP', name));
  
  // send command to production
  var body = Utilities.formatString('%s\n%s', name, file.getId());
  sendCommand('Confirm Order', body);

  // hide the quote columns
  spreadsheet.hideColumn(quoteCols);

  // save the order values
  var cell = null;
  for (var i = 1; i <= orderValues.getHeight(); i++) {
    for (var j = 1; j <= orderValues.getWidth(); j++) {
      cell = orderValues.getCell(i, j);
      cell.setValue(cell.getValue());
    }
  }

  // save the order number
  cell = orderNumRange.getCell(1, 1);
  cell.setValue(cell.getValue());
}

/*
 * Sends an quote to the customer via email and hides the quote columns.
 */
function emailQuote() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet == null) {
    ui.alert('Invalid order spreadsheet: no active spreadsheet.');
    return;
  }
  
  // get information from the spreadsheet
  var customerEmail = getValue(spreadsheet, ui, 'customerEmail', '');
  if (customerEmail == null) {
    return;
  }

  var date = getValue(spreadsheet, ui, 'date', 'Date');
  if (date == null) {
    return;
  }
  
  var projectName = getValue(spreadsheet, ui, 'projectName', 'Keyword');
  if (projectName == null) {
    return;
  }

  var height = getValue(spreadsheet, ui, 'height', '');
  if (height == null) {
    return;
  }

  var width = getValue(spreadsheet, ui, 'width', '');
  if (width == null) {
    return;
  }

  var unit = getValue(spreadsheet, ui, 'unit', '');
  if (unit == null) {
    return;
  }

  var model = getValue(spreadsheet, ui, 'model', '');
  if (model == null) {
    return;
  }

  var quoteCols = getRange(spreadsheet, ui, 'quote');

  // confirm dialog
  var confirm = ui.alert(
    'Please confirm', 
    Utilities.formatString(
      'Are you sure?  This document will be shared with %s.', customerEmail), 
    ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.NO) {
    return;
  }

  // set the sharing to "anyone who has the link"
  var file = DriveApp.getFileById(spreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // send the email
  var quote = Utilities.formatString(
    '%s %sx%s %s %s', projectName, height, width, unit, model);

  var body = Utilities.formatString(
    'Dear valued customer,\n\nYour quote for %s is available at the following address:\n%s', 
    quote, file.getUrl());

  GmailApp.sendEmail(
    customerEmail, 
    Utilities.formatString(
      'Quote %s %s', 
      Utilities.formatDate(date, 'MDT', 'MM-dd-yy'), 
      quote), 
    body);

  // hide the quote columns
  spreadsheet.hideColumn(quoteCols);
}
