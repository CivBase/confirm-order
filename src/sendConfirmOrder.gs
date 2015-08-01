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
  var range = spreadsheet.getRangeByName(name);
  if (range == null) {
    ui.alert('Invalid order spreadsheet: missing namned range "' + name + '".');
    return null;
  }
  
  var value = range.getValue();
  if (value == invalid || value == null) {
    ui.alert('Invalid order spreadsheet: bad value for "' + name + '".');
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
  var originalName = Utilities.formatString('PO %d %s %s', orderNum, customerCode, projectName);
  var name = originalName;
  var files = DriveApp.getFilesByName(name);
  
  var i = 0;
  while (files.hasNext()) {
    i += 1;
    name = formatString('%s (%d)', originalName, i);
    files = DriveApp.getFilesByName(name);
  }
  
  return name;
}

/*
 * Publishes the spreadsheet by renaming it and making it sharable by link.  Also creates an
 * associated Google Drive folder and some pre-defined contents.
 */
function confirmOrder() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet == null) {
    ui.alert('Invalid order spreadsheet: no active spreadsheet.');
    return;
  }
  
  // parse spreadsheet for info
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
  
  // confirm dialog
  var confirm = ui.alert(
    'Please confirm', 'Are you sure?  This action is not reversable.', ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.NO) {
    return;
  }
  
  // modify spreadsheet name and sharing
  var name = getName(orderNum, customerCode, projectName);
  var file = DriveApp.getFileById(spreadsheet.getId());
  file.setName(name);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // create stuff
  var folder = DriveApp.createFolder(name);
  folder.createFolder(Utilities.formatString('%s %s', name, 'TECH'));
  folder.createFolder(Utilities.formatString('%s %s', name, 'SHIP'));
  
  // send command to production
  var body = name;
  GmailApp.sendEmail('production@example.com', 'CMD - Confirm Order', body);
}
