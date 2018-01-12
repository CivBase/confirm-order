// Version 1.3.0

/*
 * Retrieves a label by name if it exists.  Otherwise, creates and returns a 
 * new label with the given name.
 *
 * @param message: name of a label
 * @type  message: String
 * @return: the specified label instance
 * @rtype: GmailLabel
 */
function getLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (label != undefined) {
    return label;
  }
  
  return GmailApp.createLabel(name);
}

/*
 * Creates a folder and its contents based on the contents of a "Confirm Order" 
 * command message.  The folder contents include several subfolders, a 
 * readme.txt file, and a project order spreadsheet based on the template file.
 * 
 * @param message: the command message
 * @type  message: GmailMessage
 */
function cmdConfirmOrder(message) {
  // read the email message
  var body = message.getPlainBody().split('\n');

  // get/create parent folder
  var parentName = '进行生产任务 Current Order';
  var parents = DriveApp.getFoldersByName(parentName);
  var parent;
  if (parents.hasNext()) {
    parent = parents.next();
  }
  else {
    Logger.log(
      'WARNING: Could not find a folder named "%s". Creating one now.', 
      parentName);
    parent = DriveApp.createFolder(parentName);
  }

  // create order folder structure
  var name = body[0].replace(/\s$/, '');

  var prName = name.replace(/^PO/, 'PR');
  var folder = parent.createFolder(prName);
  ship = folder.createFolder(Utilities.formatString('%s - %s', 'SHIP', prName));

  var tech = ship.createFolder(
    Utilities.formatString('%s - %s', 'TECH', prName));

  tech.createFolder('DWG 图纸');
  tech.createFolder('QC Picture 出货图片报告');
  tech.createFolder('Software 电脑加载文件');

  // copy readme
  var readmeName = 'current order readme.txt';
  var readmes = DriveApp.getFilesByName(readmeName);
  if (!readmes.hasNext()) {
    Logger.log(
      'ERROR: Could not find a file named "%s". Skipping this step.', 
      readmeName);
  }
  else {
    readmes.next().makeCopy('readme.txt', tech);
  }

  // get the PO file ID
  var fileId = '';
  if (body.length > 1) {
    fileId = body[1].replace(/\s$/g, '');
  }

  // create the PO file
  var templateName = 'PR Template 模板';
  var templates = DriveApp.getFilesByName(templateName);
  if (!templates.hasNext()) {
    Logger.log(
      'ERROR: Could not find a file named "%s". Skipping this step.', 
      templateName);
  }
  else {
    var template = templates.next();
    var file = template.makeCopy(prName, folder);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var spreadsheet = SpreadsheetApp.open(file);

    // set the PO file ID value
    var rangeName = 'orderFileId';
    var range = spreadsheet.getRangeByName(rangeName);
    if (range == null) {
      Logger.log(
        'ERROR: Could not find named range "%s". Skipping this step.', 
        rangeName);
    }
    else {
      range.setValue(fileId);
    }

    // add a schedule entry
    var scheduleName = 'Chipstar Schedule 订单进度';
    var schedules = DriveApp.getFilesByName(scheduleName);
    if (!schedules.hasNext()) {
      Logger.log(
        'ERROR: Could not find a file named "%s". Skipping this step.', 
        scheduleName);
    }
    else {
      var schedule = schedules.next();
      var spreadsheet = SpreadsheetApp.open(schedule);
      var sheetName = 'current';
      var sheet = spreadsheet.getSheetByName(sheetName);

      if (sheet == null) {
        Logger.log(
          'ERROR: Could not find a sheet named "%s". Skipping this step.', 
          sheetName);
      }
      else {
        // determine the last populated row
        var range = sheet.getRange('A2:B')
        lastRow = range.getHeight();
        for (var i = range.getHeight(); i > 0; i--) {
          if (range.getCell(i, 1).getValue() != null) {
            lastRow = i;
            break;
          }

          if (range.getCell(i, 2).getValue() != null) {
            lastRow = i;
            break;
          }
        }

        lastRow++;

        // populate the first cell with the file ID
        var cell = range.getCell(lastRow, 1);
        cell.setValue(file.getId());

        // populate the remaining cells using the importrange function
        cell = range.getCell(lastRow, 1);
        cell.setValue('=importrange(A' + lastRow + ',"Production Order!A1:N1")');
      }
    }
  }
}

/*
 * Checks emails for unread command messages and executes the associated 
 * commands.  Should be executed periodically by a time-driven trigger.  Relies 
 * on an external filter to automatically label command messages.
 */
function emailCommandListener() {
  // retrieve threads with unhandled comamnds
  var command = getLabel('command');
  var executed = getLabel('executed');
  var threads = command.getThreads();
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    
    // execute unhandled commands
    var messages = thread.getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (message.isUnread()) {
        var subject = message.getSubject();
        
        // confirm order
        if (subject === 'CMD - Confirm Order') {
          cmdConfirmOrder(message);
        }
        else {
          Logger.log('ERROR: Could not understand command "%s".', subject);
        }
        
        message.markRead();
      }
    }
    
    // archive thread
    thread.removeLabel(command);
    thread.addLabel(executed);
    thread.markRead();
    thread.moveToArchive();
  }
}
