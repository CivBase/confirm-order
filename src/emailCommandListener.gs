// Version 1.0.0

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

  // create folders
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

  var name = body[0];
  var folder = parent.createFolder(name);
  var tech = folder.createFolder(
    Utilities.formatString('%s - %s', 'TECH', name));
  folder.createFolder(Utilities.formatString('%s - %s', 'SHIP', name));
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
    fileId = body[1];
  }

  // create the PO file
  var templateName = 'PO Template 模板测试';
  var templates = DriveApp.getFilesByName(templateName);
  if (!templates.hasNext()) {
    Logger.log(
      'ERROR: Could not find a file named "%s". Skipping this step.', 
      templateName);
  }
  else {
    var template = templates.next();
    var file = template.makeCopy(name, folder);
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
    var scheduleName = '齐博达订单进度 Schedule';
    var schedules = DriveApp.getFilesByName(scheduleName);
    if (!schedules.hasNext()) {
      Logger.log(
        'ERROR: Could not find a file named "%s". Skipping this step.', 
        scheduleName);
    }
    else {
      var schedule = schedules.next();
      var spreadsheet = SpreadsheetApp.open(schedule);

      var rangeName = 'orderFileIds';
      var range = spreadsheet.getRangeByName(rangeName);
      if (range == null) {
        Logger.log(
          'ERROR: Could not find named range "%s". Skipping this step.', 
          rangeName);
      }
      else {
        var values = range.getValues();
        var cellIndex = 0;
        while (cellIndex < values.length && values[cellIndex][0] !== '') {
          cellIndex++;
        }

        if (cellIndex < values.length) {
          var cell = range.getCell(cellIndex + 1, 1);
          cell.setValue(file.getId());
        }
        else {
          Logger.log(
            'ERROR: No empty cells available in range "%s".', rangeName);
        }
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
