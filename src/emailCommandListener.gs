/*
 * Retrieves a label by name if it exists.  Otherwise, creates and returns a new label with the
 * given name.
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
 * Creates a folder and its contents based on the contents of a "Confirm Order" command message.
 *
 * @param message: the command message
 * @type  message: GmailMessage
 */
function cmdConfirmOrder(message) {
  var name = message.getPlainBody();
  var folder = DriveApp.createFolder(name);
  folder.createFolder(Utilities.formatString('%s %s', name, 'TECH'));
  folder.createFolder(Utilities.formatString('%s %s', name, 'SHIP'));
}

/*
 * Checks emails for unread command messages and executes the associated commands.  Should be 
 * executed periodically by a time-driven trigger.  Relies on an external filter to automatically
 * label command messages.
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
        
        // confirm order
        if (message.getSubject() === 'CMD - Confirm Order') {
          cmdConfirmOrder(message);
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
