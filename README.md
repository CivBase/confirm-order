# Google Sheets Confirm Order Script Setup
These instructions describe how to create a button on the client's order sheet that will automatically confirm the order.  When the button is clicked, it will execute a script which renames the order spreadsheet file, enables the sheet to be viewed with a link, creates empty directories for the client (named after the file), and submits a command to the production email address.

## Create the script on the order sheet.
Each Google Sheets file has an embedded Apps Script project.  The confirm order code needs to be added to this project so that the order sheet can access it.

#### 1. Open the script editor associated with the order sheet.
Open the Google Sheets editor with the sheet you want to edit and select `Tools > Script editor...`.

![tools-script-editor](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/1-tools-script-editor.png)

#### 2. Rename the Apps Script project.
Change the name of the project to something descriptive, such as "Confirm Order".

![change-name](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/2-change-name.png)

#### 3. Copy the code into the project.
Copy the code from the [`sendConfirmOrder.gs` script on GitHub](https://github.com/CivBase/confirm-order/blob/master/src/sendConfirmOrder.gs) and paste it into the `Code.gs` script in the open project.

![copy-paste-code](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/3-copy-paste-code.png)

#### 4. Change the email address to production.
Change the email address near the bottom of the script to match the production email address and save the script.

![change-email](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/4-change-email.png)

## Created named cell ranges.
The script reads information from the order sheet by checking the value of named cell ranges.  These named ranges serve as identifiers for particular information in the script.

#### 5. Define a named cell range for `customerCode`.
Back in the order sheet, right click on the cell that contains the customer code and select `Define named range...`.

![define-named-range](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/5-define-named-range.png)

#### 6. Save the cell range as `customerCode`.
Change the name of the cell range to `customerCode` and click `Done`.  It is very important that you use this exact name for the script to work.

![name-range](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/6-name-range.png)

#### 7. Define a named cell range for `projectName`.
Right click on the cell that contains the project name and select `Define named range...`.

#### 8. Save the cell range as `projectName`.
Change the name of the cell range to `projectName` and click `Done`.  It is very important that you use this exact name for the script to work.

#### 9. Define a named cell range for `orderNum`.
Right click on the cell that contains the customer order number and select `Define named range...`.

#### 10. Save the cell range as `orderNum`.
Change the name of the cell range to `orderNum` and click `Done`.  It is very important that you use this exact name for the script to work.

## Create a "Confirm Order" button.
A button needs to be available in the order sheet to allow users to execute the confirm order logic.

#### 11. Open the drawing dialog.
Select `Insert > Drawing...`.

![insert-drawing](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/7-insert-drawing.png)

#### 12. Draw a button.
Draw the button to look however you want.  After saving and closing the drawing, move it into position on the sheet.  It canont be embedded in a cell.

![draw-button](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/8-draw-button.png)

#### 13. Open the button's assign script dialog.
Click on the arrow in the top-right corner of the button, then select `Assign script...` from the dropdown menu.

![assign-script](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/9-assign-script.png)

#### 14. Assign the script to the button.
Enter the function name `confirmOrder` into the dialog and select `OK`.  The `confirmOrder` function will be called every time the button is clicked, enabling the script to execute.

![script-function](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/10-script-function.png)



# GMail Command Listener Setup
These instructions describe how to setup the production GMail account to respond to command emails.  Command emails are sent when a client account attempts to execute logic on the production account, since they do not have direct access.

## Create a filter for command emails.
The email command listener script only inspects emails with the `command` label.  When they are finished, the script re-labels them as `executed`.  A filter will automatically label command emails so that the script can identify them.  It will also archive the emails to keep the inbox clear.

#### 1. Open the GMail settings page.
On the production GMail account, click on the gear button and choose `Settings` from the dropdown menu.

![gmail-settings](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/11-gmail-settings.png)

#### 2. Navigate to the filters settings page.
Click the `Filters` tab link at the top of the settings page.

![filters-settings](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/12-filters-settings.png)

#### 3. Create a new filter.
Click the `Create a new filter` button at the bottom-right corner of the page.

![create-new-filter](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/13-create-new-filter.png)

#### 4. Set the filter subject.
Enter `CMD - ` into the `Subject` field of the filter creation dialog.  This will cause the filter to only select emails with "CMD - " in the subject.

![filter-subject](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/14-filter-subject.png)

#### 5. Confirm the filter.
Select `OK` from the filter creation confirmation dialog.

![confirm-create-filter](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/15-confirm-create-filter.png)

#### 6. Set the filter actions to archive and label as `command`.
Check the `Skip the Inbox (Archive it)` and `Apply the label:` options.  Choose `command` as the designated label.  If the command label doesn't exist, create a new one.  Be sure to use all lowercase letters.  When the correct actions have been selected, click the `Create filter` button.

![filter-actions](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/16-filter-actions.png)

## Create the script file with a time-driven trigger.
Unlike with the confirm order script, the email command listener script is not attached to any files.  Instead, a new script file needs to be created and a trigger needs to be set to activate the script on a regular basis.

#### 7. Create a new Google Apps Script file in Google Drive.
From Google Drive, select `New > More > Google Apps Script`.

![new-apps-script](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/17-new-apps-script.png)

#### 8. Rename the Apps Script project.
Change the name of the project to something descriptive, such as "Email Command Listener".

![change-name](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/18-change-name.png)

#### 9. Copy the code into the project.
Copy the code from the [`emailCommandListener.gs` script on GitHub](https://github.com/CivBase/confirm-order/blob/master/src/emailCommandListener.gs) and paste it into the `Code.gs` script in the open project.

![copy-paste-code](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/19-copy-paste-code.png)

#### 10. Open the project triggers dialog.
Click on the clock button to open the project triggers dialog.

![project-triggers](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/20-project-triggers.png)

#### 11. Add a new trigger.
Click the `No triggers set up. Click here to add one now.` link to create a new trigger.

![add-new-trigger](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/21-add-new-trigger.png)

#### 12. Setup the trigger so that it executes the script once very minute.
Select `emailCommandListener` function to run as a `Time-driven` event on a `Minutes timer` `Every minute`.  This will enable the script to execute every minute.  Close the dialog with the `Save` button.

![trigger-settings](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/22-trigger-settings.png)

#### 13. Authorize the script.
Before the script can start executing, you must authorize the script by continuing through two more dialogs.

![authorization](https://raw.githubusercontent.com/CivBase/confirm-order/master/screenshots/23-authorization.png)

## Make sure everything is in place.
For the script to completely function, a few files and folders need to exist with the correct names in the production account's Google Drive.

#### 14. Ensure all necessary files and folders are in place.
The production Google Drive should include the documents `confirm order readme.txt`, `PR Template 模板`, and `Chipstar Schedule 订单进度` in the root directory as well as the folder `进行生产任务 Current Order`.  The `PR Template 模板` file must be a Google Sheet which contains the named cell range `orderFileId`.  The `Chipstar Schedule 订单进度` file must be a Google Sheet which contains the named cell range `orderFileIds`.
