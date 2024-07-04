### Steps to Add an Apps Script to Google Sheets

1. **Open Your Google Sheet:**
   - Go to [Google Sheets](https://sheets.google.com) and open the spreadsheet where you want to add the script.

2. **Open the Script Editor:**
   - In the top menu, click on `Extensions` > `Apps Script`.

3. **Create a New Script Project:**
   - A new tab will open, taking you to the Google Apps Script editor.
   - By default, a new script project will be created with a file named `Code.gs`.

4. **Write Your Script:**
   - In the `Code.gs` file, write or paste the script you want to use.

5. **Save the Script:**
   - Click the floppy disk icon or go to `File` > `Save` to save your script.

6. **Authorize the Script (if needed):**
   - If your script requires authorization, a dialog will appear the first time you run it. Click `Review Permissions`, sign in to your Google account, and click `Allow`.

7. **Run the Script:**
   - In the Script Editor, click the play button (triangle icon) to run your script.
   - You can also assign the script to a menu item or a button in your Google Sheet for easier access.

8. **Close the Script Editor:**
   - Once you're done, you can close the Script Editor tab and return to your Google Sheet.

### Assigning a Script to a Custom Menu (Optional)

1. **Add a Custom Menu:**
   - To create a custom menu that runs your script, add the following code to your `Code.gs` file:

     ```javascript
     function onOpen() {
       var ui = SpreadsheetApp.getUi();
       ui.createMenu('Custom Menu')
           .addItem('Run Script', 'copyNamedRangeToNewSpreadsheet')
           .addToUi();
     }
     ```

2. **Save and Refresh:**
   - Save the script and refresh your Google Sheet. You will see a new menu called `Custom Menu` in the top menu bar, which will run your script when clicked.

By following these steps, you can add and run a custom Apps Script in your Google Sheets.

### Sheets

1. **create-performance-evaluation-sheet.js**
   
<img width="1144" alt="image" src="https://github.com/ZeeshanAhmadKhalil/google-docs-app-scripts/assets/41861952/abb5e647-ef42-4f5c-b438-72970218aaca">
