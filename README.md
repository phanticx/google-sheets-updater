# google-sheets-updater

Basic GSheets updater written in Java for use with [a classmate's attendance app](https://bepresent.app/) to automate updates for Computer Club.

Reads excel file using Apache POI and GSheets using Google's API, appends selected sheet with new data.
To use with your own file format, change [Main.readExcelFile()](https://github.com/phanticx/google-sheets-updater/blob/main/src/main/java/com/phanticx/Main.java#L47) to your liking. 

## Usage
1. Clone git repository
2. Set up Google Cloud OAuth and environment [here](https://developers.google.com/sheets/api/quickstart/java).
    * Add the path to your credentials.json file to [Main.credentialsPath](https://github.com/phanticx/google-sheets-updater/blob/main/src/main/java/com/phanticx/Main.java#L32)
3. Download your excel file in .xlsx format
    * Add the path to your excel file to [Main.excelFilePath](https://github.com/phanticx/google-sheets-updater/blob/main/src/main/java/com/phanticx/Main.java#L33)
4. Set your spreadsheet ID and sheet name at [Main.spreadsheetId](https://github.com/phanticx/google-sheets-updater/blob/main/src/main/java/com/phanticx/Main.java#L35) and [Main.range](https://github.com/phanticx/google-sheets-updater/blob/main/src/main/java/com/phanticx/Main.java#L36), respectively
5. Run Main.main()
