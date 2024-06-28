```markdown
# Google Calendar Time Tracker

Easily track your time spent on different categories of events in your Google Calendar with this Google Apps Script project.

## Features:

- **Event Classification:** Automatically classify calendar events based on their color and a customizable category table.
- **Work Hours Calculation:** Calculate the hours dedicated to each category within a specified date range.
- **Data Export to Spreadsheet:** Export your time tracking data to a Google Sheets spreadsheet.
- **User-Friendly Interface:** Provides a sidebar in Google Sheets for easy setup and data updates.

## Installation:

1. **Copy the Folder:** Make a copy of this project folder to your Google Drive.
2. **Open the Spreadsheet:** Open the `[Spreadsheet File Name].gsheet` file.
3. **Authorize the Script:** The first time you run the script, you'll need to authorize it to access your calendar and spreadsheet.
4. **Set Up Categories:** Use the sidebar to access and customize the event category color table.

## Usage:

1. **Open the Sidebar:** In the spreadsheet, go to "Add-ons" > "Google Calendar Time Tracker" > "Show Sidebar".
2. **Select Date Range:** In the sidebar, select the date range for which you want to track your time.
3. **Update Spreadsheet:** Click the "Update Sheet" button to process your calendar data and update the spreadsheet.

## Customization:

- **Color Table:** Customize event categories and their corresponding colors in the color table accessible from the sidebar.
- **Destination Spreadsheet:** Change the target spreadsheet and sheet by modifying the `spreadsheetId` and `sheetName` variables in the `[Code File Name].gs` file.

## Notes:

- This script requires access to your Google Calendar and Google Sheets.
- Ensure that the "Working Hours" event is correctly set up in your calendar for accurate time tracking.

## Contributions:

Contributions are welcome! Please fork the repository and submit a pull request with your changes.
