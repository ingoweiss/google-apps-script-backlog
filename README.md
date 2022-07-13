# Google Apps Script Backlog Utilities

A Google Doc is a very flexible format for creating and evolving an (initial) backlog.

However, one might want to use the backlog data in other tools to visualize (user story mapping etc.) or enrich (estimation, resourcing etc.) the backlog. It is cumbersome and error prone to keep multiple documents in sync.

This script adds menu commands to a Google Doc backlog that help with a multi-doc backlog workflow.

## Menu Commands:
### Backlog > Export Stories

1. Exports story data (ID and name) to the "Backlog Export" tab of a connected spreadsheet
2. Exports story data (ID and name) to a JSON file*

The following needs to be satisfied for the "Export Stories" command to work correctly:

1. The "Connect Spreadsheet" command needs to be run
2. The connected Google Sheet needs to have a "Backlog Export" tab with a header row ("ID" and "Name")
3. Story titles need to be formated as a level 3 heading
4. Story titles need to be in the format "[PREFIX]-[NUMBER]: [Name]" (Example: "ACCOUNT-01: User Logs In")

*The JSON file is created using the base name of the Google Doc, at the top level Google Drive. After the first export, the export file can be moved into any folder and the script should continue to work

### Backlog > Connect Spreadsheet

1. Asks for the ID of the Google Sheet that should be used to export data to

### Backlog > Open Connected Spreadsheet

1. Opens the connected Google Sheet