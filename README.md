# Excel_Unlocker

Simple python script to unlock protected Excel files.
Each Excel file is a hidden .ZIP archive containing .XML files for each worksheet.
Within these .XML is a wheetProtection tag locked sheet and a workbookProtection tag
containing the hashed passwords etc.
Retrieving and cracking these passwords are of course possible, but by simply removing the
SheetProtection / workbookProtection tags, the workbook will become completely unlocked.

A temporary folder is created in the target files working directory. The target file is extraced 
to the temp folder, remove the SheetProtection / workbookProtection and rezip it into the
original file format.
