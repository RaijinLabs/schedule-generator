# Schedule-Generator

Schedule Generator is a script for use in Google Sheets written in Google Script.
Its primary function is the formatting of a plain weekly schedule into visible and distinct schedule blocks.

# Setup
### Initial
  - This program requires two sheets in the spreadsheet that it will be running on
  - Name the first sheet `Class Data` and the second sheet `Schedule`
  - Go to `Tools` > `Script Editor` in the toolbar for the spreadsheet
  - Paste the code into the script editor
### Setting Up Class Data
  - Replicate this top row on the `Class Data` sheet
  
| Class Name | Room Label | Days(MTWThFS) |Time In|Time Out|
| ------ | ------ | ------ |------ |------ |
  - Select a cell for a button, then go to `Insert` > `Drawing`
  - Pick any shape from the drawing screen and customize however you like. Once finished hit `Save & Close`
  - Click on your new button and along with the blue outline along the border, a column on three dots should appear on the upper right hand corner of the shape.
  - Click the set of three dots and click `Assign script...`
  - Type `readAndMakeSchedule` on the textbox that appears. Click `OK`
##### And that's it!

# How To Use
 - Enter schedule data under class data (ie; Activity name under class name, Activity Location under Room Label, etc)
 - Hit the button
 - **Done!**
 - The schedule will pop up in the `Schedule` sheet and it can be safely copied into a new Google Sheet for safekeeping.
 
# License 
 - GPL v3.0

 
