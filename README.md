This code is Google Apps Script code intended to be used with a spreadsheet.
It helps you create workflows for things like homework trackers where we want
to have a spreadsheet template copied into a new sheet for each day, so we can
put each day's homework on a dated tab.

The sheet allows the user to copy tabs, delete tabs en masse, and to automate
adding tabs each day. We can either add a tab each day at a given time, or we
can work in a more complex manner by e.g. adding 5 tabs for the full week each
Monday, or by Adding a tab for Monday each Friday.

To use:

1. Create a new spreadsheet 
2. Create a TEMPLATE tab with the template you would like copied for each day.
3. Open Google Apps Script (Extensions menu => Apps Script)
4. Paste the contents of Code.gs into the file.
5. Save the script.
6. Reload the spreadsheet so that the "Date Tabs" menu appears.
7. Use the menu items to run the program.