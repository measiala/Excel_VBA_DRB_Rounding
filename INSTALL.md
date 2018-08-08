# Installing Excel DRB Rounding Add-in

To install, perform the following steps in order:

1. Download the Excel Add-in from this repository to a location that is accessible to you (e.g., H: drive)
2. Open Excel and create a new workbook (this just activates the top ribbon)
3. In the "Developer" Tab, locate and click on the "Excel Add-ins" button.
4. A window called "Add-ins" should pop up. Click on the "Browse" button.
5. An Open dialog window will pop up. Browse to the location where you saved the Add-in
   file in step 1. Select the file and click "Open".
6. A dialog box will appear asking if you'd like to copy this to your Add-ins folder.
   Click yes.
7. Once complete, you should see an Add-in called "DRB Rounding Macro" appear in the list.
   Check the box next to where you see the Add-in listed. The rounding macros are now
   accessible to you from within any workbook.

To make it easier to use the macros, I suggest the following to add the macros to your ribbon.

1. In the "File" tab, locate and click on "Options" all the way towards the bottom. 
   A window called "Excel Options" will appear.
2. Locate and click on "Customize Ribbon" on the left sidebar. This will open the ribbon
   customization pane.
3. On the right half of the pane (under "Customize the Ribbon"), locate and click on the
   "New Tab" button. This will create a new tab called "New Tab (custom)" that contains a group
   called "New Group (custom)".
4. Click on the "New Tab (custom)" and the click on the "Rename" button towards the bottom.
   This will pop up a dialog asking for the "Display Name" for the tab. Type in your preferred name.
   For the purposes of these instructions, I'm going to call it "DRB". Click "OK" when
   you have entered the name you like.
5. Click on "New Group (custom)" and repeat step 4 to rename the group to whatever you'd like. You
   can always change it later if you change your mind. I'm going to call it "Rounding"
6. The last steps is to add the macros themselves to the ribbon at the location you just created.
   If the group you just renamed in step 5 (my "Rounding" group) is not currently selected, select it now.
7. In the pop up list on the left column towards the top, just beneath "Choose commands from", select
   "Macros" in place of the default "Popular Commands". This should give you a list of all macros 
   available to you including the 3 macros that begin with "DRB_".
8. Select "DRB_Round_Count" and then click on the "Add >>" button in the center of the pane. Repeat for
   "DRB_Round_Estimate" and "DRB_Restore_Selection". Once complete these 3 macros should all appear
   under the custom group you created above (my "Rounding" group).
9. (optional) If you want to shorten the names of the entries, just select them and click the "Rename"
   button. For example, you could select "DRB_Round_Count", click on the "Rename" button, and rename it
   to "Counts". I have renamed mine to "Round Counts", "Round Estimates", and "Restore Original".

You are done! These settings should remain from session to session so you only need to do them once. DRB rounding is now just a click away!!