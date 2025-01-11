# networks_assets
Python tkinter app which takes model &amp; unique serials as scanned input and translates to readable text in an inventory spreadsheet.

# Features and Flow
GUI Layout:

# Main window with:
+ (Add) and - (Remove) buttons.
Log View to display scanned details.
Treeview for Inventory display.
Input dialog for scanning serial numbers, which will:
Auto-refresh after each scan for seamless user interaction.
Close when "Finished" is pressed.

#Excel Sheets:

Models sheet for model and associated serial numbers.
Inventory sheet to log scanned items (model and unique serial number).
Timestamps sheet to log actions with timestamp, model, serial, and action type.

#Logic:

+ triggers the input dialog.
First scanned serial looks up Models and finds the associated model.
Model is recorded in Inventory and action logged in Timestamps.
Second scanned serial is logged in the Inventory row and Timestamps.
- will trigger an input dialog to remove an item by its serial number.

#Functions:

Handle serial number scanning for both steps (model and unique serial).
Update the sheets and log actions.
Manage the inventory and logviews in real time.

#Error Handling:

Handle cases where scanned serials don't match the Models sheet.
Ensure unique serial numbers are not duplicated.
