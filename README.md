# Features and Flow

## GUI Layout
- **Main Window**:
  - `+` (Add) and `-` (Remove) buttons.
  - Log View to display scanned details.
  - Treeview for displaying the inventory.
- **Input Dialog** (for scanning serial numbers):
  - Automatically refreshes after each scan for seamless user interaction.
  - Closes when the "Finished" button is pressed.

---

## Excel Sheets
1. **Models Sheet**:
   - Stores models and associated serial numbers.
2. **Inventory Sheet**:
   - Logs scanned items, including the model and unique serial numbers.
3. **Timestamps Sheet**:
   - Logs actions with:
     - Timestamp
     - Model
     - Serial
     - Action type

---

## Logic
- `+` Button:
  - Triggers the input dialog.
  - First scanned serial looks up the **Models Sheet** to find the associated model.
  - Records the model in the **Inventory Sheet** and logs the action in the **Timestamps Sheet**.
  - Second scanned serial is added to the same inventory row and logged in **Timestamps**.
- `-` Button:
  - Triggers an input dialog to remove an item by its serial number.

---

## Functions
- Handle serial number scanning for both:
  1. Model lookup.
  2. Logging the unique serial number.
- Update Excel sheets and log actions.
- Manage the inventory and log views in real time.

---

## Error Handling
- Handles cases where scanned serials do not match entries in the **Models Sheet**.
- Ensures unique serial numbers are not duplicated.
