# Excel VBA – HR Web Data Entry Automation (20 Records at a Time)

## What is this?
This Excel VBA macro **automatically enters employee data from Excel into a web portal**.

Instead of manually copying and pasting values one by one, this macro:
- Reads employee details from Excel
- Switches to Google Chrome
- Types the data into the web form
- Moves between fields using the TAB key
- Repeats this for **20 employees at a time**

This is useful for **HR, admin, and data-entry tasks** where the same form must be filled repeatedly.

---

## What does the macro do?
When you run the macro:

1. It starts from the **currently selected row** in Excel
2. Reads data from **5 specific columns**
3. Switches to **Google Chrome**
4. Types the values into the web form in order
5. Presses **TAB** to move to the next field
6. Repeats this process for **up to 20 rows**

If it finds an empty row, it **automatically stops**.

---

## Excel Column Mapping (VERY IMPORTANT)

The macro reads data from the following Excel columns:

| Excel Column | Used For |
|-------------|---------|
| Column C | Field 1 (First input field on the website) |
| Column E | Field 2 |
| Column F | Field 3 |
| Column G | Field 4 |
| Column H | Field 5 |

**The order of columns must match the order of fields on the website form.**

---

## Requirements
- Windows PC
- Microsoft Excel (Macro-enabled)
- Google Chrome browser
- Web form that allows keyboard input and TAB navigation

---

## How to Use (Step-by-Step for Beginners)

### STEP 1: Prepare your Excel file
1. Open Excel
2. Enter employee data:
   - Column C → Field 1
   - Column E → Field 2
   - Column F → Field 3
   - Column G → Field 4
   - Column H → Field 5
3. Make sure data starts from a single row (example: row 2)

---

### STEP 2: Enable macros in Excel
1. Open Excel
2. Go to **File → Options → Trust Center**
3. Click **Trust Center Settings**
4. Open **Macro Settings**
5. Select **Enable all macros**
6. Click OK

---

### STEP 3: Add the VBA code
1. Press **ALT + F11**
2. Click **Insert → Module**
3. Paste the VBA code into the window
4. Save the file as **Excel Macro-Enabled Workbook (.xlsm)**

---

### STEP 4: Open the web form
1. Open **Google Chrome**
2. Navigate to the web portal form
3. Click inside the **first input field**
4. Make sure the cursor is blinking in the first field

⚠️ Do NOT click anywhere else after this.

---

### STEP 5: Select starting row in Excel
1. Go back to Excel
2. Click on the row where you want to start (example: row 2)
3. Make sure that row contains valid data

---

### STEP 6: Run the macro
1. Press **ALT + F8**
2. Select `Sending20Keys`
3. Click **Run**

The macro will:
- Switch to Chrome
- Start typing data automatically
- Fill up to **20 records**

⛔ Do NOT touch keyboard or mouse while it is running.

---

## ⏱️ Timing Notes
- The macro waits 1 second between each record
- If your system or website is slow, you may need to increase the delay

Example:
```vba
Application.Wait Now + TimeValue("0:00:02")
