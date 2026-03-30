# Integrated Poultry Shop POS (v2.0)

This document provides an implementation blueprint for an **AppSheet + Google Sheets** system tailored to poultry operations with fractional retailing, mixed payment modes, and rapid admin-driven price changes.

---

## 1) Google Sheets / AppSheet Data Model

Create one Google Sheet file with four tabs (tables), then add them in AppSheet.

## Table A: `Inventory` (Master List)

**Columns**
- `Item_ID` (Text, key, `UNIQUEID()`)
- `Item_Name` (Text)
- `Item_Code` (Enum: SP, FP, AB, etc.)
- `Category` (Enum: Birds, Feeds, Medication, Equipment)
- `Base_Unit` (Enum, default `Bag`)
- `Bag_Weight_Kg` (Decimal, default `25`)  
  > Needed for Kilo conversions.
- `Cost_Price` (Price)
- `Selling_Price` (Price)
- `Current_Stock_Decimal` (Decimal)
- `Reorder_Level` (Decimal)
- `Last_Updated_By` (Email)
- `Last_Updated_At` (DateTime)

**Behavior**
- `Selling_Price` must be editable in an admin-only interface.
- `Current_Stock_Decimal` supports fractional values (e.g., 12.75 bags).

## Table B: `Sales` (POS Transactions)

**Columns**
- `Sale_ID` (Text, key, `UNIQUEID()`)
- `Date` (DateTime, `NOW()`)
- `Item_Link` (Ref → `Inventory`)
- `Unit_Type` (Enum: Bag, Half, Quarter, Kilo, Decimal)
- `Qty_Value` (Decimal)
- `Payment_Method` (Enum: Cash, POS, Transfer)
- `Unit_Factor` (Virtual column, Decimal)
- `Inventory_Deduction` (Virtual column, Decimal)
- `Total_Amount` (Virtual column, Price)
- `Attendant_Email` (Email, `USEREMAIL()`)
- `Commission_Earned` (Price or Decimal)

## Table C: `Bookings_DOB` (Day-Old Birds)

**Columns**
- `Booking_ID` (Text, key, `UNIQUEID()`)
- `Date_Receipt` (DateTime)
- `Customer_Name` (Text)
- `Phone_No` (Phone)
- `Bird_Type` (Enum: BX, CX, TX, F.TX, NX)
- `Qty` (Number)
- `Deposit` (Price)
- `Balance` (Price)
- `Expected_Date` (Date)
- `Remarks` (LongText)
- `Status` (Enum: Pending, Confirmed, Collected, Cancelled)

## Table D: `Daily_Stock_Check`

**Columns**
- `Check_ID` (Text, key, `UNIQUEID()`)
- `Date` (Date)
- `Item_Link` (Ref → `Inventory`)
- `Opening_Stock` (Decimal)
- `Closing_Stock` (Decimal)
- `Variance` (Virtual/App formula)

**Variance formula**
```appsheet
[Closing_Stock] - [Opening_Stock]
```

---

## 2A) AppSheet Unit Conversion Logic

Use a reusable factor that maps each sales unit into bag-equivalent quantity.

## `Unit_Factor` (Virtual Column in `Sales`)

```appsheet
SWITCH(
  [Unit_Type],
  "Bag", 1,
  "Half", 0.5,
  "Quarter", 0.25,
  "Kilo", 1 / IF(ISBLANK([Item_Link].[Bag_Weight_Kg]), 25, [Item_Link].[Bag_Weight_Kg]),
  "Decimal", 1,
  1
)
```

> For `Decimal`, staff enters exact bag-equivalent in `Qty_Value` (e.g., `0.30`).

## `Inventory_Deduction` (Virtual Column in `Sales`)

```appsheet
[Qty_Value] * [Unit_Factor]
```

### Example validation
- Half bag sale (`Qty_Value = 1`, `Unit_Type = Half`) ⇒ deduction `0.5`
- Quarter (`Qty_Value = 2`, `Unit_Type = Quarter`) ⇒ deduction `0.5`
- Kilo (`Qty_Value = 10`, `Bag_Weight_Kg = 25`) ⇒ deduction `10 * (1/25) = 0.4`

## `Total_Amount` (Virtual Column in `Sales`)

```appsheet
[Inventory_Deduction] * [Item_Link].[Selling_Price]
```

## Inventory update action (important)
Create an action on `Inventory` table:
- **Action name:** `Deduct_From_Inventory`
- **Type:** Data: set the values of some columns in this row
- **Set this column:** `Current_Stock_Decimal`
- **Formula:**

```appsheet
[Current_Stock_Decimal] - ANY(SELECT(Sales[Inventory_Deduction], [Sale_ID] = [_THISROW-1].[Sale_ID]))
```

Preferred approach: Trigger this via a Bot on new Sales rows using “Referenced action.”

---

## 2B) Admin “Market Update” Feature (Bulk Price Update)

Goal: update all **Feeds** prices in one operation.

## Option 1 (Recommended): Percentage markup action in AppSheet

### Step 1: Settings table
Create table `Admin_Settings`:
- `Setting_ID` (key)
- `Feed_Price_Factor` (Decimal, e.g., `1.08` for +8%)
- `Updated_By`
- `Updated_At`

### Step 2: Action on `Inventory`
Create action `Apply_Feed_Market_Update`:
- For records in `Inventory`
- **Only if this condition is true:**
```appsheet
[Category] = "Feeds"
```
- **Set column:** `Selling_Price`
- **Formula:**
```appsheet
ROUND([Selling_Price] * ANY(Admin_Settings[Feed_Price_Factor]), 0.01)
```

### Step 3: Group action
Create an “execute action on a set of rows” action that targets all `Inventory` rows where Category is Feeds.

Row filter for target set:
```appsheet
SELECT(Inventory[Item_ID], [Category] = "Feeds")
```

### Step 4: Admin-only visibility
In the action/view security expression:
```appsheet
IN(USEREMAIL(), {"admin@yourshop.com", "owner@yourshop.com"})
```

This gives one-tap bulk feed repricing.

## Option 2: CSV import overwrite
Maintain a “Price Upload” sheet and import updated feed prices in bulk. Good for external accountant workflows, but less controlled than in-app action.

---

## 2C) SMS Automation for `Bookings_DOB` (Google Apps Script)

Below is a production-ready starter script. It triggers on edits, checks deposit > 0, and sends SMS once per booking using a `SMS_Sent` column.

> **Prerequisite columns in `Bookings_DOB`:** add `SMS_Sent` (TRUE/FALSE), default FALSE.

```javascript
/**
 * Trigger: Installable onEdit trigger
 * Sheet: Bookings_DOB
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Bookings_DOB') return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = e.range.getRow();
  if (row === 1) return;

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const idx = (name) => headers.indexOf(name);

  const phoneIdx = idx('Phone_No');
  const qtyIdx = idx('Qty');
  const birdIdx = idx('Bird_Type');
  const depIdx = idx('Deposit');
  const balIdx = idx('Balance');
  const sentIdx = idx('SMS_Sent');

  if ([phoneIdx, qtyIdx, birdIdx, depIdx, balIdx, sentIdx].some(i => i === -1)) {
    throw new Error('Missing required columns in Bookings_DOB sheet.');
  }

  const phone = String(rowValues[phoneIdx] || '').trim();
  const qty = rowValues[qtyIdx] || 0;
  const birdType = String(rowValues[birdIdx] || '').trim();
  const deposit = Number(rowValues[depIdx] || 0);
  const balance = Number(rowValues[balIdx] || 0);
  const smsSent = rowValues[sentIdx] === true;

  if (!phone || deposit <= 0 || smsSent) return;

  const message = `Booking Confirmed: ${qty} ${birdType} chicks. Deposit Recieved: ${deposit}. Bal: ${balance}. Please keep your receipt!`;

  // Replace with your SMS provider function.
  const ok = sendSmsWithTermii(phone, message);

  if (ok) {
    sheet.getRange(row, sentIdx + 1).setValue(true);
  }
}

/**
 * Example SMS via Termii API
 * Set script properties: TERMII_API_KEY, TERMII_SENDER_ID
 */
function sendSmsWithTermii(to, message) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('TERMII_API_KEY');
  const sender = PropertiesService.getScriptProperties().getProperty('TERMII_SENDER_ID') || 'POULTRY';
  if (!apiKey) throw new Error('TERMII_API_KEY not set in Script Properties.');

  const payload = {
    to: to,
    from: sender,
    sms: message,
    type: 'plain',
    channel: 'generic',
    api_key: apiKey
  };

  const res = UrlFetchApp.fetch('https://api.ng.termii.com/api/sms/send', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code >= 200 && code < 300) return true;

  Logger.log('SMS failed: ' + code + ' ' + res.getContentText());
  return false;
}
```

### Setup steps
1. Open Google Sheet → Extensions → Apps Script.
2. Paste script.
3. Add script properties (`TERMII_API_KEY`, optional `TERMII_SENDER_ID`).
4. Create **Installable Trigger** for `onEdit`.
5. Test by entering Deposit > 0 in a `Bookings_DOB` row.

---

## 2D) Looker Studio Dashboard: Payment Mode Pie Chart

Create a pie chart for **Cash vs POS vs Transfer**.

1. Open Looker Studio → Create report.
2. Connect data source to Google Sheet containing `Sales` tab.
3. Ensure field types:
   - `Date` = Date/DateTime
   - `Payment_Method` = Text/Dimension
   - `Total_Amount` = Currency/Metric
4. Insert → **Pie chart**.
5. Set:
   - **Dimension:** `Payment_Method`
   - **Metric:** `SUM(Total_Amount)`
6. Add Date range control for daily/weekly/monthly reconciliation.
7. Optional filter: include only completed/valid sales if status column exists.
8. Style: show percentage + value labels for quick till-vs-bank analysis.

---

## 3) AppSheet Dropdown Master Values

Use these in Enum/Valid_If expressions.

## Birds (`Bird_Type` / relevant product columns)
- Broilers (BX)
- Cockerels (CX)
- Turkeys (TX)
- Foreign Turkeys (F.TX)
- Noilers (NX)

## Feeds
- Starter Plus (SP)
- Finisher Plus (FT)
- Starter Crumble (SC)
- Finisher Pellet (FP)
- Grower (G)
- Layer (L)

## Medication
- Antibiotics (AB)
- Anticocci (AC)
- Fowl Pox (FP)
- Tylo (T)

## Equipment
- Drinkers (Small, Medium, Large, 3in1)
- Feeders (Wooden Size 1-4, Tinker Size 1-3)
- Stands (Wooden, Tinker)

---

## Recommended Governance & Controls

- Restrict all pricing actions and Inventory manual edits to admin emails.
- Enable audit fields (`Updated_By`, `Updated_At`) in all mutable tables.
- Add stock floor validation to stop negative inventory:
```appsheet
[Item_Link].[Current_Stock_Decimal] >= [Inventory_Deduction]
```
- Add shift-level reports using `Daily_Stock_Check` variance and attendant-level sales totals.

