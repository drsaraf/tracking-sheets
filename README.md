# Tracking Status Lookup for Google Sheets

Track packages directly in your Google Sheets spreadsheet. Auto-detects carrier (USPS, UPS, FedEx, DHL) and fetches real-time status.

## Installation

### Step 1: Open Apps Script
1. Open your Google Sheet
2. Go to **Extensions → Apps Script**

### Step 2: Add the Code
1. Delete any existing code in the editor
2. Copy the entire contents of `Code.gs` and paste it
3. Click **Save** (💾 icon or Ctrl+S)

### Step 3: Configure (if needed)
Edit the `CONFIG` object at the top if your columns are different:
```javascript
const CONFIG = {
  TRACKING_COLUMN: 4,  // Column D
  STATUS_COLUMN: 5,    // Column E  
  START_ROW: 2,        // Skip header
  SHEET_NAME: null     // Active sheet
};
```

### Step 4: Authorize
1. Close the Apps Script editor
2. Reload your spreadsheet
3. Click **📦 Tracking → Update All Status**
4. Click through the authorization prompts (first time only)

## Usage

### Menu Options

After installation, you'll see a **📦 Tracking** menu:

- **Update All Status** - Scans all tracking numbers, updates status column
- **Update Selected Rows** - Only updates rows you've selected
- **Settings** - View configuration and available functions

### Custom Functions

Use these formulas directly in cells:

| Function | Example | Returns |
|----------|---------|---------|
| `=TRACKSTATUS(D2)` | Status text | "Delivered", "In Transit", etc. |
| `=TRACKURL(D2)` | Tracking URL | Full carrier tracking link |
| `=CARRIER(D2)` | Carrier name | "USPS", "UPS", "FedEx", "DHL" |

### Automatic Features

- ✅ **Skips delivered** - Already delivered items won't be re-checked
- 🔗 **Clickable links** - Status becomes a hyperlink to full tracking page
- 🚀 **Auto-detect carrier** - Identifies carrier from tracking number format

## Supported Carriers

| Carrier | Tracking Format |
|---------|-----------------|
| USPS | 20-22 digits (starts with 94, 93, 92, etc.) |
| UPS | 1Z + 16 alphanumeric characters |
| FedEx | 12, 15, or 20-22 digits |
| DHL | 10-11 digits or starts with JJD |

## Status Values

- **Delivered** - Package delivered (with date if available)
- **Out for Delivery** - On truck for delivery today
- **In Transit** - Moving through carrier network
- **Shipped / Picked Up** - Carrier has package
- **Label Created** - Shipping label created, not yet shipped
- **Check Link** - Status unclear, click link to check manually
- **Error** - Could not fetch status

## Rate Limiting

The script adds small delays between requests to avoid being blocked by carrier websites. For large batches (100+ tracking numbers), run in smaller groups.

## Troubleshooting

**"Authorization required"**
- Click through the Google authorization prompts
- You may need to click "Advanced" → "Go to [script name]" for first-time auth

**Status shows "Error" or "Check Link"**
- Carrier website may be blocking automated requests
- Click the link to check status manually
- Try again later

**Menu doesn't appear**
- Reload the spreadsheet
- Make sure you saved the Apps Script code
