/**
 * Tracking Status Lookup for Google Sheets
 * Automatically fetches shipping status for USPS, UPS, FedEx, DHL
 */

// Configuration - adjust these to match your sheet
const CONFIG = {
  TRACKING_COLUMN: 4,  // Column D (1-indexed)
  STATUS_COLUMN: 5,    // Column E (1-indexed)
  START_ROW: 2,        // Skip header row
  SHEET_NAME: null     // null = active sheet, or specify name like "Sheet1"
};

// Standardized status values
const STATUS = {
  DELIVERED: 'Delivered',
  OUT_FOR_DELIVERY: 'Out for Delivery',
  IN_TRANSIT: 'In Transit',
  EXCEPTION: 'Exception',
  WAITING_PICKUP: 'Waiting for Pickup',
  LABEL_CREATED: 'Label Created',
  NOT_FOUND: 'Not Found'
};

/**
 * Adds custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Tracking')
    .addItem('Update All Status', 'updateAllTrackingStatus')
    .addItem('Update Selected Rows', 'updateSelectedTrackingStatus')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Detect carrier from tracking number format
 */
function detectCarrier(trackingNumber) {
  const tn = trackingNumber.toString().trim().toUpperCase().replace(/\s/g, '');
  
  // UPS: 1Z + 16 alphanumeric
  if (/^1Z[A-Z0-9]{16}$/i.test(tn)) {
    return { carrier: 'UPS', url: `https://www.ups.com/track?tracknum=${tn}` };
  }
  
  // FedEx: 12, 15, or 20-22 digits
  if (/^\d{12}$/.test(tn) || /^\d{15}$/.test(tn)) {
    return { carrier: 'FedEx', url: `https://www.fedex.com/fedextrack/?trknbr=${tn}` };
  }
  
  // DHL: 10-11 digits or starts with JJD
  if (/^\d{10,11}$/.test(tn) || /^JJD/i.test(tn)) {
    return { carrier: 'DHL', url: `https://www.dhl.com/us-en/home/tracking/tracking-express.html?submit=1&tracking-id=${tn}` };
  }
  
  // USPS: 20-22 digits, often starts with 94, 93, 92, 91, 70, 23, 03
  if (/^(94|93|92|91|70|23|03)\d{18,20}$/.test(tn) || /^\d{20,22}$/.test(tn)) {
    return { carrier: 'USPS', url: `https://tools.usps.com/go/TrackConfirmAction?tLabels=${tn}` };
  }
  
  // USPS: 13 characters international
  if (/^[A-Z]{2}\d{9}[A-Z]{2}$/i.test(tn)) {
    return { carrier: 'USPS', url: `https://tools.usps.com/go/TrackConfirmAction?tLabels=${tn}` };
  }
  
  return { carrier: 'Unknown', url: null };
}

/**
 * Fetch tracking status from carrier API/website
 */
function fetchTrackingStatus(trackingNumber) {
  const { carrier, url } = detectCarrier(trackingNumber);
  
  if (carrier === 'Unknown' || !url) {
    return { status: STATUS.NOT_FOUND, url: null, carrier };
  }
  
  try {
    let status = STATUS.NOT_FOUND;
    
    if (carrier === 'USPS') {
      status = fetchUSPSStatus(trackingNumber);
    } else if (carrier === 'UPS') {
      status = fetchUPSStatus(trackingNumber);
    } else if (carrier === 'FedEx') {
      status = fetchFedExStatus(trackingNumber);
    } else if (carrier === 'DHL') {
      status = fetchDHLStatus(trackingNumber);
    }
    
    return { status, url, carrier };
  } catch (e) {
    Logger.log('Error fetching status: ' + e.toString());
    return { status: STATUS.EXCEPTION, url, carrier };
  }
}

/**
 * USPS Status - using their tracking page
 */
function fetchUSPSStatus(trackingNumber) {
  try {
    const url = `https://tools.usps.com/go/TrackConfirmAction?tLabels=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    const html = response.getContentText();
    
    // Delivered
    if (/class="[^"]*delivered[^"]*"/i.test(html) || /Delivered,/i.test(html) || />Delivered</i.test(html)) {
      return STATUS.DELIVERED;
    }
    
    // Out for Delivery
    if (/Out for Delivery/i.test(html)) {
      return STATUS.OUT_FOR_DELIVERY;
    }
    
    // Exception / Alert
    if (/Alert/i.test(html) || /Exception/i.test(html) || /Delivery Attempt/i.test(html) || /No Access/i.test(html) || /Undeliverable/i.test(html)) {
      return STATUS.EXCEPTION;
    }
    
    // Waiting for Pickup
    if (/Available for Pickup/i.test(html) || /Ready for Pickup/i.test(html) || /Held at Post Office/i.test(html)) {
      return STATUS.WAITING_PICKUP;
    }
    
    // In Transit
    if (/In Transit/i.test(html) || /Arrived at/i.test(html) || /Departed/i.test(html) || /Processed/i.test(html) || /In-Transit/i.test(html) || /Origin/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    // Shipped / Accepted
    if (/Accepted/i.test(html) || /USPS in possession/i.test(html) || /Picked Up/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    // Label Created
    if (/Pre-Shipment/i.test(html) || /Label Created/i.test(html) || /Shipping Label Created/i.test(html)) {
      return STATUS.LABEL_CREATED;
    }
    
    // Not Found
    if (/Status Not Available/i.test(html) || /not found/i.test(html) || /no record/i.test(html)) {
      return STATUS.NOT_FOUND;
    }
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('USPS error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * UPS Status
 */
function fetchUPSStatus(trackingNumber) {
  try {
    const url = `https://www.ups.com/track?tracknum=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    const html = response.getContentText();
    
    // Delivered
    if (/delivered/i.test(html) && /status/i.test(html)) {
      return STATUS.DELIVERED;
    }
    
    // Out for Delivery
    if (/out for delivery/i.test(html)) {
      return STATUS.OUT_FOR_DELIVERY;
    }
    
    // Exception
    if (/exception/i.test(html) || /delivery attempt/i.test(html)) {
      return STATUS.EXCEPTION;
    }
    
    // Waiting for Pickup
    if (/ready for pickup/i.test(html) || /will call/i.test(html) || /access point/i.test(html)) {
      return STATUS.WAITING_PICKUP;
    }
    
    // In Transit
    if (/in transit/i.test(html) || /on the way/i.test(html) || /departed/i.test(html) || /arrived/i.test(html) || /origin scan/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    // Label Created
    if (/label created/i.test(html) || /order processed/i.test(html) || /shipment ready/i.test(html)) {
      return STATUS.LABEL_CREATED;
    }
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('UPS error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * FedEx Status
 */
function fetchFedExStatus(trackingNumber) {
  try {
    const url = `https://www.fedex.com/fedextrack/?trknbr=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    const html = response.getContentText();
    
    // Delivered
    if (/delivered/i.test(html)) {
      return STATUS.DELIVERED;
    }
    
    // Out for Delivery
    if (/on fedex vehicle/i.test(html) || /out for delivery/i.test(html)) {
      return STATUS.OUT_FOR_DELIVERY;
    }
    
    // Exception
    if (/exception/i.test(html) || /delivery exception/i.test(html) || /delay/i.test(html)) {
      return STATUS.EXCEPTION;
    }
    
    // Waiting for Pickup
    if (/ready for pickup/i.test(html) || /at fedex/i.test(html) || /hold at/i.test(html)) {
      return STATUS.WAITING_PICKUP;
    }
    
    // In Transit
    if (/in transit/i.test(html) || /departed/i.test(html) || /arrived/i.test(html) || /at local/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    // Label Created
    if (/shipment information sent/i.test(html) || /label created/i.test(html)) {
      return STATUS.LABEL_CREATED;
    }
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('FedEx error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * DHL Status
 */
function fetchDHLStatus(trackingNumber) {
  try {
    const url = `https://www.dhl.com/us-en/home/tracking/tracking-express.html?submit=1&tracking-id=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    const html = response.getContentText();
    
    // Delivered
    if (/delivered/i.test(html)) {
      return STATUS.DELIVERED;
    }
    
    // Out for Delivery
    if (/with delivery courier/i.test(html) || /out for delivery/i.test(html)) {
      return STATUS.OUT_FOR_DELIVERY;
    }
    
    // Exception
    if (/exception/i.test(html) || /shipment on hold/i.test(html)) {
      return STATUS.EXCEPTION;
    }
    
    // Waiting for Pickup
    if (/available for pickup/i.test(html) || /at service point/i.test(html)) {
      return STATUS.WAITING_PICKUP;
    }
    
    // In Transit
    if (/in transit/i.test(html) || /departed/i.test(html) || /arrived/i.test(html) || /processed/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    // Label Created
    if (/shipment picked up/i.test(html)) {
      return STATUS.IN_TRANSIT;
    }
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('DHL error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Update all tracking statuses in the sheet
 */
function updateAllTrackingStatus() {
  const sheet = CONFIG.SHEET_NAME 
    ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME)
    : SpreadsheetApp.getActiveSheet();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No data found in sheet.');
    return;
  }
  
  const trackingRange = sheet.getRange(CONFIG.START_ROW, CONFIG.TRACKING_COLUMN, lastRow - CONFIG.START_ROW + 1, 1);
  const statusRange = sheet.getRange(CONFIG.START_ROW, CONFIG.STATUS_COLUMN, lastRow - CONFIG.START_ROW + 1, 1);
  
  const trackingNumbers = trackingRange.getValues();
  const currentStatuses = statusRange.getValues();
  
  let updated = 0;
  let skipped = 0;
  
  for (let i = 0; i < trackingNumbers.length; i++) {
    const tn = trackingNumbers[i][0];
    const currentStatus = currentStatuses[i][0].toString().toLowerCase();
    
    // Skip empty cells
    if (!tn || tn.toString().trim() === '') {
      continue;
    }
    
    // Skip already delivered
    if (currentStatus.includes('delivered')) {
      skipped++;
      continue;
    }
    
    // Fetch new status
    const result = fetchTrackingStatus(tn.toString());
    
    // Create HYPERLINK formula
    const cell = sheet.getRange(CONFIG.START_ROW + i, CONFIG.STATUS_COLUMN);
    if (result.url) {
      cell.setFormula(`=HYPERLINK("${result.url}","${result.status}")`);
    } else {
      cell.setValue(result.status);
    }
    
    updated++;
    
    // Small delay to avoid rate limiting
    if (updated % 5 === 0) {
      Utilities.sleep(1000);
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Updated ${updated} tracking numbers.\n⏭️ Skipped ${skipped} delivered items.`);
}

/**
 * Update only selected rows
 */
function updateSelectedTrackingStatus() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  let updated = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const tn = sheet.getRange(row, CONFIG.TRACKING_COLUMN).getValue();
    const currentStatus = sheet.getRange(row, CONFIG.STATUS_COLUMN).getValue().toString().toLowerCase();
    
    if (!tn || tn.toString().trim() === '' || currentStatus.includes('delivered')) {
      continue;
    }
    
    const result = fetchTrackingStatus(tn.toString());
    const cell = sheet.getRange(row, CONFIG.STATUS_COLUMN);
    
    if (result.url) {
      cell.setFormula(`=HYPERLINK("${result.url}","${result.status}")`);
    } else {
      cell.setValue(result.status);
    }
    
    updated++;
  }
  
  SpreadsheetApp.getUi().alert(`✅ Updated ${updated} tracking numbers.`);
}

/**
 * Custom function for single cell lookup
 * Usage: =TRACKSTATUS(D2)
 */
function TRACKSTATUS(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') {
    return '';
  }
  
  const result = fetchTrackingStatus(trackingNumber.toString());
  return result.status;
}

/**
 * Custom function that returns tracking URL
 * Usage: =TRACKURL(D2)
 */
function TRACKURL(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') {
    return '';
  }
  
  const { url } = detectCarrier(trackingNumber.toString());
  return url || '';
}

/**
 * Custom function for carrier detection
 * Usage: =CARRIER(D2)
 */
function CARRIER(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') {
    return '';
  }
  
  const { carrier } = detectCarrier(trackingNumber.toString());
  return carrier;
}

/**
 * Show settings dialog
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #333; }
      .setting { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
      .info { color: #666; font-size: 12px; margin-top: 5px; }
      .statuses { background: #f5f5f5; padding: 15px; border-radius: 8px; margin-top: 20px; }
      code { background: #e0e0e0; padding: 2px 6px; border-radius: 3px; }
    </style>
    <h2>📦 Tracking Settings</h2>
    <p>Current configuration:</p>
    <div class="setting">
      <label>Tracking Number Column:</label>
      <input type="text" value="D (column 4)" disabled>
      <div class="info">Edit CONFIG in Apps Script to change</div>
    </div>
    <div class="setting">
      <label>Status Column:</label>
      <input type="text" value="E (column 5)" disabled>
    </div>
    <div class="statuses">
      <h3>Possible Status Values:</h3>
      <ul>
        <li><strong>Delivered</strong> - Package delivered</li>
        <li><strong>Out for Delivery</strong> - On truck for delivery today</li>
        <li><strong>In Transit</strong> - Moving through carrier network</li>
        <li><strong>Exception</strong> - Problem with delivery</li>
        <li><strong>Waiting for Pickup</strong> - At carrier location for pickup</li>
        <li><strong>Label Created</strong> - Shipping label created, not shipped yet</li>
        <li><strong>Not Found</strong> - Tracking number not recognized</li>
      </ul>
    </div>
  `)
    .setWidth(400)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}
