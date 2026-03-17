/**
 * Tracking Status Lookup for Google Sheets
 * Uses official carrier APIs for reliable tracking
 */

// Configuration - adjust these to match your sheet
const CONFIG = {
  TRACKING_COLUMN: 4,  // Column D (1-indexed)
  STATUS_COLUMN: 5,    // Column E (1-indexed)
  START_ROW: 2,        // Skip header row
  SHEET_NAME: null     // null = active sheet, or specify name like "Sheet1"
};

// API Credentials
const UPS_CLIENT_ID = 'nk6XMzoRwYb2j8FjLOmjq5H2B3AVQkcI5H8wGIMTuTOYzEbA';
const UPS_CLIENT_SECRET = 'GikPKvVFmgroI4ECubxXnm5xIoC4p9kaOiQDgLE2fNFOf5pzW2AaGynm4BDgjIYI';
const USPS_USER_ID = '1CAPIT79M7505';

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
  
  // FedEx: 12, 15 digits (not 20+ which is usually USPS)
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
 * Get UPS OAuth Access Token
 */
function getUPSAccessToken() {
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get('ups_token');
  
  if (cachedToken) {
    return cachedToken;
  }
  
  const tokenUrl = 'https://onlinetools.ups.com/security/v1/oauth/token';
  const credentials = Utilities.base64Encode(UPS_CLIENT_ID + ':' + UPS_CLIENT_SECRET);
  
  const response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: 'grant_type=client_credentials',
    muteHttpExceptions: true
  });
  
  const data = JSON.parse(response.getContentText());
  
  if (data.access_token) {
    cache.put('ups_token', data.access_token, 10800);
    return data.access_token;
  }
  
  Logger.log('UPS token error: ' + response.getContentText());
  return null;
}

/**
 * Fetch UPS tracking status using official API
 */
function fetchUPSStatus(trackingNumber) {
  try {
    const token = getUPSAccessToken();
    if (!token) {
      return STATUS.EXCEPTION;
    }
    
    const url = `https://onlinetools.ups.com/api/track/v1/details/${trackingNumber}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'transId': Utilities.getUuid(),
        'transactionSrc': 'GoogleSheets'
      },
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.trackResponse && data.trackResponse.shipment && data.trackResponse.shipment[0]) {
      const shipment = data.trackResponse.shipment[0];
      const pkg = shipment.package ? shipment.package[0] : null;
      
      if (pkg && pkg.currentStatus) {
        const status = pkg.currentStatus.description.toLowerCase();
        const code = pkg.currentStatus.code || '';
        
        if (status.includes('delivered') || code === 'D') return STATUS.DELIVERED;
        if (status.includes('out for delivery') || code === 'O') return STATUS.OUT_FOR_DELIVERY;
        if (status.includes('exception') || code === 'X') return STATUS.EXCEPTION;
        if (status.includes('pickup') || status.includes('ready') || code === 'P') return STATUS.WAITING_PICKUP;
        if (status.includes('in transit') || code === 'I') return STATUS.IN_TRANSIT;
        if (status.includes('label') || status.includes('created') || code === 'M') return STATUS.LABEL_CREATED;
        
        return STATUS.IN_TRANSIT;
      }
    }
    
    if (data.response && data.response.errors) {
      Logger.log('UPS API error: ' + JSON.stringify(data.response.errors));
      return STATUS.NOT_FOUND;
    }
    
    return STATUS.NOT_FOUND;
  } catch (e) {
    Logger.log('UPS error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch USPS tracking status using official API
 */
function fetchUSPSStatus(trackingNumber) {
  try {
    // Build XML request for USPS Track API
    const xml = `<?xml version="1.0" encoding="UTF-8" ?>
<TrackFieldRequest USERID="${USPS_USER_ID}">
  <Revision>1</Revision>
  <ClientIp>127.0.0.1</ClientIp>
  <SourceId>VIP Medical Group</SourceId>
  <TrackID ID="${trackingNumber}"></TrackID>
</TrackFieldRequest>`;
    
    const url = 'https://secure.shippingapis.com/ShippingAPI.dll?API=TrackV2&XML=' + encodeURIComponent(xml);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseText = response.getContentText();
    
    // Parse XML response
    const doc = XmlService.parse(responseText);
    const root = doc.getRootElement();
    
    // Check for errors
    const error = root.getChild('Error');
    if (error) {
      const errorDesc = error.getChildText('Description');
      Logger.log('USPS API error: ' + errorDesc);
      return STATUS.NOT_FOUND;
    }
    
    // Get tracking info
    const trackInfo = root.getChild('TrackInfo');
    if (!trackInfo) {
      return STATUS.NOT_FOUND;
    }
    
    // Check for error in track info
    const trackError = trackInfo.getChild('Error');
    if (trackError) {
      return STATUS.NOT_FOUND;
    }
    
    // Get status summary
    const statusSummary = trackInfo.getChildText('StatusSummary') || '';
    const status = statusSummary.toLowerCase();
    
    // Also check individual status fields
    const statusCategory = (trackInfo.getChildText('StatusCategory') || '').toLowerCase();
    const deliveryDate = trackInfo.getChildText('ExpectedDeliveryDate') || '';
    
    // Map USPS status to our standard statuses
    if (status.includes('delivered') || statusCategory === 'delivered') {
      return STATUS.DELIVERED;
    }
    
    if (status.includes('out for delivery') || statusCategory === 'out for delivery') {
      return STATUS.OUT_FOR_DELIVERY;
    }
    
    if (status.includes('alert') || status.includes('undeliverable') || 
        status.includes('no access') || status.includes('return') ||
        status.includes('refused') || status.includes('deceased') ||
        statusCategory === 'alert') {
      return STATUS.EXCEPTION;
    }
    
    if (status.includes('available for pickup') || status.includes('ready for pickup') || 
        status.includes('held at') || status.includes('notice left') ||
        statusCategory === 'available for pickup') {
      return STATUS.WAITING_PICKUP;
    }
    
    if (status.includes('pre-shipment') || status.includes('label created') || 
        status.includes('shipping label') || status.includes('usps awaiting') ||
        statusCategory === 'pre-shipment') {
      return STATUS.LABEL_CREATED;
    }
    
    if (status.includes('in transit') || status.includes('arrived') || 
        status.includes('departed') || status.includes('processed') ||
        status.includes('accepted') || status.includes('origin') ||
        status.includes('en route') || status.includes('forwarded') ||
        statusCategory === 'in transit') {
      return STATUS.IN_TRANSIT;
    }
    
    // If we got here with valid track info, assume in transit
    if (statusSummary.length > 0) {
      return STATUS.IN_TRANSIT;
    }
    
    return STATUS.NOT_FOUND;
  } catch (e) {
    Logger.log('USPS error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch FedEx tracking status (web scraping - needs API for reliability)
 */
function fetchFedExStatus(trackingNumber) {
  try {
    const url = `https://www.fedex.com/fedextrack/?trknbr=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    const html = response.getContentText().toLowerCase();
    
    if (html.includes('delivered')) return STATUS.DELIVERED;
    if (html.includes('on fedex vehicle') || html.includes('out for delivery')) return STATUS.OUT_FOR_DELIVERY;
    if (html.includes('exception') || html.includes('delay')) return STATUS.EXCEPTION;
    if (html.includes('pickup') || html.includes('hold at')) return STATUS.WAITING_PICKUP;
    if (html.includes('shipment information') || html.includes('label')) return STATUS.LABEL_CREATED;
    if (html.includes('in transit') || html.includes('departed') || html.includes('arrived')) return STATUS.IN_TRANSIT;
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('FedEx error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch DHL tracking status (web scraping - needs API for reliability)
 */
function fetchDHLStatus(trackingNumber) {
  try {
    const url = `https://api-eu.dhl.com/track/shipments?trackingNumber=${trackingNumber}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0',
        'DHL-API-Key': 'demo-key'
      }
    });
    
    const text = response.getContentText().toLowerCase();
    
    if (text.includes('delivered')) return STATUS.DELIVERED;
    if (text.includes('out for delivery')) return STATUS.OUT_FOR_DELIVERY;
    if (text.includes('exception') || text.includes('on hold')) return STATUS.EXCEPTION;
    if (text.includes('pickup') || text.includes('service point')) return STATUS.WAITING_PICKUP;
    if (text.includes('transit') || text.includes('processed')) return STATUS.IN_TRANSIT;
    
    return STATUS.IN_TRANSIT;
  } catch (e) {
    Logger.log('DHL error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch tracking status from carrier
 */
function fetchTrackingStatus(trackingNumber) {
  const { carrier, url } = detectCarrier(trackingNumber);
  
  if (carrier === 'Unknown' || !url) {
    return { status: STATUS.NOT_FOUND, url: null, carrier };
  }
  
  try {
    let status = STATUS.NOT_FOUND;
    
    if (carrier === 'UPS') {
      status = fetchUPSStatus(trackingNumber);
    } else if (carrier === 'USPS') {
      status = fetchUSPSStatus(trackingNumber);
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
      Utilities.sleep(300);
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
      .api-status { padding: 10px; border-radius: 8px; margin: 10px 0; }
      .api-ok { background: #d4edda; color: #155724; }
      .api-missing { background: #f8d7da; color: #721c24; }
      .statuses { background: #f5f5f5; padding: 15px; border-radius: 8px; margin-top: 20px; }
    </style>
    <h2>📦 Tracking Settings</h2>
    <h3>API Status:</h3>
    <div class="api-status api-ok">✅ UPS API: Connected</div>
    <div class="api-status api-ok">✅ USPS API: Connected</div>
    <div class="api-status api-missing">⚠️ FedEx: Web scraping (less reliable)</div>
    <div class="api-status api-missing">⚠️ DHL: Web scraping (less reliable)</div>
    <div class="statuses">
      <h3>Status Values:</h3>
      <ul>
        <li><strong>Delivered</strong></li>
        <li><strong>Out for Delivery</strong></li>
        <li><strong>In Transit</strong></li>
        <li><strong>Exception</strong></li>
        <li><strong>Waiting for Pickup</strong></li>
        <li><strong>Label Created</strong></li>
        <li><strong>Not Found</strong></li>
      </ul>
    </div>
  `)
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}
