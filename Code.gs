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
const USPS_CONSUMER_KEY = 'XBmzFyoWGifi2bNA0sXzmP61Lyr8aa2tLzrymH7Mazrs5K3A';
const USPS_CONSUMER_SECRET = 'msBghXqWOmhAYyGiT9RfxNFNnPWUSg2Gimqu32AMxIO2nTv36wNElzGgD5mZIt35';
const FEDEX_API_KEY = 'l7dbcd6af277524dceb0a5e3765da4b211';
const FEDEX_SECRET_KEY = 'b7b6d9ac71fb4c258dc64ed3e1594522';

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
  
  // FedEx: 12, 15 digits
  if (/^\d{12}$/.test(tn) || /^\d{15}$/.test(tn)) {
    return { carrier: 'FedEx', url: `https://www.fedex.com/fedextrack/?trknbr=${tn}` };
  }
  
  // DHL: 10-11 digits or starts with JJD
  if (/^\d{10,11}$/.test(tn) || /^JJD/i.test(tn)) {
    return { carrier: 'DHL', url: `https://www.dhl.com/us-en/home/tracking/tracking-express.html?submit=1&tracking-id=${tn}` };
  }
  
  // USPS: 20-22 digits
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
  if (cachedToken) return cachedToken;
  
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
 * Get USPS OAuth Access Token (new API platform)
 */
function getUSPSAccessToken() {
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get('usps_token');
  if (cachedToken) return cachedToken;
  
  const tokenUrl = 'https://api.usps.com/oauth2/v3/token';
  
  const response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: `grant_type=client_credentials&client_id=${USPS_CONSUMER_KEY}&client_secret=${USPS_CONSUMER_SECRET}`,
    muteHttpExceptions: true
  });
  
  const data = JSON.parse(response.getContentText());
  Logger.log('USPS token response: ' + JSON.stringify(data));
  
  if (data.access_token) {
    // Cache for 50 minutes (typical token validity is 60 min)
    cache.put('usps_token', data.access_token, 3000);
    return data.access_token;
  }
  
  Logger.log('USPS token error: ' + response.getContentText());
  return null;
}

/**
 * Get FedEx OAuth Access Token
 */
function getFedExAccessToken() {
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get('fedex_token');
  if (cachedToken) return cachedToken;
  
  const tokenUrl = 'https://apis.fedex.com/oauth/token';
  
  const response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: `grant_type=client_credentials&client_id=${FEDEX_API_KEY}&client_secret=${FEDEX_SECRET_KEY}`,
    muteHttpExceptions: true
  });
  
  const data = JSON.parse(response.getContentText());
  if (data.access_token) {
    cache.put('fedex_token', data.access_token, 3000);
    return data.access_token;
  }
  
  Logger.log('FedEx token error: ' + response.getContentText());
  return null;
}

/**
 * Fetch UPS tracking status
 */
function fetchUPSStatus(trackingNumber) {
  try {
    const token = getUPSAccessToken();
    if (!token) return STATUS.EXCEPTION;
    
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
      const pkg = data.trackResponse.shipment[0].package ? data.trackResponse.shipment[0].package[0] : null;
      
      if (pkg && pkg.currentStatus) {
        const status = pkg.currentStatus.description.toLowerCase();
        const code = pkg.currentStatus.code || '';
        
        if (status.includes('delivered') || code === 'D') return STATUS.DELIVERED;
        if (status.includes('out for delivery') || code === 'O') return STATUS.OUT_FOR_DELIVERY;
        if (status.includes('exception') || code === 'X') return STATUS.EXCEPTION;
        if (status.includes('pickup') || code === 'P') return STATUS.WAITING_PICKUP;
        if (status.includes('in transit') || code === 'I') return STATUS.IN_TRANSIT;
        if (status.includes('label') || code === 'M') return STATUS.LABEL_CREATED;
        
        return STATUS.IN_TRANSIT;
      }
    }
    
    return STATUS.NOT_FOUND;
  } catch (e) {
    Logger.log('UPS error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch USPS tracking status (new API platform)
 */
function fetchUSPSStatus(trackingNumber) {
  try {
    const token = getUSPSAccessToken();
    if (!token) {
      Logger.log('USPS: No token available');
      return STATUS.NOT_FOUND;
    }
    
    const url = `https://api.usps.com/tracking/v3/tracking/${trackingNumber}?expand=DETAIL`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const responseText = response.getContentText();
    Logger.log('USPS response: ' + responseText.substring(0, 500));
    
    const data = JSON.parse(responseText);
    
    // Check for errors
    if (data.error || data.errors) {
      Logger.log('USPS API error: ' + JSON.stringify(data.error || data.errors));
      return STATUS.NOT_FOUND;
    }
    
    // New USPS API response structure
    if (data.trackingNumber) {
      const statusCategory = (data.statusCategory || '').toLowerCase();
      const statusSummary = (data.status || '').toLowerCase();
      
      // Check status category first (more reliable)
      if (statusCategory === 'delivered' || statusSummary.includes('delivered')) {
        return STATUS.DELIVERED;
      }
      if (statusCategory === 'out for delivery' || statusSummary.includes('out for delivery')) {
        return STATUS.OUT_FOR_DELIVERY;
      }
      if (statusCategory === 'alert' || statusCategory === 'exception' || 
          statusSummary.includes('alert') || statusSummary.includes('undeliverable') || 
          statusSummary.includes('return')) {
        return STATUS.EXCEPTION;
      }
      if (statusCategory === 'available for pickup' || statusSummary.includes('pickup') || 
          statusSummary.includes('held at') || statusSummary.includes('notice left')) {
        return STATUS.WAITING_PICKUP;
      }
      if (statusCategory === 'pre-shipment' || statusSummary.includes('label') || 
          statusSummary.includes('pre-shipment') || statusSummary.includes('shipping label created')) {
        return STATUS.LABEL_CREATED;
      }
      if (statusCategory === 'in transit' || statusSummary.includes('transit') || 
          statusSummary.includes('arrived') || statusSummary.includes('departed') || 
          statusSummary.includes('processed') || statusSummary.includes('accepted') ||
          statusSummary.includes('in-transit') || statusSummary.includes('origin')) {
        return STATUS.IN_TRANSIT;
      }
      
      // Default to in transit if we have valid tracking data
      if (data.status) {
        return STATUS.IN_TRANSIT;
      }
    }
    
    return STATUS.NOT_FOUND;
  } catch (e) {
    Logger.log('USPS error: ' + e);
    return STATUS.NOT_FOUND;
  }
}

/**
 * Fetch FedEx tracking status
 */
function fetchFedExStatus(trackingNumber) {
  try {
    const token = getFedExAccessToken();
    if (!token) {
      Logger.log('FedEx: No token available');
      return STATUS.EXCEPTION;
    }
    
    const url = 'https://apis.fedex.com/track/v1/trackingnumbers';
    
    const payload = {
      trackingInfo: [{
        trackingNumberInfo: {
          trackingNumber: trackingNumber
        }
      }],
      includeDetailedScans: false
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'X-locale': 'en_US'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.output && data.output.completeTrackResults && data.output.completeTrackResults[0]) {
      const result = data.output.completeTrackResults[0];
      
      if (result.trackResults && result.trackResults[0]) {
        const track = result.trackResults[0];
        
        if (track.latestStatusDetail) {
          const statusCode = track.latestStatusDetail.code || '';
          const statusDesc = (track.latestStatusDetail.description || '').toLowerCase();
          
          if (statusCode === 'DL' || statusDesc.includes('delivered')) return STATUS.DELIVERED;
          if (statusCode === 'OD' || statusDesc.includes('out for delivery')) return STATUS.OUT_FOR_DELIVERY;
          if (statusCode === 'DE' || statusCode === 'SE' || statusDesc.includes('exception')) return STATUS.EXCEPTION;
          if (statusCode === 'HL' || statusDesc.includes('hold') || statusDesc.includes('pickup')) return STATUS.WAITING_PICKUP;
          if (statusCode === 'IT' || statusDesc.includes('in transit')) return STATUS.IN_TRANSIT;
          if (statusCode === 'LB' || statusDesc.includes('label') || statusDesc.includes('shipment information')) return STATUS.LABEL_CREATED;
          
          return STATUS.IN_TRANSIT;
        }
      }
    }
    
    return STATUS.NOT_FOUND;
  } catch (e) {
    Logger.log('FedEx error: ' + e);
    return STATUS.EXCEPTION;
  }
}

/**
 * Fetch DHL tracking status
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
    Logger.log('Error: ' + e.toString());
    return { status: STATUS.EXCEPTION, url, carrier };
  }
}

/**
 * Update all tracking statuses
 */
function updateAllTrackingStatus() {
  const sheet = CONFIG.SHEET_NAME 
    ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME)
    : SpreadsheetApp.getActiveSheet();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No data found.');
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
    
    if (!tn || tn.toString().trim() === '') continue;
    if (currentStatus.includes('delivered')) { skipped++; continue; }
    
    const result = fetchTrackingStatus(tn.toString());
    const cell = sheet.getRange(CONFIG.START_ROW + i, CONFIG.STATUS_COLUMN);
    
    if (result.url) {
      cell.setFormula(`=HYPERLINK("${result.url}","${result.status}")`);
    } else {
      cell.setValue(result.status);
    }
    
    updated++;
    if (updated % 5 === 0) Utilities.sleep(500);
  }
  
  SpreadsheetApp.getUi().alert(`✅ Updated ${updated} tracking numbers.\n⏭️ Skipped ${skipped} delivered.`);
}

/**
 * Update selected rows only
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
    
    if (!tn || tn.toString().trim() === '' || currentStatus.includes('delivered')) continue;
    
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
 * Custom functions
 */
function TRACKSTATUS(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') return '';
  return fetchTrackingStatus(trackingNumber.toString()).status;
}

function TRACKURL(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') return '';
  return detectCarrier(trackingNumber.toString()).url || '';
}

function CARRIER(trackingNumber) {
  if (!trackingNumber || trackingNumber.toString().trim() === '') return '';
  return detectCarrier(trackingNumber.toString()).carrier;
}

/**
 * Settings dialog
 */
function showSettings() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #333; }
      .api-status { padding: 10px; border-radius: 8px; margin: 10px 0; }
      .api-ok { background: #d4edda; color: #155724; }
      .api-warn { background: #fff3cd; color: #856404; }
      .statuses { background: #f5f5f5; padding: 15px; border-radius: 8px; margin-top: 20px; }
    </style>
    <h2>📦 Tracking Settings</h2>
    <h3>API Status:</h3>
    <div class="api-status api-ok">✅ UPS: Official API</div>
    <div class="api-status api-ok">✅ USPS: Official API (new platform)</div>
    <div class="api-status api-ok">✅ FedEx: Official API</div>
    <div class="api-status api-warn">⚠️ DHL: Web scraping</div>
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
