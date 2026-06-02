/**
 * ----------------------------------------------------------------------------
 * 📲 Google Apps Script - Mobile Portfolio Manager Web App & Public API
 * ----------------------------------------------------------------------------
 * 
 * INSTRUCTIONS:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1oDZbVB_zgXJ0OGlVDwHL8DWzPFbhfpTAtmILkf3Of1U/edit
 * 2. Go to "Extensions" in the top menu, then click "Apps Script".
 * 3. Delete any code in the editor, and paste this entire code.
 * 4. Click the Save icon (floppy disk).
 * 5. Click the "Deploy" button (top right) -> "New deployment".
 * 6. Click the gear icon next to "Select type" and choose "Web app".
 * 7. Configure:
 *    - Description: Mobile Portfolio Manager API
 *    - Execute as: "Me (your-email@gmail.com)"
 *    - Who has access: "Anyone"
 * 8. Click "Deploy". Authorize permissions when prompted.
 * 9. Copy the "Web app URL" (it will look like: https://script.google.com/macros/s/XXXX/exec).
 * 10. You can paste this URL into your website Settings or your .env file as PORTAL_MOBILE_URL=your_copied_url
 */

function doGet(e) {
  const action = e.parameter.action;
  
  // JSON API endpoints for static site integration
  if (action === 'get') {
    return ContentService.createTextOutput(JSON.stringify(getStocks()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'add') {
    const name = e.parameter.name;
    const price = e.parameter.price;
    const exchange = e.parameter.exchange || 'NSE';
    const category = e.parameter.category || 'Indians';
    const res = addStock(name, price, exchange, category);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'delete') {
    const index = parseInt(e.parameter.index);
    const res = deleteStock(index);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Default: Serve HTML for direct mobile browser interface
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('📈 Mobile Portfolio Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Support POST fallback if needed
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter;
  }
  
  const action = data.action;
  if (action === 'add') {
    const res = addStock(data.name, data.price, data.exchange, data.category);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'delete') {
    const res = deleteStock(parseInt(data.index));
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Get list of stocks from 'Stocks' worksheet
function getStocks() {
  try {
    const sheet = GetOrCreateStocksSheet();
    const rows = sheet.getDataRange().getValues();
    const stocks = [];
    
    // Skip header row
    for (let i = 1; i < rows.length; i++) {
      const name = rows[i][0];
      const price = rows[i][1];
      const exchange = rows[i][2] || 'NSE';
      const category = rows[i][3] || 'Indians';
      if (name) {
        stocks.push({
          index: i + 1, // Sheet row number (2-based, i=1 means Row 2)
          stockName: name.toString().trim(),
          buyPrice: price ? parseFloat(price) : null,
          exchange: exchange.toString().trim(),
          category: category.toString().trim()
        });
      }
    }
    return { success: true, data: stocks };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Add a stock to 'Stocks' worksheet
function addStock(name, price, exchange, category) {
  try {
    const sheet = GetOrCreateStocksSheet();
    sheet.appendRow([
      name, 
      price ? parseFloat(price) : '', 
      exchange || 'NSE', 
      category || 'Indians'
    ]);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Delete a stock by sheet row index
function deleteStock(rowIndex) {
  try {
    const sheet = GetOrCreateStocksSheet();
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function GetOrCreateStocksSheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  
  if (!ss) {
    ss = SpreadsheetApp.openById('1oDZbVB_zgXJ0OGlVDwHL8DWzPFbhfpTAtmILkf3Of1U');
  }
  
  let sheet = ss.getSheetByName('Stocks');
  if (!sheet) {
    sheet = ss.insertSheet('Stocks');
    sheet.appendRow(['Stock Name (Symbol)', 'Buy Price', 'Exchange', 'Category']);
    
    // Style Header
    const headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1F497D');
    headerRange.setFontColor('#FFFFFF');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 120);
  }
  return sheet;
}
