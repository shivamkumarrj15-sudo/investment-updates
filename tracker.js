// ============================================================================
// 📈 Investment Tracker Bot v2.0 - Daily Analysis and FII/DII Parser
// ============================================================================

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const FormData = require('form-data');

// Configurations
const CONFIG = {
  excelFile: path.join(__dirname, 'stocks.xlsx'),
  googleSheetTemplateUrl: 'https://docs.google.com/spreadsheets/d/1oDZbVB_zgXJ0OGlVDwHL8DWzPFbhfpTAtmILkf3Of1U/export?format=xlsx',
  mobilePortalUrl: process.env.PORTAL_MOBILE_URL,
  portalWebsiteUrl: process.env.PORTAL_WEBSITE_URL,
  newsApi: {
    key: process.env.NEWS_API_KEY,
    baseUrl: 'https://newsapi.org/v2/everything',
  },
  openRouter: {
    key: process.env.OPENROUTER_API_KEY,
    baseUrl: 'https://openrouter.ai/api/v1/chat/completions',
  },
  email: {
    sender: process.env.SENDER_EMAIL,
    password: process.env.SENDER_APP_PASSWORD,
    receiver: process.env.RECEIVER_EMAIL,
  }
};

// Check for required configuration keys
if (!CONFIG.newsApi.key || !CONFIG.openRouter.key || !CONFIG.email.sender || !CONFIG.email.password || !CONFIG.email.receiver) {
  console.error('❌ Error: Missing environment variables in .env. Please check your credentials.');
  process.exit(1);
}

// NSE Official Holidays 2026 (for automated runs skipping holidays)
const NSE_HOLIDAYS_2026 = [
  '2026-01-26', // Republic Day
  '2026-03-03', // Holi
  '2026-03-26', // Ram Navami
  '2026-03-31', // Mahavir Jayanti
  '2026-04-03', // Good Friday
  '2026-04-14', // Ambedkar Jayanti
  '2026-05-01', // Maharashtra Day
  '2026-05-28', // Bakri Id
  '2026-06-26', // Muharram
  '2026-09-14', // Ganesh Chaturthi
  '2026-10-02', // Gandhi Jayanti
  '2026-10-20', // Dussehra
  '2026-11-10', // Diwali Balipratipada
  '2026-11-23', // Gurunanak Jayanti
  '2026-12-25'  // Christmas
];

// Check if today is weekend or holiday in Indian Standard Time (IST)
const todayCheck = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
const dayOfWeek = todayCheck.getDay(); // 0 = Sunday, 6 = Saturday
const yyyy = todayCheck.getFullYear();
const mm = String(todayCheck.getMonth() + 1).padStart(2, '0');
const dd = String(todayCheck.getDate()).padStart(2, '0');
const todayStr = `${yyyy}-${mm}-${dd}`;

if (dayOfWeek === 0 || dayOfWeek === 6) {
  console.log(`📅 Today is Weekend in IST (Day ${dayOfWeek}). Skipping automated run.`);
  process.exit(0);
}

if (NSE_HOLIDAYS_2026.includes(todayStr)) {
  console.log(`📅 Today (${todayStr}) is an official NSE holiday in India. Skipping automated run.`);
  process.exit(0);
}

/**
 * Standard HTTP headers for fetching data from NSE (prevents 403 blocks)
 */
const NSE_HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.9',
  'Referer': 'https://www.nseindia.com/'
};

/**
 * Helper to parse Date into DDMMYYYY string
 */
function formatDateToDDMMYYYY(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}${month}${year}`;
}

/**
 * Helper to format date for human display (e.g. May 27)
 */
function formatDateToHuman(date) {
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

/**
 * Clean up the stock name to make a strong search query for News API.
 */
function getSearchQuery(stockName) {
  const match = stockName.match(/^(.*?)\s*\(([^)]+)\)$/);
  if (match) {
    const name = match[1].trim();
    const symbol = match[2].trim();
    const cleanName = name.replace(/(Inc\.|Ltd\.|Corporation|Limited|Co\.)/gi, '').trim();
    return `("${cleanName}" OR "${symbol}") AND (stock OR earnings OR market OR finance)`;
  }
  const cleanName = stockName.replace(/(Inc\.|Ltd\.|Corporation|Limited|Co\.)/gi, '').trim();
  return `"${cleanName}" AND (stock OR finance OR market)`;
}

/**
 * Fetches FII/DII CSV files for the last 3 trading days from NSE
 */
async function fetchLast3TradingDays() {
  console.log('📅 Finding the last 3 active trading days from NSE archives...');
  const tradingDays = [];
  let currentDate = new Date();
  let attempts = 0;
  
  // Look back up to 15 days to collect 3 successful trading days
  while (tradingDays.length < 3 && attempts < 15) {
    attempts++;
    const dayOfWeek = currentDate.getDay();
    
    // Skip Saturdays (6) and Sundays (0)
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      currentDate.setDate(currentDate.getDate() - 1);
      continue;
    }
    
    const ddmmyyyy = formatDateToDDMMYYYY(currentDate);
    const url = `https://archives.nseindia.com/content/nsccl/fao_participant_oi_${ddmmyyyy}.csv`;
    
    try {
      console.log(`   Trying date: ${formatDateToHuman(currentDate)} (${ddmmyyyy})...`);
      const response = await axios.get(url, { headers: NSE_HEADERS, timeout: 8000 });
      
      if (response.status === 200 && response.data && response.data.includes('Client Type')) {
        console.log(`   ✅ Successful! Trading Day ${tradingDays.length + 1} Found.`);
        tradingDays.push({
          date: new Date(currentDate),
          dateStr: ddmmyyyy,
          humanDate: formatDateToHuman(currentDate),
          csvText: response.data
        });
      }
    } catch (err) {
      // 404 means the market was closed (holiday) or data is not uploaded yet for today
    }
    
    currentDate.setDate(currentDate.getDate() - 1);
  }
  
  if (tradingDays.length < 3) {
    throw new Error(`Could only find ${tradingDays.length} trading days in the last 15 days. Check your network or NSE archives.`);
  }
  
  return tradingDays;
}

/**
 * Parses the NSE participant wise OI CSV
 */
function parseNseCsv(csvText) {
  const lines = csvText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
  const data = {};
  
  let headerIndex = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].toLowerCase().includes('client type') || lines[i].toLowerCase().includes('future index')) {
      headerIndex = i;
      break;
    }
  }
  
  if (headerIndex === -1) {
    throw new Error('Invalid CSV structure. Header row not found.');
  }
  
  const headers = lines[headerIndex].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  
  for (let i = headerIndex + 1; i < lines.length; i++) {
    const rowCells = lines[i].split(',').map(c => c.trim().replace(/^"|"$/g, ''));
    const clientType = rowCells[0];
    if (['client', 'dii', 'fii', 'pro', 'total'].includes(clientType.toLowerCase())) {
      const rowData = {};
      headers.forEach((header, idx) => {
        if (idx > 0) {
          rowData[header] = parseFloat(rowCells[idx]) || 0;
        }
      });
      data[clientType.toLowerCase()] = rowData;
    }
  }
  return data;
}

/**
 * Downloads the Google Sheets template workbook
 */
async function downloadTemplate() {
  const tempPath = path.join(__dirname, 'temp_template.xlsx');
  console.log(`📥 Downloading Google Sheet template from: ${CONFIG.googleSheetTemplateUrl}`);
  
  try {
    const response = await axios({
      method: 'get',
      url: CONFIG.googleSheetTemplateUrl,
      responseType: 'stream'
    });
    
    const writer = fs.createWriteStream(tempPath);
    response.data.pipe(writer);
    
    await new Promise((resolve, reject) => {
      writer.on('finish', resolve);
      writer.on('error', reject);
    });
    
    console.log('✅ Template downloaded successfully!');
    return tempPath;
  } catch (error) {
    console.error('⚠️ Template download failed. Falling back to local stocks.xlsx if present.');
    return null;
  }
}

/**
 * Reads the portfolio list (Stocks and Buy Prices) from a workbook object
 */
function getPortfolioFromWorkbook(workbook) {
  const portfolio = [];
  try {
    const sheet = workbook.getWorksheet('Stocks');
    if (sheet) {
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          const name = row.getCell(1).value;
          const price = row.getCell(2).value;
          const exchange = row.getCell(3).value || 'NSE';
          const category = row.getCell(4).value || 'Indians';
          if (name) {
            portfolio.push({
              stockName: name.toString().trim(),
              buyPrice: price ? parseFloat(price) : null,
              exchange: exchange.toString().trim(),
              category: category.toString().trim()
            });
          }
        }
      });
    }
  } catch (e) {
    console.warn('⚠️ Could not parse Stocks sheet from workbook:', e.message);
  }
  return portfolio;
}

/**
 * Preserves the user's custom portfolio list from local stocks.xlsx if present
 */
async function getSavedPortfolio() {
  if (fs.existsSync(CONFIG.excelFile)) {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(CONFIG.excelFile);
      return getPortfolioFromWorkbook(workbook);
    } catch (e) {
      console.warn('⚠️ Could not read local portfolio file. Using default template stocks.');
    }
  }
  return [];
}

/**
 * Writes the parsed CSV tables into the 'data' sheet of the workbook
 */
function writeDataToDataSheet(sheet, tradingDays) {
  // Columns Mapping for writing
  const headers = [
    'Client Type', 'Future Index Long', 'Future Index Short', 'Future Stock Long', 'Future Stock Short',
    'Option Index Call Long', 'Option Index Put Long', 'Option Index Call Short', 'Option Index Put Short',
    'Option Stock Call Long', 'Option Stock Put Long', 'Option Stock Call Short', 'Option Stock Put Short',
    'Total Long Contracts', 'Total Short Contracts'
  ];
  
  // Set Column widths
  sheet.columns = headers.map(h => ({ header: h, width: 22 }));

  // Helper to write a table at a specific start row
  const writeTable = (startRow, titleDate, parsedData) => {
    // 1. Title Row
    sheet.getCell(startRow, 1).value = `Participant wise Open Interest (no. of contracts) in Equity Derivatives as on ${titleDate}`;
    sheet.getCell(startRow, 1).font = { bold: true, name: 'Segoe UI', size: 11 };
    
    // 2. Header Row
    const headerRow = sheet.getRow(startRow + 1);
    headers.forEach((h, idx) => {
      headerRow.getCell(idx + 1).value = h;
      headerRow.getCell(idx + 1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.getCell(idx + 1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E293B' }
      };
    });
    
    // 3. Data Rows (Client, DII, FII, Pro, TOTAL)
    const participants = ['client', 'dii', 'fii', 'pro', 'total'];
    participants.forEach((p, pIdx) => {
      const targetRow = sheet.getRow(startRow + 2 + pIdx);
      targetRow.getCell(1).value = p.charAt(0).toUpperCase() + p.slice(1);
      
      const pData = parsedData[p] || {};
      headers.forEach((h, hIdx) => {
        if (hIdx > 0) {
          targetRow.getCell(hIdx + 1).value = pData[h] || 0;
        }
      });
    });
  };

  // Day 0: Today (Bottom table, starts at Row 26 in template)
  const d0Parsed = parseNseCsv(tradingDays[0].csvText);
  writeTable(26, tradingDays[0].humanDate, d0Parsed);

  // Day 1: Yesterday (Top table, starts at Row 3 in template)
  const d1Parsed = parseNseCsv(tradingDays[1].csvText);
  writeTable(3, tradingDays[1].humanDate, d1Parsed);

  // Day 2: Two days ago (Middle table, starts at Row 15 in template)
  const d2Parsed = parseNseCsv(tradingDays[2].csvText);
  // Note: Middle table in template starts at row 15.
  writeTable(15, tradingDays[2].humanDate, d2Parsed);

  sheet.views = [{ showGridLines: true }];
  console.log('✅ Wrote 3 days of raw NSE OI data to the "data" worksheet.');
}

/**
 * Re-writes/corrects the formulas and labels on the 'overall data' sheet
 */
function writeCalculationsToOverallSheet(sheet) {
  sheet.views = [{ showGridLines: true }];

  // Column structure for changes
  sheet.getCell('B3').value = 'Future call';
  sheet.getCell('C3').value = 'future put';
  sheet.getCell('D3').value = 'option call';
  sheet.getCell('E3').value = 'option put ';

  sheet.getCell('B18').value = 'Future call';
  sheet.getCell('C18').value = 'future put';
  sheet.getCell('D18').value = 'option call';
  sheet.getCell('E18').value = 'option put ';

  // Labeled rows
  // Table 1 (Today vs Yesterday) - Left Table (Rows 4-7)
  const rowLabelsLeft = {
    4: 'Client',
    5: 'FII',
    6: 'DII',
    7: 'PROS'
  };
  Object.keys(rowLabelsLeft).forEach(rNum => {
    sheet.getCell(`A${rNum}`).value = rowLabelsLeft[rNum];
  });

  // Table 2 (Yesterday vs 2 days ago) - Left Table (Rows 19-22)
  const rowLabelsLeftYesterday = {
    19: 'Client',
    20: 'FII',
    21: 'DII',
    22: 'PROS'
  };
  Object.keys(rowLabelsLeftYesterday).forEach(rNum => {
    sheet.getCell(`A${rNum}`).value = rowLabelsLeftYesterday[rNum];
  });

  // Table 3 (Right Summary Table) - Rows 10-12
  sheet.getCell('I9').value = 'over al data';
  sheet.getCell('J9').value = 'future';
  sheet.getCell('K9').value = 'option';
  sheet.getCell('L9').value = 'Total';
  
  sheet.getCell('I10').value = 'Client';
  sheet.getCell('I11').value = 'FII';
  sheet.getCell('I12').value = 'PROS';
  sheet.getCell('I13').value = 'TOTAL';

  // --- Write Formulas ---

  // Yesterday vs 2 days ago changes (Row 19-22)
  // Client (Row 19)
  sheet.getCell('B19').value = { formula: 'data!F5-data!F17' }; // Option Index Call Long
  sheet.getCell('C19').value = { formula: 'data!G5-data!G17' }; // Option Index Put Long
  sheet.getCell('D19').value = { formula: 'data!J5-data!J17' }; // Option Stock Call Long
  sheet.getCell('E19').value = { formula: 'data!K5-data!K17' }; // Option Stock Put Long

  // FII (Row 20)
  sheet.getCell('B20').value = { formula: 'data!F7-data!F19' };
  sheet.getCell('C20').value = { formula: 'data!G7-data!G19' };
  sheet.getCell('D20').value = { formula: 'data!J7-data!J19' };
  sheet.getCell('E20').value = { formula: 'data!K7-data!K19' };

  // DII (Row 21)
  sheet.getCell('B21').value = { formula: 'data!F6-data!F18' };
  sheet.getCell('C21').value = { formula: 'data!G6-data!G18' };
  sheet.getCell('D21').value = { formula: 'data!J6-data!J18' };
  sheet.getCell('E21').value = { formula: 'data!K6-data!K18' };

  // PROS (Row 22)
  sheet.getCell('B22').value = { formula: 'data!F8-data!F20' };
  sheet.getCell('C22').value = { formula: 'data!G8-data!G20' };
  sheet.getCell('D22').value = { formula: 'data!J8-data!J20' };
  sheet.getCell('E22').value = { formula: 'data!K8-data!K20' };


  // Today vs Yesterday changes (Row 4-7)
  // Client (Row 4) - Today Client (Row 28) vs Yesterday Client (Row 5)
  sheet.getCell('B4').value = { formula: 'data!F28-data!F5' };
  sheet.getCell('C4').value = { formula: 'data!G28-data!G5' };
  sheet.getCell('D4').value = { formula: 'data!J28-data!J5' };
  sheet.getCell('E4').value = { formula: 'data!K28-data!K5' };

  // FII (Row 5) - Today FII (Row 30) vs Yesterday FII (Row 7)
  sheet.getCell('B5').value = { formula: 'data!F30-data!F7' };
  sheet.getCell('C5').value = { formula: 'data!G30-data!G7' };
  sheet.getCell('D5').value = { formula: 'data!J30-data!J7' };
  sheet.getCell('E5').value = { formula: 'data!K30-data!K7' };

  // DII (Row 6) - Today DII (Row 29) vs Yesterday DII (Row 6)
  sheet.getCell('B6').value = { formula: 'data!F29-data!F6' };
  sheet.getCell('C6').value = { formula: 'data!G29-data!G6' };
  sheet.getCell('D6').value = { formula: 'data!J29-data!J6' };
  sheet.getCell('E6').value = { formula: 'data!K29-data!K6' };

  // PROS (Row 7) - Today Pro (Row 31) vs Yesterday Pro (Row 8)
  sheet.getCell('B7').value = { formula: 'data!F31-data!F8' };
  sheet.getCell('C7').value = { formula: 'data!G31-data!G8' };
  sheet.getCell('D7').value = { formula: 'data!J31-data!J8' };
  sheet.getCell('E7').value = { formula: 'data!K31-data!K8' };


  // Right side Summary Table (Rows 10-12) - Future/Option Net calculations
  // Client (Row 10)
  sheet.getCell('J10').value = { formula: 'B4-C4' };
  sheet.getCell('K10').value = { formula: 'D4-E4' };
  sheet.getCell('L10').value = { formula: 'J10+K10' };

  // FII (Row 11)
  sheet.getCell('J11').value = { formula: 'B5-C5' };
  sheet.getCell('K11').value = { formula: 'D5-E5' };
  sheet.getCell('L11').value = { formula: 'J11+K11' };

  // PROS (Row 12) - FIXED to point to row 7 instead of row 6!
  sheet.getCell('J12').value = { formula: 'B7-C7' };
  sheet.getCell('K12').value = { formula: 'D7-E7' };
  sheet.getCell('L12').value = { formula: 'J12+K12' };

  // TOTALS (Row 13)
  sheet.getCell('J13').value = { formula: 'J10+J11+J12' };
  sheet.getCell('K13').value = { formula: 'K10+K11+K12' };
  sheet.getCell('L13').value = { formula: 'L10+L11+L12' };

  console.log('✅ Wrote and corrected all formulas in "overall data" worksheet.');
}

/**
 * Calculates FII & Pro positions to predict overall market direction
 */
async function analyzeFiiProMarketTrend(tradingDays) {
  console.log('🤖 Starting dynamic 3-day participant Open Interest analysis with AI...');
  const d0 = parseNseCsv(tradingDays[0].csvText);
  const d1 = parseNseCsv(tradingDays[1].csvText);
  const d2 = parseNseCsv(tradingDays[2].csvText);

  const getStats = (dayData, prevDayData, p) => {
    const fl0 = dayData[p]['Future Index Long'] || 0;
    const fs0 = dayData[p]['Future Index Short'] || 0;
    const cl0 = dayData[p]['Option Index Call Long'] || 0;
    const cs0 = dayData[p]['Option Index Call Short'] || 0;
    const pl0 = dayData[p]['Option Index Put Long'] || 0;
    const ps0 = dayData[p]['Option Index Put Short'] || 0;

    const fl1 = prevDayData[p]['Future Index Long'] || 0;
    const fs1 = prevDayData[p]['Future Index Short'] || 0;
    const cl1 = prevDayData[p]['Option Index Call Long'] || 0;
    const cs1 = prevDayData[p]['Option Index Call Short'] || 0;
    const pl1 = prevDayData[p]['Option Index Put Long'] || 0;
    const ps1 = prevDayData[p]['Option Index Put Short'] || 0;

    const netFuture0 = fl0 - fs0;
    const netFuture1 = fl1 - fs1;
    const netFutureChange = netFuture0 - netFuture1;

    const netOption0 = (cl0 - cs0) - (pl0 - ps0);
    const netOption1 = (cl1 - cs1) - (pl1 - ps1);
    const netOptionChange = netOption0 - netOption1;

    return {
      futureLong: fl0,
      futureShort: fs0,
      netFuture: netFuture0,
      futureChange: netFutureChange,
      
      callLong: cl0,
      callShort: cs0,
      putLong: pl0,
      putShort: ps0,
      netOption: netOption0,
      optionChange: netOptionChange
    };
  };

  const fii = getStats(d0, d1, 'fii');
  const pro = getStats(d0, d1, 'pro');
  const client = getStats(d0, d1, 'client');
  const dii = getStats(d0, d1, 'dii');

  const fiiD2 = getStats(d1, d2, 'fii');
  const proD2 = getStats(d1, d2, 'pro');
  const clientD2 = getStats(d1, d2, 'client');
  const diiD2 = getStats(d1, d2, 'dii');

  // Build data description for 3 days to send to AI
  const promptData = `
Analyze participant-wise Open Interest (OI) over the last 3 trading days:

Trading Day 0 (Latest - ${tradingDays[0].humanDate}):
- FII: Net Futures: ${fii.netFuture.toLocaleString()} (Change from Day 1: ${fii.futureChange >= 0 ? '+' : ''}${fii.futureChange.toLocaleString()}), Net Options: ${fii.netOption.toLocaleString()} (Change from Day 1: ${fii.optionChange >= 0 ? '+' : ''}${fii.optionChange.toLocaleString()})
- PRO: Net Futures: ${pro.netFuture.toLocaleString()} (Change from Day 1: ${pro.futureChange >= 0 ? '+' : ''}${pro.futureChange.toLocaleString()}), Net Options: ${pro.netOption.toLocaleString()} (Change from Day 1: ${pro.optionChange >= 0 ? '+' : ''}${pro.optionChange.toLocaleString()})
- CLIENT: Net Futures: ${client.netFuture.toLocaleString()} (Change from Day 1: ${client.futureChange >= 0 ? '+' : ''}${client.futureChange.toLocaleString()}), Net Options: ${client.netOption.toLocaleString()} (Change from Day 1: ${client.optionChange >= 0 ? '+' : ''}${client.optionChange.toLocaleString()})
- DII: Net Futures: ${dii.netFuture.toLocaleString()} (Change from Day 1: ${dii.futureChange >= 0 ? '+' : ''}${dii.futureChange.toLocaleString()}), Net Options: ${dii.netOption.toLocaleString()} (Change from Day 1: ${dii.optionChange >= 0 ? '+' : ''}${dii.optionChange.toLocaleString()})

Trading Day 1 (${tradingDays[1].humanDate}):
- FII: Net Futures: ${fiiD2.netFuture.toLocaleString()} (Change from Day 2: ${fiiD2.futureChange >= 0 ? '+' : ''}${fiiD2.futureChange.toLocaleString()}), Net Options: ${fiiD2.netOption.toLocaleString()} (Change from Day 2: ${fiiD2.optionChange >= 0 ? '+' : ''}${fiiD2.optionChange.toLocaleString()})
- PRO: Net Futures: ${proD2.netFuture.toLocaleString()} (Change from Day 2: ${proD2.futureChange >= 0 ? '+' : ''}${proD2.futureChange.toLocaleString()}), Net Options: ${proD2.netOption.toLocaleString()} (Change from Day 2: ${proD2.optionChange >= 0 ? '+' : ''}${proD2.optionChange.toLocaleString()})
- CLIENT: Net Futures: ${clientD2.netFuture.toLocaleString()} (Change from Day 2: ${clientD2.futureChange >= 0 ? '+' : ''}${clientD2.futureChange.toLocaleString()}), Net Options: ${clientD2.netOption.toLocaleString()} (Change from Day 2: ${clientD2.optionChange >= 0 ? '+' : ''}${clientD2.optionChange.toLocaleString()})
- DII: Net Futures: ${diiD2.netFuture.toLocaleString()} (Change from Day 2: ${diiD2.futureChange >= 0 ? '+' : ''}${diiD2.futureChange.toLocaleString()}), Net Options: ${diiD2.netOption.toLocaleString()} (Change from Day 2: ${diiD2.optionChange >= 0 ? '+' : ''}${diiD2.optionChange.toLocaleString()})

Trading Day 2 (${tradingDays[2].humanDate}):
- FII: Net Futures: ${(fiiD2.netFuture - fiiD2.futureChange).toLocaleString()}, Net Options: ${(fiiD2.netOption - fiiD2.optionChange).toLocaleString()}
- PRO: Net Futures: ${(proD2.netFuture - proD2.futureChange).toLocaleString()}, Net Options: ${(proD2.netOption - proD2.optionChange).toLocaleString()}
- CLIENT: Net Futures: ${(clientD2.netFuture - clientD2.futureChange).toLocaleString()}, Net Options: ${(clientD2.netOption - clientD2.optionChange).toLocaleString()}
- DII: Net Futures: ${(diiD2.netFuture - diiD2.futureChange).toLocaleString()}, Net Options: ${(diiD2.netOption - diiD2.optionChange).toLocaleString()}
`;

  const systemPrompt = `You are an expert financial analyst focusing on Indian stock market derivatives.
Analyze the participant-wise Open Interest (OI) futures and options trends of the last 3 trading days (FII, PRO, CLIENT, DII).
Generate a daily market prediction (BULLISH/BEARISH/RANGEBOUND) and output it in a JSON object.

Follow these strict output instructions:
1. Rationale must be in HINGLISH (Hindi written in English alphabet, e.g. "FIIs ne pichle 3 dino me continuous heavy futures selling ki hai, jabki Retail client heavy longs me hain..."). Mention the specific behavior of FII, DII, Pro, and Client over the last 3 days to support your opinion.
2. Return ONLY a valid JSON object. Do not include markdown code block syntax like \`\`\`json.

JSON Schema:
{
  "direction": "UP (BULLISH)" or "DOWN (BEARISH)" or "NEUTRAL / RANGEBOUND",
  "score": "Strong Bullish 🚀" or "Bullish 🐂" or "Neutral 😐" or "Bearish 🐻" or "Strong Bearish 🚨",
  "rationale": "3-4 sentence detailed market overview in Hinglish describing the 3-day derivative trend.",
  "fii_stance": "BULLISH (LONG)" or "BEARISH (SHORT)" or "NEUTRAL / SIDEWAYS",
  "pro_stance": "BULLISH (LONG)" or "BEARISH (SHORT)" or "NEUTRAL / SIDEWAYS",
  "client_stance": "BULLISH (LONG)" or "BEARISH (SHORT)" or "NEUTRAL / SIDEWAYS"
}`;

  const models = [
    'google/gemma-4-31b-it:free',
    'google/gemma-4-26b-a4b-it:free',
    'openrouter/free',
    'meta-llama/llama-3.3-70b-instruct:free',
    'meta-llama/llama-3.2-3b-instruct:free',
    'qwen/qwen3-coder:free',
    'z-ai/glm-4.5-air:free',
    'nousresearch/hermes-3-llama-3.1-405b:free',
    'liquid/lfm-2.5-1.2b-instruct:free',
    'openai/gpt-oss-20b:free'
  ];

  let direction = 'NEUTRAL / RANGEBOUND';
  let score = 'Neutral 😐';
  let rationale = 'Dynamic trend analysis failed to load.';
  let fiiStance = 'NEUTRAL / SIDEWAYS';
  let proStance = 'NEUTRAL / SIDEWAYS';
  let clientStance = 'NEUTRAL / SIDEWAYS';

  for (let i = 0; i < models.length; i++) {
    const modelName = models[i];
    console.log(`   🤖 Querying OpenRouter model: ${modelName} for market trend...`);
    try {
      const response = await axios.post(CONFIG.openRouter.baseUrl, {
        model: modelName,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: promptData }
        ]
      }, {
        headers: {
          'Authorization': `Bearer ${CONFIG.openRouter.key}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': 'https://github.com/google/gemini-investment-tracker',
          'X-Title': 'Investment Tracker Bot'
        },
        timeout: 25000
      });

      if (response.data?.error) {
        throw new Error(`OpenRouter Error: ${response.data.error.message || JSON.stringify(response.data.error)}`);
      }

      let content = response.data?.choices?.[0]?.message?.content;
      if (!content) {
        throw new Error('Empty response from model.');
      }

      let jsonMatch = content.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('No JSON object found in response.');
      }

      const parsed = JSON.parse(jsonMatch[0].trim());
      direction = parsed.direction || 'NEUTRAL / RANGEBOUND';
      score = parsed.score || 'Neutral 😐';
      rationale = parsed.rationale || 'Analysis overview not found.';
      fiiStance = parsed.fii_stance || 'NEUTRAL / SIDEWAYS';
      proStance = parsed.pro_stance || 'NEUTRAL / SIDEWAYS';
      clientStance = parsed.client_stance || 'NEUTRAL / SIDEWAYS';

      console.log(`   ✅ Market analysis successfully processed using model: ${modelName}`);
      break;
    } catch (error) {
      const errorMsg = error.response?.data?.error?.message || error.message;
      console.warn(`   ⚠️ Model ${modelName} failed: ${errorMsg}`);
      if (i === models.length - 1) {
        console.error('   ❌ All models failed for trend analysis. Falling back to default calculation rules.');
        // Fallback rule-based calculations if OpenRouter is completely down
        const smartMoneyOptionChange = fii.optionChange + pro.optionChange;
        if (fii.netOption + pro.netOption < -200000 || smartMoneyOptionChange < -50000) {
          direction = 'DOWN (BEARISH)';
          score = 'Strong Bearish 🚨';
          rationale = `FIIs aur Pros ne heavy options short kiye hain. Retail client opposite position long me hain. Yeh market down hone ka historical fallback indicator hai.`;
          fiiStance = 'BEARISH (SHORT)';
          proStance = 'BEARISH (SHORT)';
        } else if (fii.netOption + pro.netOption > 200000 || smartMoneyOptionChange > 50000) {
          direction = 'UP (BULLISH)';
          score = 'Strong Bullish 🚀';
          rationale = `FII aur Pro ne combined heavy call positions and long build-up kiya hai. Retail short hai. Yeh market upward momentum ka signal hai.`;
          fiiStance = 'BULLISH (LONG)';
          proStance = 'BULLISH (LONG)';
        } else {
          rationale = `FII aur Pro option positions mixed hain. Client aur Smart Money conflicting direction me evenly split hain, isliye market rangebound/sideways console karegi.`;
        }
      }
    }
  }

  // Update stances in returned objects
  fii.stance = fiiStance;
  pro.stance = proStance;
  client.stance = clientStance;

  return {
    fii,
    pro,
    client,
    smartMoneyOptionChange: fii.optionChange + pro.optionChange,
    smartMoneyFutureChange: fii.futureChange + pro.futureChange,
    direction,
    score,
    rationale,
    date: tradingDays[0].humanDate
  };
}

/**
 * Uses OpenRouter AI to analyze stock news and make Hinglish EXIT/CONTINUE decisions
 */
async function analyzeStockWithExitRule(stockName, buyPrice, newsArticles) {
  console.log(`🤖 AI analysis starting for "${stockName}"...`);
  
  let newsText = 'No news articles found from the past 7 days. Standard market evaluation requested.';
  if (newsArticles.length > 0) {
    newsText = newsArticles.map((art, idx) => {
      return `[Article ${idx + 1}] Source: ${art.source}\nTitle: ${art.title}\nDescription: ${art.description || 'No description'}\n`;
    }).join('\n');
  }

  const systemPrompt = `You are a professional financial advisor. Analyze the stock based on its news, quarterly results (Q1, Q2, Q3 if available), and purchase buy price.
IMPORTANT instructions for outputs:
1. All descriptions, summaries, future growth, and outlooks must be written in HINGLISH (Hindi written in English alphabet, e.g. "Stock me acchi growth dikh rahi hai...", "Company ke profits achhe hain isliye investment continue rakhna chahiye").
2. Explicitly cover any recent Q1, Q2, or Q3 earnings results/news if mentioned. Explain if the company has high upside growth potential (high target up) or if there are risks.
3. Perform a strict risk assessment. Recommending whether to HOLD (represented as "CONTINUE") or "EXIT" the stock. If there is BAD NEWS (scams, regulatory penalties, major profit declines, or poor earnings) set action to "EXIT". Otherwise, if future potential looks high or news/earnings are good/neutral, set action to "CONTINUE" (meaning HOLD).
4. Provide a suggested target exit/take-profit price in the 'exit_price' field (e.g. ₹250 or 'N/A').
5. Return ONLY a valid JSON object. Do not include markdown code block syntax like \`\`\`json.

JSON Schema:
{
  "summary": "2-3 sentence Hinglish news summary. Mention Q1/Q2/Q3 quarterly results if found in news.",
  "sentiment": "Bullish, Bearish, or Neutral",
  "geopolitical_risk": "Low, Medium, or High",
  "action": "CONTINUE or EXIT",
  "exit_price": "Suggested target price (e.g. ₹220 or 'N/A').",
  "future_growth": "1-2 sentence description in Hinglish of future potential, target upsides, and quarterly outlook.",
  "outlook": "2-3 sentence Hinglish rationale recommending if user should HOLD (CONTINUE) or EXIT (SELL) and why."
}`;

  const userPrompt = `Stock: ${stockName}
Buy Price: ${buyPrice || 'N/A'}
Recent news:\n${newsText}`;

  // Fallback free model sequence
  const models = [
    'google/gemma-4-31b-it:free',
    'google/gemma-4-26b-a4b-it:free',
    'openrouter/free',
    'meta-llama/llama-3.3-70b-instruct:free',
    'meta-llama/llama-3.2-3b-instruct:free',
    'qwen/qwen3-coder:free',
    'z-ai/glm-4.5-air:free',
    'nousresearch/hermes-3-llama-3.1-405b:free',
    'liquid/lfm-2.5-1.2b-instruct:free',
    'openai/gpt-oss-20b:free'
  ];

  for (let i = 0; i < models.length; i++) {
    const currentModel = models[i];
    console.log(`   🤖 Querying model: ${currentModel}...`);
    try {
      const response = await axios.post(CONFIG.openRouter.baseUrl, {
        model: currentModel,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
      }, {
        headers: {
          'Authorization': `Bearer ${CONFIG.openRouter.key}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': 'https://github.com/google/gemini-investment-tracker',
          'X-Title': 'Investment Tracker Bot'
        },
        timeout: 25000
      });

      if (response.data?.error) {
        throw new Error(`OpenRouter API Error: ${response.data.error.message || JSON.stringify(response.data.error)}`);
      }

      let content = response.data?.choices?.[0]?.message?.content;
      if (!content) {
        throw new Error('Empty response content from OpenRouter.');
      }

      // Robust JSON extraction
      let jsonMatch = content.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        throw new Error('No JSON object found in response.');
      }
      const cleaned = jsonMatch[0].trim();
      const analysis = JSON.parse(cleaned);
      console.log(`   ✅ Analysis completed successfully using model: ${currentModel}`);
      return {
        summary: analysis.summary || 'Summary generate nahi ho payi.',
        sentiment: analysis.sentiment || 'Neutral',
        geopolitical_risk: analysis.geopolitical_risk || 'Low',
        action: (analysis.action || 'CONTINUE').toUpperCase(),
        exit_price: analysis.exit_price || 'N/A',
        future_growth: analysis.future_growth || 'Growth details generate nahi ho payi.',
        outlook: analysis.outlook || 'Outlook generate nahi ho payi.'
      };
    } catch (error) {
      const errorMsg = error.response?.data?.error?.message || error.message;
      console.warn(`   ⚠️ Model ${currentModel} failed: ${errorMsg}`);
      if (i === models.length - 1) {
        console.error('   ❌ All fallback models failed.');
        return {
          summary: 'News fetch and check errors ki vajah se automatic update generate nahi ho payi.',
          sentiment: 'Neutral',
          geopolitical_risk: 'Medium',
          action: 'CONTINUE',
          exit_price: 'N/A',
          future_growth: 'API Connection failed, details available soon.',
          outlook: 'AI Analysis call errors aayi hai. Kripya OpenRouter key or connection settings check karein.'
        };
      }
      console.log('   🔄 Trying next model in fallback list...');
    }
  }
}

/**
 * Creates/Updates the formatted dated report sheet
 */
async function writeReportToExcel(workbook, reportData) {
  const dateStr = new Date().toISOString().split('T')[0];
  const sheetName = `Report_${dateStr}`;
  
  const existingSheet = workbook.getWorksheet(sheetName);
  if (existingSheet) {
    workbook.removeWorksheet(sheetName);
  }
  
  const sheet = workbook.addWorksheet(sheetName);
  sheet.views = [{ showGridLines: true }];
  
  // Set Columns
  sheet.columns = [
    { header: 'Stock Name', key: 'stockName', width: 25 },
    { header: 'Exchange', key: 'exchange', width: 15 },
    { header: 'Category', key: 'category', width: 15 },
    { header: 'Buy Price', key: 'buyPrice', width: 15 },
    { header: 'AI Recommendation / Action', key: 'action', width: 25 },
    { header: 'Suggested Exit Price', key: 'exitPrice', width: 22 },
    { header: 'Future Growth Prospect', key: 'futureGrowth', width: 35 },
    { header: 'Sentiment', key: 'sentiment', width: 15 },
    { header: 'Geopolitical Risk', key: 'geopoliticalRisk', width: 20 },
    { header: 'Hinglish News Summary', key: 'summary', width: 45 },
    { header: 'Hinglish Rationale & Outlook', key: 'outlook', width: 45 }
  ];
  
  // Format Header
  const headerRow = sheet.getRow(1);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Segoe UI' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' } // Sleek slate-900 color
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  
  // Fill data
  reportData.forEach((row, index) => {
    const excelRow = sheet.addRow({
      stockName: row.stockName,
      exchange: row.exchange || 'NSE',
      category: row.category || 'Indians',
      buyPrice: row.buyPrice || 'N/A',
      action: row.action,
      exitPrice: row.exit_price || 'N/A',
      futureGrowth: row.future_growth || 'N/A',
      sentiment: row.sentiment,
      geopoliticalRisk: row.geopolitical_risk,
      summary: row.summary,
      outlook: row.outlook
    });
    
    excelRow.height = 65; // Wrap text height
    
    excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font = { name: 'Segoe UI', size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      
      // Border
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
      
      // Zebra striping
      if (index % 2 === 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8FAFC' }
        };
      }
      
      // Format Action cell (Col 5)
      if (colNumber === 5) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { bold: true, name: 'Segoe UI', size: 10 };
        const act = row.action.toUpperCase();
        if (act === 'CONTINUE') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
          cell.font.color = { argb: 'FF065F46' };
        } else if (act === 'EXIT' || act === 'EXIT / SELL' || act === 'SELL') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          cell.font.color = { argb: 'FF991B1B' };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
          cell.font.color = { argb: 'FF475569' };
        }
      }

      // Format Buy Price cell (Col 4)
      if (colNumber === 4) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }

      // Format Suggested Exit Price (Col 6)
      if (colNumber === 6) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { bold: true, color: { argb: 'FF0284C7' } };
      }
      
      // Sentiment formatting (Col 8)
      if (colNumber === 8) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
      
      // Geopolitical Risk formatting (Col 9)
      if (colNumber === 9) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        const risk = row.geopolitical_risk.toUpperCase();
        if (risk === 'HIGH') {
          cell.font = { bold: true, color: { argb: 'FFEF4444' } };
        } else if (risk === 'MEDIUM') {
          cell.font = { bold: true, color: { argb: 'FFF59E0B' } };
        } else {
          cell.font = { color: { argb: 'FF10B981' } };
        }
      }
    });
  });
  
  await workbook.xlsx.writeFile(CONFIG.excelFile);
  console.log(`✅ Saved updated stock report sheet "${sheetName}".`);
}

/**
 * Uploads the workbook to tmpfiles.org
 */
async function uploadReport() {
  console.log('📤 Uploading Excel file to tmpfiles.org...');
  const form = new FormData();
  form.append('file', fs.createReadStream(CONFIG.excelFile));
  
  try {
    const response = await axios.post('https://tmpfiles.org/api/v1/upload', form, {
      headers: {
        ...form.getHeaders()
      }
    });
    
    if (response.data?.status === 'success' && response.data?.data?.url) {
      const dlUrl = response.data.data.url.replace('https://tmpfiles.org/', 'https://tmpfiles.org/dl/');
      console.log(`✅ File uploaded successfully! Link: ${dlUrl}`);
      return dlUrl;
    } else {
      console.error(`⚠️ Upload failed: ${JSON.stringify(response.data)}`);
      return null;
    }
  } catch (error) {
    console.error(`❌ Upload error: ${error.message}`);
    return null;
  }
}

/**
 * Sends a daily HTML email with FII/DII analysis and portfolio actions in HINGLISH
 */
async function sendEmailReport(reportData, fiiTrend, fileLink) {
  console.log(`📧 Sending report to ${CONFIG.email.receiver}...`);
  
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: CONFIG.email.sender,
      pass: CONFIG.email.password
    }
  });

  try {
    await transporter.verify();
    console.log('   SMTP verified.');
  } catch (err) {
    console.error('   ❌ SMTP Connection failed:', err.message);
    return false;
  }

  // Compile stock rows and exit warnings
  let stockRows = '';
  let exitAlerts = '';
  let exitCount = 0;
  
  reportData.forEach((row, index) => {
    let actColor = '#475569';
    let actBg = '#f1f5f9';
    if (row.action === 'CONTINUE') {
      actColor = '#065f46';
      actBg = '#d1fae5';
    } else if (row.action === 'EXIT' || row.action === 'SELL') {
      actColor = '#991b1b';
      actBg = '#fee2e2';
      exitCount++;
      exitAlerts += `
        <div style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 15px; border-radius: 8px; margin-bottom: 12px; font-size: 14px; line-height: 1.5; color: #7f1d1d;">
          <strong>🚨 EXIT Recommendation for: <span style="font-size: 15px; text-decoration: underline;">${row.stockName} (${row.exchange || 'NSE'})</span></strong><br/>
          <strong>Buy Price:</strong> ₹${row.buyPrice || 'N/A'} | <strong>Suggested Exit Price:</strong> <span style="font-weight: 800; color: #dc2626;">${row.exit_price || 'N/A'}</span><br/>
          <strong>Future Growth Prospect:</strong> ${row.future_growth || 'N/A'}<br/>
          <strong>Reason / Rationale (Hinglish):</strong> ${row.outlook || 'N/A'}
        </div>
      `;
    }
    
    const rowBg = index % 2 === 1 ? '#f8fafc' : '#ffffff';
    
    stockRows += `
      <tr style="background-color: ${rowBg}; border-bottom: 1px solid #cbd5e1;">
        <td style="padding: 12px; border: 1px solid #cbd5e1; font-weight: bold; color: #0f172a;">${row.stockName}</td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;"><span style="background-color: #dbeafe; color: #1e40af; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;">${row.exchange || 'NSE'}</span></td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;"><span style="background-color: #d1fae5; color: #065f46; padding: 3px 8px; border-radius: 4px; font-size: 11px;">${row.category || 'Indians'}</span></td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center; color: #0284c7; font-weight: 500;">${row.buyPrice ? '₹' + row.buyPrice : 'N/A'}</td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">
          <span style="background-color: ${actBg}; color: ${actColor}; padding: 6px 12px; border-radius: 9999px; font-weight: bold; font-size: 11px;">
            ${row.action}
          </span>
        </td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #dc2626;">${row.exit_price || 'N/A'}</td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; font-size: 12px; color: #334155;">${row.future_growth || 'N/A'}</td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center; color: #475569;">${row.sentiment}</td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: ${row.geopolitical_risk === 'High' ? '#ef4444' : row.geopolitical_risk === 'Medium' ? '#f59e0b' : '#10b981'};">
          ${row.geopolitical_risk}
        </td>
        <td style="padding: 12px; border: 1px solid #cbd5e1; font-size: 12px; color: #334155;">${row.summary}</td>
      </tr>
    `;
  });

  let exitSection = '';
  if (exitCount > 0) {
    exitSection = `
      <div style="margin-bottom: 30px;">
        <h3 style="color: #991b1b; margin-top: 0; font-size: 16px; border-bottom: 2px solid #fee2e2; padding-bottom: 8px;">🚨 Urgent Sell / Exit Recommendations</h3>
        ${exitAlerts}
      </div>
    `;
  } else {
    exitSection = `
      <div style="background-color: #f0fdf4; border-left: 4px solid #22c55e; padding: 15px; border-radius: 8px; margin-bottom: 30px; font-size: 14px; color: #166534;">
        <strong>✅ No Exit Recommendations:</strong> Aapke portfolio me koi bhi stock exit karne ki recommendation nahi hai. Sabhi stocks ko hold kiya jaa sakta hai.
      </div>
    `;
  }

  // Weekly Confirmation Section (Every Friday)
  const isFriday = new Date().getDay() === 5;
  let weeklyConfirmationSection = '';
  if (isFriday) {
    weeklyConfirmationSection = `
      <div style="background-color: #eff6ff; border: 1px dashed #3b82f6; padding: 20px; border-radius: 12px; text-align: center; margin-top: 25px; margin-bottom: 25px;">
        <h4 style="margin: 0 0 8px 0; color: #1e40af; font-size: 15px; font-weight: 700;">📋 Weekly Portfolio Confirmation Required</h4>
        <p style="color: #1e40af; font-size: 13px; margin: 0 0 15px 0; line-height: 1.5;">
          Aaj Friday hai. Kripya confirm karein ki kya aapne inme se kisi stock ko Exit/Sell kiya hai. Agar kiya hai, toh aap use direct mobile phone par niche di gayi Google Sheet link se delete kar sakte hain, taaki aane wale hafte me use download/track na kiya jaye.
        </p>
        <a href="https://docs.google.com/spreadsheets/d/1oDZbVB_zgXJ0OGlVDwHL8DWzPFbhfpTAtmILkf3Of1U/edit?usp=drivesdk" style="background-color: #3b82f6; color: #ffffff; padding: 8px 16px; border-radius: 6px; font-weight: bold; text-decoration: none; font-size: 12px; display: inline-block;">
          📝 Open Google Sheet to Edit Watchlist
        </a>
      </div>
    `;
  }

  const linkSection = `
    <div style="text-align: center; margin: 25px 0; display: flex; gap: 15px; justify-content: center; flex-wrap: wrap;">
      ${CONFIG.portalWebsiteUrl ? `
        <a href="${CONFIG.portalWebsiteUrl}" style="background-color: #0f172a; color: #ffffff; padding: 14px 28px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 16px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2); display: inline-block; border: 1px solid rgba(255,255,255,0.15);">
          🌐 Open Portfolio Website (GitHub Pages)
        </a>
      ` : ''}
      ${fileLink ? `
        <a href="${fileLink}" style="background-color: #2563eb; color: #ffffff; padding: 14px 28px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); display: inline-block;">
          📥 Download Updated Excel Report
        </a>
      ` : ''}
    </div>
    ${fileLink ? `<p style="color: #64748b; font-size: 12px; text-align: center; margin-top: -15px; margin-bottom: 25px;">Excel link valid for 24 hours: <a href="${fileLink}" style="color: #2563eb;">${fileLink}</a></p>` : ''}
  `;

  const htmlBody = `
    <div style="background-color: #f1f5f9; padding: 30px; font-family: 'Segoe UI', -apple-system, sans-serif;">
      <div style="max-width: 950px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
        
        <!-- Header Banner -->
        <div style="background: linear-gradient(135deg, #020617, #1e3a8a); padding: 40px 30px; text-align: center; color: #ffffff;">
          <h1 style="margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -0.5px;">📈 Daily Investment & FII/DII/Pro Market Analysis</h1>
          <p style="margin: 10px 0 0 0; color: #93c5fd; font-size: 15px; font-weight: 500;">Participant OI Trends & Geopolitical Risk Analysis</p>
        </div>
        
        <!-- Mail Intro in Hinglish -->
        <div style="padding: 35px 30px;">
          <p style="color: #334155; font-size: 15px; line-height: 1.6; margin-bottom: 25px; margin-top: 0;">
            Hi, <br/><br/>
            Aapki daily investment tracking report aur **FII / Pro / Client derivative data analysis** ready hai. Niche market positions aur predictions ki information table format me di gayi hai.
          </p>

          <!-- 🚨 Urgent Sell / Exit Recommendations -->
          ${exitSection}

          <!-- FII/DII Analysis Section -->
          <div style="background-color: #f8fafc; border-left: 4px solid #3b82f6; padding: 20px; border-radius: 0 12px 12px 0; margin-bottom: 30px;">
            <h3 style="color: #1e3a8a; margin: 0 0 15px 0; font-size: 17px; font-weight: 700;">📊 Participant wise Open Interest Summary (as on ${fiiTrend.date})</h3>
            
            <div style="margin-bottom: 20px; font-size: 14px; color: #334155;">
              <strong>Aaj Market View:</strong> 
              <span style="background-color: ${fiiTrend.direction.includes('UP') ? '#d1fae5' : fiiTrend.direction.includes('DOWN') ? '#fee2e2' : '#f1f5f9'}; color: ${fiiTrend.direction.includes('UP') ? '#065f46' : fiiTrend.direction.includes('DOWN') ? '#991b1b' : '#475569'}; padding: 4px 10px; border-radius: 4px; font-weight: bold; display: inline-block; margin-left: 5px;">
                ${fiiTrend.direction}
              </span> (${fiiTrend.score})
            </div>

            <div style="overflow-x: auto; margin-bottom: 15px;">
              <table style="width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; background-color: #ffffff;">
                <thead>
                  <tr style="background-color: #1e293b; color: #ffffff;">
                    <th style="padding: 10px; border: 1px solid #cbd5e1;">Participant</th>
                    <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">Net Futures (OI)</th>
                    <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">Futures Change</th>
                    <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">Net Options (OI)</th>
                    <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">Options Change</th>
                    <th style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">Daily Stance</th>
                  </tr>
                </thead>
                <tbody>
                  <!-- FII -->
                  <tr>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; font-weight: bold; color: #0f172a;">FII (Foreign Inst.)</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.fii.netFuture >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.fii.netFuture.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.fii.futureChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.fii.futureChange >= 0 ? '+' : ''}${fiiTrend.fii.futureChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.fii.netOption >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.fii.netOption.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.fii.optionChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.fii.optionChange >= 0 ? '+' : ''}${fiiTrend.fii.optionChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">
                      <span style="background-color: ${fiiTrend.fii.stance.includes('BULLISH') ? '#d1fae5' : fiiTrend.fii.stance.includes('BEARISH') ? '#fee2e2' : '#f1f5f9'}; color: ${fiiTrend.fii.stance.includes('BULLISH') ? '#065f46' : fiiTrend.fii.stance.includes('BEARISH') ? '#991b1b' : '#475569'}; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 10px;">
                        ${fiiTrend.fii.stance}
                      </span>
                    </td>
                  </tr>
                  <!-- PRO -->
                  <tr style="background-color: #f8fafc;">
                    <td style="padding: 10px; border: 1px solid #cbd5e1; font-weight: bold; color: #0f172a;">PRO (Proprietary)</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.pro.netFuture >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.pro.netFuture.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.pro.futureChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.pro.futureChange >= 0 ? '+' : ''}${fiiTrend.pro.futureChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.pro.netOption >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.pro.netOption.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.pro.optionChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.pro.optionChange >= 0 ? '+' : ''}${fiiTrend.pro.optionChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">
                      <span style="background-color: ${fiiTrend.pro.stance.includes('BULLISH') ? '#d1fae5' : fiiTrend.pro.stance.includes('BEARISH') ? '#fee2e2' : '#f1f5f9'}; color: ${fiiTrend.pro.stance.includes('BULLISH') ? '#065f46' : fiiTrend.pro.stance.includes('BEARISH') ? '#991b1b' : '#475569'}; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 10px;">
                        ${fiiTrend.pro.stance}
                      </span>
                    </td>
                  </tr>
                  <!-- CLIENT -->
                  <tr>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; font-weight: bold; color: #0f172a;">CLIENT (Retailers)</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.client.netFuture >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.client.netFuture.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.client.futureChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.client.futureChange >= 0 ? '+' : ''}${fiiTrend.client.futureChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.client.netOption >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.client.netOption.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; font-weight: bold; color: ${fiiTrend.client.optionChange >= 0 ? '#10b981' : '#ef4444'}">${fiiTrend.client.optionChange >= 0 ? '+' : ''}${fiiTrend.client.optionChange.toLocaleString()}</td>
                    <td style="padding: 10px; border: 1px solid #cbd5e1; text-align: center;">
                      <span style="background-color: ${fiiTrend.client.stance.includes('BULLISH') ? '#d1fae5' : fiiTrend.client.stance.includes('BEARISH') ? '#fee2e2' : '#f1f5f9'}; color: ${fiiTrend.client.stance.includes('BULLISH') ? '#065f46' : fiiTrend.client.stance.includes('BEARISH') ? '#991b1b' : '#475569'}; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 10px;">
                        ${fiiTrend.client.stance}
                      </span>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            
            <p style="color: #475569; font-size: 14px; line-height: 1.6; margin: 0; padding-top: 10px; border-top: 1px solid #e2e8f0;">
              <strong>Market Rationale (Hinglish):</strong> ${fiiTrend.rationale}
            </p>
          </div>
          
          <!-- Summary Table -->
          <h3 style="color: #0f172a; margin-bottom: 15px; font-size: 16px; border-bottom: 2px solid #f1f5f9; padding-bottom: 8px;">💼 Portfolio Stock Tracker</h3>
          <div style="overflow-x: auto; margin-bottom: 30px;">
            <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; border: 1px solid #cbd5e1;">
              <thead>
                <tr style="background-color: #0f172a; color: #ffffff;">
                  <th style="padding: 12px; border: 1px solid #cbd5e1;">Stock</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Exchange</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Category</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Buy Price</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Recommendation</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Suggested Exit Price</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1;">Future Growth</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Sentiment</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1; text-align: center;">Geopolitical Risk</th>
                  <th style="padding: 12px; border: 1px solid #cbd5e1;">News Summary (Hinglish)</th>
                </tr>
              </thead>
              <tbody>
                ${stockRows}
              </tbody>
            </table>
          </div>
          
          <!-- Download & Website Buttons -->
          ${linkSection}

          <!-- Weekly Friday Confirmation -->
          ${weeklyConfirmationSection}
          
          <!-- Interactive Settings Portal -->
          <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 25px; border-radius: 12px; text-align: center; margin-top: 25px; box-shadow: inset 0 2px 4px rgba(0,0,0,0.02);">
            <h4 style="margin: 0 0 10px 0; color: #1e293b; font-size: 15px; font-weight: 700;">⚙️ Yahan par aap apne stocks add ya delete kar sakte hain:</h4>
            <p style="color: #475569; font-size: 13px; margin: 0 0 20px 0; line-height: 1.5;">
              Niche diye gaye options se aap naye stocks list me add kar sakte hain taaki unki news aur report aati rahe:
            </p>
            <div style="display: flex; gap: 15px; justify-content: center; flex-wrap: wrap;">
              ${CONFIG.mobilePortalUrl ? `
                <a href="${CONFIG.mobilePortalUrl}" style="background-color: #2563eb; color: #ffffff; padding: 12px 22px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 14px; display: inline-block; box-shadow: 0 2px 4px rgba(37, 99, 235, 0.3);">
                  📲 Open Mobile Portfolio App (Google Web App)
                </a>
              ` : `
                <div style="background-color: #f1f5f9; color: #64748b; padding: 12px 22px; border-radius: 8px; font-weight: bold; font-size: 14px; display: inline-block; border: 1px dashed #cbd5e1;">
                  📲 Mobile App Not Configured (.env PORTAL_MOBILE_URL empty)
                </div>
              `}
              ${CONFIG.portalWebsiteUrl ? `
                <a href="${CONFIG.portalWebsiteUrl}" style="background-color: #0f172a; color: #ffffff; padding: 12px 22px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 14px; display: inline-block; box-shadow: 0 2px 4px rgba(15, 23, 42, 0.2);">
                  🌐 Open Portfolio Website (GitHub Pages)
                </a>
              ` : ''}
              <a href="https://docs.google.com/spreadsheets/d/1oDZbVB_zgXJ0OGlVDwHL8DWzPFbhfpTAtmILkf3Of1U/edit?usp=drivesdk" style="background-color: #22c55e; color: #ffffff; padding: 12px 22px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 14px; display: inline-block; box-shadow: 0 2px 4px rgba(34, 197, 94, 0.2);">
                🟢 Open Google Sheet Database (Works on Phone)
              </a>
            </div>
          </div>
        </div>
        
        <!-- Footer -->
        <div style="background-color: #f8fafc; padding: 25px 30px; text-align: center; border-top: 1px solid #f1f5f9;">
          <p style="margin: 0; color: #64748b; font-size: 11px;">This is an automated investment analysis bot using OpenRouter AI & NSE data.</p>
          <p style="margin: 4px 0 0 0; color: #94a3b8; font-size: 11px;">Receiver Email: ${CONFIG.email.receiver}</p>
        </div>
      </div>
    </div>
  `;

  try {
    const mailOptions = {
      from: `"📈 Market & Portfolio Update" <${CONFIG.email.sender}>`,
      to: CONFIG.email.receiver,
      subject: `📈 Daily Market Update: View is ${fiiTrend.direction.split(' ')[0]} | ${fiiTrend.humanDate || fiiTrend.date}`,
      html: htmlBody,
      attachments: [
        {
          filename: `stocks_report_${new Date().toISOString().split('T')[0]}.xlsx`,
          path: CONFIG.excelFile
        }
      ]
    };

    const info = await transporter.sendMail(mailOptions);
    console.log(`✅ Mail sent successfully! Message ID: ${info.messageId}`);
    return true;
  } catch (error) {
    console.error(`❌ Failed to send email: ${error.message}`);
    return false;
  }
}

/**
 * Main Orchestration Process
 */
async function runTracker() {
  console.log('\n=========================================');
  console.log('🚀 Starting Investment Tracker v2.0 Daily Process');
  console.log('=========================================\n');

  try {
    // 1. Fetch NSE trading days
    const tradingDays = await fetchLast3TradingDays();
    
    // 2. Perform FII & Pro Option Trend analysis
    const fiiTrend = await analyzeFiiProMarketTrend(tradingDays);
    console.log(`\n📊 Smart Money Trend Prediction: ${fiiTrend.direction} (FII Net Options: ${fiiTrend.fii.netOption.toLocaleString()}, Pro Net Options: ${fiiTrend.pro.netOption.toLocaleString()})`);
    
    // 3. Download Google Sheets template
    const tempFile = await downloadTemplate();
    const workbook = new ExcelJS.Workbook();
    let loadedFromGoogle = false;
    
    if (tempFile) {
      await workbook.xlsx.readFile(tempFile);
      loadedFromGoogle = true;
      // Delete temporary download file
      try { fs.unlinkSync(tempFile); } catch (e) {}
    } else if (fs.existsSync(CONFIG.excelFile)) {
      console.log('📖 Using existing local stocks.xlsx as base.');
      await workbook.xlsx.readFile(CONFIG.excelFile);
    } else {
      throw new Error('Google sheet template download failed and no local stocks.xlsx was found.');
    }

    // 4. Load watchlist from workbook (prioritizing Google Sheets if loaded)
    let stocksToProcess = getPortfolioFromWorkbook(workbook);
    if (stocksToProcess.length === 0) {
      console.log('📖 Loading watchlist from local stocks.xlsx database...');
      // Try to load from local file
      const localPortfolio = await getSavedPortfolio();
      if (localPortfolio.length > 0) {
        stocksToProcess = localPortfolio;
      } else {
        console.log('📖 Watchlist empty. Using default stock tracking list.');
        stocksToProcess = [
          { stockName: 'Apple Inc. (AAPL)', buyPrice: 175, exchange: 'NSE', category: 'Indians' },
          { stockName: 'Tesla, Inc. (TSLA)', buyPrice: 180, exchange: 'NSE', category: 'Indians' },
          { stockName: 'NVIDIA Corporation (NVDA)', buyPrice: 900, exchange: 'NSE', category: 'Indians' },
          { stockName: 'Microsoft Corporation (MSFT)', buyPrice: 420, exchange: 'NSE', category: 'Indians' },
          { stockName: 'Reliance Industries Limited', buyPrice: 2400, exchange: 'NSE', category: 'Indians' },
          { stockName: 'Tata Consultancy Services (TCS)', buyPrice: 3800, exchange: 'NSE', category: 'Indians' }
        ];
      }
    } else {
      console.log(`📋 Loaded ${stocksToProcess.length} stocks from downloaded Google Sheet template.`);
    }
    
    // 5. Update data Sheet with CSV data
    let dataSheet = workbook.getWorksheet('data');
    if (!dataSheet) {
      dataSheet = workbook.addWorksheet('data');
    }
    writeDataToDataSheet(dataSheet, tradingDays);
    
    // 6. Correct/Update calculations in overall data sheet
    let overallSheet = workbook.getWorksheet('overall data');
    if (!overallSheet) {
      overallSheet = workbook.addWorksheet('overall data');
    }
    writeCalculationsToOverallSheet(overallSheet);

    // 7. Re-write Stocks sheet preserving user portfolio details
    let stocksSheet = workbook.getWorksheet('Stocks');
    if (stocksSheet) {
      workbook.removeWorksheet('Stocks');
    }
    stocksSheet = workbook.addWorksheet('Stocks');
    stocksSheet.columns = [
      { header: 'Stock Name (Symbol)', key: 'stockName', width: 35 },
      { header: 'Buy Price', key: 'buyPrice', width: 15 },
      { header: 'Exchange', key: 'exchange', width: 15 },
      { header: 'Category', key: 'category', width: 15 }
    ];
    
    // Style Stocks Header
    const headerRow = stocksSheet.getRow(1);
    headerRow.height = 25;
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F497D' }
      };
    });
    
    // Fill Stocks rows using loaded list
    stocksToProcess.forEach(s => {
      stocksSheet.addRow({
        stockName: s.stockName,
        buyPrice: s.buyPrice,
        exchange: s.exchange || 'NSE',
        category: s.category || 'Indians'
      });
    });
    stocksSheet.views = [{ showGridLines: true }];
    console.log(`📋 Preserved portfolio watchlists: ${stocksToProcess.length} stocks processed.`);

    // 8. Analyze news and details for each stock
    const reportData = [];
    for (const stock of stocksToProcess) {
      console.log(`\n-----------------------------------------`);
      // A. Fetch news
      const newsArticles = await fetchStockNews(stock.stockName);
      
      // B. Analyze with OpenRouter
      const analysis = await analyzeStockWithExitRule(stock.stockName, stock.buyPrice, newsArticles);
      
      reportData.push({
        stockName: stock.stockName,
        buyPrice: stock.buyPrice,
        exchange: stock.exchange || 'NSE',
        category: stock.category || 'Indians',
        ...analysis
      });
    }

    console.log(`\n-----------------------------------------`);
    // 9. Write report sheet to Excel workbook
    await writeReportToExcel(workbook, reportData);
    
    // 10. Upload file for sharing link
    const fileLink = await uploadReport();
    
    // 11. Send email report
    const emailSent = await sendEmailReport(reportData, fiiTrend, fileLink);
    
    if (emailSent) {
      console.log('\n=========================================');
      console.log('✅ Daily Investment Tracker Bot completed successfully!');
      console.log('=========================================\n');
    } else {
      console.log('\n=========================================');
      console.log('⚠️ Process finished but email dispatch failed.');
      console.log('=========================================\n');
    }

  } catch (error) {
    console.error('\n❌ Tracker crashed due to error:', error.stack || error.message);
  }
}

// Helper to fetch news (copied from previous code for standalone run support)
async function fetchStockNews(stockName) {
  const query = getSearchQuery(stockName);
  const date7DaysAgo = new Date();
  date7DaysAgo.setDate(date7DaysAgo.getDate() - 7);
  const fromDateStr = date7DaysAgo.toISOString().split('T')[0];
  
  console.log(`🔍 Fetching news for "${stockName}" since ${fromDateStr}...`);
  
  try {
    const response = await axios.get(CONFIG.newsApi.baseUrl, {
      params: {
        q: query,
        from: fromDateStr,
        sortBy: 'relevance',
        language: 'en',
        pageSize: 5,
        apiKey: CONFIG.newsApi.key
      },
      headers: {
        'User-Agent': 'InvestmentTrackerBot/2.0'
      }
    });
    
    if (response.data.status === 'ok') {
      const articles = response.data.articles || [];
      console.log(`   Found ${articles.length} news articles.`);
      return articles.map(art => ({
        title: art.title,
        description: art.description,
        source: art.source?.name || 'Unknown',
        url: art.url,
        publishedAt: art.publishedAt
      }));
    } else {
      console.warn(`   ⚠️ News API returned error: ${response.data.message}`);
      return [];
    }
  } catch (error) {
    console.error(`   ❌ Failed to fetch news: ${error.response?.data?.message || error.message}`);
    return [];
  }
}

// Run process
runTracker();
