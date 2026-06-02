const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const app = express();
const PORT = 3000;
const EXCEL_FILE = path.join(__dirname, 'stocks.xlsx');

app.use(bodyParser.json());
app.use(express.static(__dirname));

// Ensure workbook exists with 4-column structure
async function getOrCreateWorkbook() {
  const workbook = new ExcelJS.Workbook();
  if (!fs.existsSync(EXCEL_FILE)) {
    const sheet = workbook.addWorksheet('Stocks');
    sheet.columns = [
      { header: 'Stock Name (Symbol)', key: 'stockName', width: 35 },
      { header: 'Buy Price', key: 'buyPrice', width: 15 },
      { header: 'Exchange', key: 'exchange', width: 15 },
      { header: 'Category', key: 'category', width: 15 }
    ];
    // Add default indian-focused demo portfolio
    sheet.addRow({ stockName: 'Apple Inc. (AAPL)', buyPrice: 175, exchange: 'NSE', category: 'Indians' });
    sheet.addRow({ stockName: 'Tesla, Inc. (TSLA)', buyPrice: 180, exchange: 'NSE', category: 'Indians' });
    sheet.addRow({ stockName: 'NVIDIA Corporation (NVDA)', buyPrice: 900, exchange: 'NSE', category: 'Indians' });
    sheet.addRow({ stockName: 'Reliance Industries Limited', buyPrice: 2400, exchange: 'NSE', category: 'Indians' });
    sheet.addRow({ stockName: 'Tata Consultancy Services (TCS)', buyPrice: 3800, exchange: 'NSE', category: 'Indians' });
    
    // Style Header
    const headerRow = sheet.getRow(1);
    headerRow.height = 25;
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F497D' }
      };
    });
    
    await workbook.xlsx.writeFile(EXCEL_FILE);
  } else {
    await workbook.xlsx.readFile(EXCEL_FILE);
    // Ensure Stocks sheet exists
    let sheet = workbook.getWorksheet('Stocks');
    if (!sheet) {
      sheet = workbook.addWorksheet('Stocks');
      sheet.columns = [
        { header: 'Stock Name (Symbol)', key: 'stockName', width: 35 },
        { header: 'Buy Price', key: 'buyPrice', width: 15 },
        { header: 'Exchange', key: 'exchange', width: 15 },
        { header: 'Category', key: 'category', width: 15 }
      ];
      await workbook.xlsx.writeFile(EXCEL_FILE);
    }
  }
  return workbook;
}

// GET /api/stocks
app.get('/api/stocks', async (req, res) => {
  try {
    const workbook = await getOrCreateWorkbook();
    const sheet = workbook.getWorksheet('Stocks');
    const stocks = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const stockName = row.getCell(1).value;
        const buyPrice = row.getCell(2).value;
        const exchange = row.getCell(3).value || 'NSE';
        const category = row.getCell(4).value || 'Indians';
        if (stockName) {
          stocks.push({
            stockName: stockName.toString().trim(),
            buyPrice: buyPrice ? parseFloat(buyPrice) : null,
            exchange: exchange.toString().trim(),
            category: category.toString().trim()
          });
        }
      }
    });
    res.json(stocks);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST /api/stocks
app.post('/api/stocks', async (req, res) => {
  try {
    const newStocks = req.body; // Array of { stockName, buyPrice, exchange, category }
    const workbook = new ExcelJS.Workbook();
    
    // Read or create
    if (fs.existsSync(EXCEL_FILE)) {
      await workbook.xlsx.readFile(EXCEL_FILE);
    } else {
      workbook.addWorksheet('Stocks');
    }
    
    // Get or recreate Stocks sheet
    let sheet = workbook.getWorksheet('Stocks');
    if (sheet) {
      workbook.removeWorksheet('Stocks');
    }
    sheet = workbook.addWorksheet('Stocks');
    sheet.columns = [
      { header: 'Stock Name (Symbol)', key: 'stockName', width: 35 },
      { header: 'Buy Price', key: 'buyPrice', width: 15 },
      { header: 'Exchange', key: 'exchange', width: 15 },
      { header: 'Category', key: 'category', width: 15 }
    ];
    
    // Style Header
    const headerRow = sheet.getRow(1);
    headerRow.height = 25;
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F497D' }
      };
    });

    // Add new data rows
    newStocks.forEach(s => {
      sheet.addRow({
        stockName: s.stockName,
        buyPrice: s.buyPrice ? parseFloat(s.buyPrice) : null,
        exchange: s.exchange || 'NSE',
        category: s.category || 'Indians'
      });
    });

    sheet.views = [{ showGridLines: true }];
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST /api/run - trigger node tracker.js and stream output
app.post('/api/run', (req, res) => {
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Transfer-Encoding', 'chunked');

  console.log('🤖 Manual Run requested via Web UI...');
  res.write('🤖 Automation starting inside background process...\n\n');

  const processRun = spawn('node', ['tracker.js'], { cwd: __dirname });

  processRun.stdout.on('data', (data) => {
    res.write(data.toString());
  });

  processRun.stderr.on('data', (data) => {
    res.write(`⚠️ ERROR: ${data.toString()}`);
  });

  processRun.on('close', (code) => {
    res.write(`\n\n=========================================\n`);
    res.write(`🏁 Process closed with exit code: ${code}\n`);
    if (code === 0) {
      res.write(`✅ Report compiled and emailed successfully!\n`);
    } else {
      res.write(`❌ Process failed. Verify configurations and try again.\n`);
    }
    res.write(`=========================================\n`);
    res.end();
  });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log('\n======================================================');
  console.log(`📈 Portfolio manager is running at: http://localhost:${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser to update watchlists.`);
  console.log('======================================================\n');
});
