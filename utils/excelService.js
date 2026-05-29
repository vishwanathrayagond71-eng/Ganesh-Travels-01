// ============================================================
// excelService.js - Hybrid Database (Excel / MongoDB) Service
// Handles transparent data operations for local Excel & MongoDB Atlas
// ============================================================

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { MongoClient } = require('mongodb');

// Path to excel-data folder
const DATA_DIR = (process.env.NETLIFY || process.env.LAMBDA_TASK_ROOT)
  ? '/tmp/excel-data'
  : path.join(__dirname, '..', 'excel-data');

// Make sure excel-data folder exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Define headers for each Excel sheet / collection
const SHEET_HEADERS = {
  users: ['id', 'name', 'email', 'phone', 'password', 'createdAt'],
  bookings: ['id', 'userId', 'userName', 'userEmail', 'destination', 'date', 'guests', 'name', 'email', 'phone', 'message', 'status', 'createdAt'],
  reviews: ['id', 'userEmail', 'userName', 'destination', 'rating', 'comment', 'status', 'createdAt'],
  contacts: ['id', 'name', 'email', 'subject', 'message', 'createdAt'],
  newsletter: ['id', 'email', 'createdAt'],
  settings: ['key', 'value', 'createdAt'],
  team: ['id', 'name', 'role', 'image', 'createdAt']
};

const MONGODB_URI = process.env.MONGODB_URI;
let mongoClient = null;
let db = null;
const isMongo = !!MONGODB_URI;

// Helper to get local excel file path
function getFilePath(sheetName) {
  return path.join(DATA_DIR, `${sheetName}.xlsx`);
}

// Initialize an excel file with headers if it doesn't exist
async function initSheet(sheetName, customWorkbook) {
  const filePath = getFilePath(sheetName);
  if (!fs.existsSync(filePath) || customWorkbook) {
    const workbook = customWorkbook || new ExcelJS.Workbook();
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      sheet = workbook.addWorksheet(sheetName);
    }
    const headers = SHEET_HEADERS[sheetName];
    sheet.spliceRows(1, sheet.rowCount); // Clear sheet rows if rewriting
    sheet.addRow(headers);
    // Style the header row
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1d3557' }
    };
    if (!customWorkbook) {
      await workbook.xlsx.writeFile(filePath);
    }
  }
}

// Sync MongoDB collection to Excel file in the background (for Admin Panel downloads)
async function syncToExcel(sheetName) {
  if (!isMongo) return;
  try {
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await initSheet(sheetName, workbook);
    const sheet = workbook.getWorksheet(sheetName);
    const headers = SHEET_HEADERS[sheetName];
    
    const collection = db.collection(sheetName);
    const docs = await collection.find({}).toArray();
    
    docs.forEach(doc => {
      const rowData = headers.map(h => doc[h] || '');
      sheet.addRow(rowData);
    });
    
    await workbook.xlsx.writeFile(filePath);
  } catch (err) {
    console.error(`[Sync Error] Failed to sync ${sheetName} to Excel:`, err);
  }
}

// Initialize database / sheets
async function initAllSheets() {
  if (isMongo) {
    console.log('🔌 Connecting to MongoDB Atlas...');
    try {
      mongoClient = new MongoClient(MONGODB_URI);
      await mongoClient.connect();
      db = mongoClient.db();
      console.log(`✅ Connected successfully to MongoDB database: "${db.databaseName}"`);

      // Seed settings if empty
      const phone = await findOne('settings', 'key', 'contact_phone');
      if (!phone) {
        await addRow('settings', { key: 'contact_phone', value: '+91 98765 43210', createdAt: new Date().toLocaleString() });
      }
      const email = await findOne('settings', 'key', 'contact_email');
      if (!email) {
        await addRow('settings', { key: 'contact_email', value: 'xyz7@gmail.com', createdAt: new Date().toLocaleString() });
      }

      // Start background sync of all collections to local Excel files for download support
      for (const sheetName of Object.keys(SHEET_HEADERS)) {
        syncToExcel(sheetName);
      }
    } catch (err) {
      console.error('❌ Failed to connect to MongoDB:', err);
      throw err;
    }
  } else {
    // Fallback Excel initialization
    for (const sheetName of Object.keys(SHEET_HEADERS)) {
      await initSheet(sheetName);
    }

    const phone = await findOne('settings', 'key', 'contact_phone');
    if (!phone) {
      await addRow('settings', { key: 'contact_phone', value: '+91 98765 43210', createdAt: new Date().toLocaleString() });
    }
    const email = await findOne('settings', 'key', 'contact_email');
    if (!email) {
      await addRow('settings', { key: 'contact_email', value: 'xyz7@gmail.com', createdAt: new Date().toLocaleString() });
    }
    console.log('✅ Excel sheets initialized successfully (Local Mode)');
  }
}

// Read all rows (returns array of objects)
async function readData(sheetName) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const docs = await collection.find({}).toArray();
      return docs.map(doc => {
        const obj = { ...doc };
        delete obj._id; // Remove MongoDB internal ID for application consistency
        return obj;
      });
    } catch (err) {
      console.error(`[DB Error] Failed to read from collection ${sheetName}:`, err);
      return [];
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    const rows = [];
    const headers = SHEET_HEADERS[sheetName];
    
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row.getCell(i + 1).value || '';
      });
      rows.push(obj);
    });
    
    return rows;
  }
}

// Add a new row / document
async function addRow(sheetName, data) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      await collection.insertOne({ ...data });
      // Sync to local Excel in background for backup/download
      syncToExcel(sheetName);
    } catch (err) {
      console.error(`[DB Error] Failed to insert into collection ${sheetName}:`, err);
      throw err;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    const headers = SHEET_HEADERS[sheetName];
    const rowData = headers.map(h => data[h] || '');
    sheet.addRow(rowData);
    
    await workbook.xlsx.writeFile(filePath);
  }
}

// Delete a row / document by ID
async function deleteRow(sheetName, id) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const query = (sheetName === 'settings') ? { key: id } : { id: id };
      const result = await collection.deleteOne(query);
      syncToExcel(sheetName);
      return result.deletedCount > 0;
    } catch (err) {
      console.error(`[DB Error] Failed to delete from collection ${sheetName}:`, err);
      return false;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    
    let rowIndexToDelete = -1;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (String(row.getCell(1).value) === String(id)) {
        rowIndexToDelete = rowNumber;
      }
    });
    
    if (rowIndexToDelete > 0) {
      sheet.spliceRows(rowIndexToDelete, 1);
      await workbook.xlsx.writeFile(filePath);
      return true;
    }
    return false;
  }
}

// Update a row / document by ID
async function updateRow(sheetName, id, newData) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      const query = (sheetName === 'settings') ? { key: id } : { id: id };
      await collection.updateOne(query, { $set: newData });
      syncToExcel(sheetName);
    } catch (err) {
      console.error(`[DB Error] Failed to update collection ${sheetName}:`, err);
      throw err;
    }
  } else {
    await initSheet(sheetName);
    const filePath = getFilePath(sheetName);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    const headers = SHEET_HEADERS[sheetName];
    
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (String(row.getCell(1).value) === String(id)) {
        headers.forEach((header, i) => {
          if (newData[header] !== undefined) {
            row.getCell(i + 1).value = newData[header];
          }
        });
      }
    });
    
    await workbook.xlsx.writeFile(filePath);
  }
}

// Find a single row / document by a field value
async function findOne(sheetName, field, value) {
  if (isMongo) {
    try {
      const collection = db.collection(sheetName);
      // Escape special regex characters in search term for safety
      const escapedValue = String(value).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      // Case-insensitive exact match
      const query = { [field]: { $regex: new RegExp(`^${escapedValue}$`, 'i') } };
      const doc = await collection.findOne(query);
      if (doc) {
        const obj = { ...doc };
        delete obj._id;
        return obj;
      }
      return null;
    } catch (err) {
      console.error(`[DB Error] Failed to find in collection ${sheetName}:`, err);
      return null;
    }
  } else {
    const rows = await readData(sheetName);
    return rows.find(row => String(row[field]).toLowerCase() === String(value).toLowerCase()) || null;
  }
}

// Generate Excel workbook in-memory and write to stream
async function writeExcelToStream(sheetName, stream) {
  const headers = SHEET_HEADERS[sheetName];
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  
  sheet.addRow(headers);
  // Style the header row
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1d3557' }
  };
  
  const data = await readData(sheetName);
  data.forEach(row => {
    const rowData = headers.map(h => row[h] !== undefined ? String(row[h]) : '');
    sheet.addRow(rowData);
  });
  
  await workbook.xlsx.write(stream);
}

module.exports = {
  initAllSheets,
  readData,
  addRow,
  deleteRow,
  updateRow,
  findOne,
  getFilePath,
  writeExcelToStream
};
