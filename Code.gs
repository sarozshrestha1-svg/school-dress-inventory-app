const SHEET_NAMES = {
  inventory: 'Inventory',
  sales: 'Sales',
  stock: 'Stock Summary',
  credit: 'Credit Sales',
  pending: 'Pending Deliveries',
  notifications: 'Notifications',
  dashboard: 'Dashboard Data'
};

const HEADERS = {
  inventory: ['Date Received', 'Supplier Name', 'School Name', 'Item Name', 'Size', 'Quantity Received', 'Remarks', 'Created At'],
  sales: ['Sale Date', 'Student Name', 'School Name', 'Item Name', 'Size', 'Quantity Sold', 'Selling Price', 'Additional Cost', 'Total Bill', 'Amount Paid', 'Remaining Amount', 'Credit Sale', 'Phone No.', 'Item Left To Give', 'Notes', 'Created At'],
  stock: ['School Name', 'Item Name', 'Size', 'Available Quantity', 'Status', 'Last Updated'],
  credit: ['Student Name', 'School Name', 'Total Bill', 'Amount Paid', 'Remaining Amount', 'Sale Date', 'Phone No.', 'Notes'],
  pending: ['Student Name', 'School Name', 'Item Sold', 'Pending Item', 'Sale Date', 'Phone No.', 'Notes'],
  notifications: ['Date', 'School Name', 'Item Name', 'Size', 'Current Stock Level', 'Status'],
  dashboard: ['Metric', 'Value', 'Last Updated']
};

function doGet(e) {
  try {
    setupSpreadsheet();
    const params = e && e.parameter ? e.parameter : {};
    if (params.payload) {
      const body = JSON.parse(params.payload || '{}');
      const result = handleAction(body);
      return scriptResponse(result, params.callback);
    }
    return scriptResponse({ ok: true, message: 'School Dress Inventory backend is running.' }, params.callback);
  } catch (error) {
    return scriptResponse({ ok: false, message: error.message }, e && e.parameter && e.parameter.callback);
  }
}

function doPost(e) {
  try {
    setupSpreadsheet();
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return jsonResponse(handleAction(body));
  } catch (error) {
    return jsonResponse({ ok: false, message: error.message });
  }
}

function handleAction(body) {
  const action = body.action;
  if (action === 'getAllData') return getAllData();
  if (action === 'addInventory') return addInventory(body.inventory || {});
  if (action === 'addSale') return addSale(body.sale || {});
  return { ok: false, message: 'Unknown action.' };
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach(function(key) {
    const name = SHEET_NAMES[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaders(sheet, HEADERS[key]);
  });
  rebuildDerivedSheets();
}

function ensureHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const missing = headers.some(function(header, index) {
    return current[index] !== header;
  });
  if (missing) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function addInventory(inventory) {
  validateRequired(inventory, ['dateReceived', 'supplierName', 'schoolName', 'itemName', 'size', 'quantityReceived']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.inventory);
    sheet.appendRow([
      cleanDate(inventory.dateReceived),
      text(inventory.supplierName),
      text(inventory.schoolName),
      text(inventory.itemName),
      text(inventory.size),
      number(inventory.quantityReceived),
      text(inventory.remarks),
      new Date()
    ]);
    rebuildDerivedSheets();
    notifyLowStockForItem(text(inventory.schoolName), text(inventory.itemName), text(inventory.size));
    return { ok: true, message: 'Inventory added.' };
  } finally {
    lock.releaseLock();
  }
}

function addSale(sale) {
  validateRequired(sale, ['saleDate', 'studentName', 'schoolName', 'itemName', 'size', 'quantitySold', 'sellingPrice', 'amountPaid']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const quantitySold = number(sale.quantitySold);
    const sellingPrice = number(sale.sellingPrice);
    const additionalCost = number(sale.additionalCost);
    const totalBill = quantitySold * sellingPrice + additionalCost;
    const amountPaid = number(sale.amountPaid);
    const remainingAmount = Math.max(totalBill - amountPaid, 0);
    const creditSale = remainingAmount > 0 ? 'Yes' : text(sale.creditSale || 'No');
    const pendingItem = text(sale.itemLeftToGive || 'Nothing Pending');
    const available = getAvailableStock(text(sale.schoolName), text(sale.itemName), text(sale.size));
    if (quantitySold > available) {
      throw new Error('Only ' + available + ' item(s) available for this school, item, and size.');
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.sales);
    sheet.appendRow([
      cleanDate(sale.saleDate),
      text(sale.studentName),
      text(sale.schoolName),
      text(sale.itemName),
      text(sale.size),
      quantitySold,
      sellingPrice,
      additionalCost,
      totalBill,
      amountPaid,
      remainingAmount,
      creditSale,
      text(sale.phoneNo),
      pendingItem,
      text(sale.notes),
      new Date()
    ]);
    rebuildDerivedSheets();
    notifyLowStockForItem(text(sale.schoolName), text(sale.itemName), text(sale.size));
    return { ok: true, message: 'Sale added.' };
  } finally {
    lock.releaseLock();
  }
}

function rebuildDerivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockRows = buildStockSummaryRows();
  replaceData(SHEET_NAMES.stock, HEADERS.stock, stockRows);

  const sales = readObjects(SHEET_NAMES.sales, HEADERS.sales);
  const creditRows = [];
  const pendingRows = [];
  sales.forEach(function(row) {
    if (number(row['Remaining Amount']) > 0 || row['Credit Sale'] === 'Yes') {
      creditRows.push([
        row['Student Name'],
        row['School Name'],
        number(row['Total Bill']),
        number(row['Amount Paid']),
        number(row['Remaining Amount']),
        row['Sale Date'],
        row['Phone No.'],
        row['Notes']
      ]);
    }
    if (row['Item Left To Give'] && row['Item Left To Give'] !== 'Nothing Pending') {
      pendingRows.push([
        row['Student Name'],
        row['School Name'],
        row['Item Name'],
        row['Item Left To Give'],
        row['Sale Date'],
        row['Phone No.'],
        row['Notes']
      ]);
    }
  });
  replaceData(SHEET_NAMES.credit, HEADERS.credit, creditRows);
  replaceData(SHEET_NAMES.pending, HEADERS.pending, pendingRows);
  replaceData(SHEET_NAMES.dashboard, HEADERS.dashboard, buildDashboardRows(stockRows, sales, creditRows, pendingRows));
  formatSheets(ss);
}

function buildStockSummaryRows() {
  const totals = {};
  readObjects(SHEET_NAMES.inventory, HEADERS.inventory).forEach(function(row) {
    const key = stockKey(row['School Name'], row['Item Name'], row.Size);
    if (!totals[key]) totals[key] = { school: row['School Name'], item: row['Item Name'], size: row.Size, qty: 0 };
    totals[key].qty += number(row['Quantity Received']);
  });
  readObjects(SHEET_NAMES.sales, HEADERS.sales).forEach(function(row) {
    const key = stockKey(row['School Name'], row['Item Name'], row.Size);
    if (!totals[key]) totals[key] = { school: row['School Name'], item: row['Item Name'], size: row.Size, qty: 0 };
    totals[key].qty -= number(row['Quantity Sold']);
  });
  return Object.keys(totals).sort().map(function(key) {
    const item = totals[key];
    const available = Math.max(item.qty, 0);
    return [item.school, item.item, item.size, available, stockStatus(available), new Date()];
  });
}

function getAvailableStock(school, item, size) {
  const rows = buildStockSummaryRows();
  const match = rows.find(function(row) {
    return row[0] === school && row[1] === item && String(row[2]) === String(size);
  });
  return match ? number(match[3]) : 0;
}

function buildDashboardRows(stockRows, sales, creditRows, pendingRows) {
  const totalInventory = stockRows.reduce(function(total, row) { return total + number(row[3]); }, 0);
  const totalSales = sales.reduce(function(total, row) { return total + number(row['Total Bill']); }, 0);
  const totalCredit = creditRows.reduce(function(total, row) { return total + number(row[4]); }, 0);
  const lowStock = stockRows.filter(function(row) { return number(row[3]) > 0 && number(row[3]) <= 5; }).length;
  const outOfStock = stockRows.filter(function(row) { return number(row[3]) === 0; }).length;
  const now = new Date();
  return [
    ['Total Inventory', totalInventory, now],
    ['Total Sales', totalSales, now],
    ['Credit Amount Outstanding', totalCredit, now],
    ['Pending Items To Give', pendingRows.length, now],
    ['Low Stock Items', lowStock, now],
    ['Out Of Stock Items', outOfStock, now]
  ];
}

function notifyLowStockForItem(school, item, size) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockRows = readObjects(SHEET_NAMES.stock, HEADERS.stock);
  const match = stockRows.find(function(row) {
    return row['School Name'] === school && row['Item Name'] === item && String(row.Size) === String(size);
  });
  if (!match) return;
  const qty = number(match['Available Quantity']);
  if (qty <= 5) {
    ss.getSheetByName(SHEET_NAMES.notifications).appendRow([
      new Date(),
      school,
      item,
      size,
      qty,
      stockStatus(qty)
    ]);
  }
}

function getAllData() {
  rebuildDerivedSheets();
  return {
    ok: true,
    stock: mapStock(readObjects(SHEET_NAMES.stock, HEADERS.stock)),
    sales: mapSales(readObjects(SHEET_NAMES.sales, HEADERS.sales)),
    credits: mapCredits(readObjects(SHEET_NAMES.credit, HEADERS.credit)),
    pending: mapPending(readObjects(SHEET_NAMES.pending, HEADERS.pending)),
    notifications: readObjects(SHEET_NAMES.notifications, HEADERS.notifications),
    dashboard: readObjects(SHEET_NAMES.dashboard, HEADERS.dashboard)
  };
}

function mapStock(rows) {
  return rows.map(function(row) {
    return {
      schoolName: row['School Name'],
      itemName: row['Item Name'],
      size: row.Size,
      availableQuantity: number(row['Available Quantity']),
      status: row.Status,
      date: toIsoDate(row['Last Updated'])
    };
  });
}

function mapSales(rows) {
  return rows.map(function(row) {
    return {
      saleDate: toIsoDate(row['Sale Date']),
      studentName: row['Student Name'],
      schoolName: row['School Name'],
      itemName: row['Item Name'],
      size: row.Size,
      quantitySold: number(row['Quantity Sold']),
      sellingPrice: number(row['Selling Price']),
      additionalCost: number(row['Additional Cost']),
      totalBill: number(row['Total Bill']),
      amountPaid: number(row['Amount Paid']),
      remainingAmount: number(row['Remaining Amount']),
      creditSale: row['Credit Sale'],
      phoneNo: row['Phone No.'],
      itemLeftToGive: row['Item Left To Give'],
      notes: row.Notes
    };
  });
}

function mapCredits(rows) {
  return rows.map(function(row) {
    return {
      studentName: row['Student Name'],
      schoolName: row['School Name'],
      totalBill: number(row['Total Bill']),
      amountPaid: number(row['Amount Paid']),
      remainingAmount: number(row['Remaining Amount']),
      saleDate: toIsoDate(row['Sale Date']),
      phoneNo: row['Phone No.'],
      notes: row.Notes
    };
  });
}

function mapPending(rows) {
  return rows.map(function(row) {
    return {
      studentName: row['Student Name'],
      schoolName: row['School Name'],
      itemSold: row['Item Sold'],
      pendingItem: row['Pending Item'],
      saleDate: toIsoDate(row['Sale Date']),
      phoneNo: row['Phone No.'],
      notes: row.Notes
    };
  });
}

function readObjects(sheetName, headers) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.filter(function(row) {
    return row.some(function(cell) { return cell !== ''; });
  }).map(function(row) {
    const object = {};
    headers.forEach(function(header, index) {
      object[header] = row[index];
    });
    return object;
  });
}

function replaceData(sheetName, headers, rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function formatSheets(ss) {
  Object.keys(SHEET_NAMES).forEach(function(key) {
    const sheet = ss.getSheetByName(SHEET_NAMES[key]);
    const headerCount = HEADERS[key].length;
    sheet.getRange(1, 1, 1, headerCount).setFontWeight('bold').setBackground('#eaf2f0');
    sheet.autoResizeColumns(1, headerCount);
  });
}

function validateRequired(object, fields) {
  fields.forEach(function(field) {
    if (object[field] === undefined || object[field] === null || object[field] === '') {
      throw new Error('Missing required field: ' + field);
    }
  });
}

function cleanDate(value) {
  if (!value) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
    const parts = String(value).split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) return value;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toIsoDate(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const date = new Date(value);
  if (!isNaN(date.getTime())) return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value);
}

function stockKey(school, item, size) {
  return [text(school).toLowerCase(), text(item).toLowerCase(), text(size).toLowerCase()].join('|');
}

function stockStatus(qty) {
  if (number(qty) === 0) return 'OUT OF STOCK';
  if (number(qty) <= 5) return 'LOW STOCK';
  return 'In Stock';
}

function number(value) {
  const parsed = Number(value);
  return isNaN(parsed) ? 0 : parsed;
}

function text(value) {
  return String(value || '').trim();
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function scriptResponse(payload, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(payload);
}
