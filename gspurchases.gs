/**
 * FASHION FIZZ BD - PURCHASES (WITH DIRECT SUPPLIER CREATION)
 */

function poShowUI() {
  const html = HtmlService.createTemplateFromFile('purchases')
    .evaluate().setTitle('Purchase Orders').setWidth(1250).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getPurchaseStartupData() {
  return {
    suppliers: soGetRangeDataAsObjects('RANGESUPPLIERS') || [],
    items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS') || [],
    pos: soGetRangeDataAsObjects('RANGEPO') || []
  };
}

/**
 * MASTER SAVE: Handles New Supplier Creation + PO Saving
 */
function soSaveOrUpdatePO(poData, items, supplier) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName('PurchaseOrders');
  const pdSheet = ss.getSheetByName('PurchaseDetails');
  
  // 1. Handle Supplier (Create if New)
  let finalSupId = poData.supId;
  let finalSupName = poData.supName;

  if (supplier && supplier.isNew) {
    finalSupId = _poAddNewSupplier(supplier);
    finalSupName = supplier.name;
  }

  // 2. Save PO Header
  // Columns: [Date, PO ID, Supplier ID, Supplier Name, Bill Num, State, City, Total Amount, Total Paid, PO Balance, PMT Status, Status]
  poSheet.appendRow([
    new Date(), 
    poData.id, 
    finalSupId, 
    finalSupName, 
    poData.billNum, 
    supplier.state || "", 
    supplier.city || "", 
    poData.total, 
    0, // Paid
    poData.total, // Balance
    "Unpaid", 
    "Pending"
  ]);

  // 3. Save Details & Sync Stock
  items.forEach(item => {
    pdSheet.appendRow([
      new Date(), 
      poData.id, 
      "D-" + Date.now() + Math.floor(Math.random()*100), 
      finalSupId, 
      finalSupName,
      supplier.state || "", 
      supplier.city || "", 
      poData.billNum, 
      item.id, 
      "", // Type
      item.category, 
      "", // Subcat
      item.name,
      item.qty, 
      item.cost, 
      item.total, 
      0, 0, item.cost, 0, item.total
    ]);
    
    // Update Inventory Stock
    _syncPurchaseStock(item.id, item.size, item.qty);
  });

  // 4. Update Supplier Financials
  supUpdateFinancials(finalSupId, poData.total, 0);

  return { success: true, message: "PO Saved & Supplier Created!" };
}

/**
 * HELPER: Creates a new supplier row
 */
function _poAddNewSupplier(s) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Suppliers');
  
  // Generate ID
  const newId = "S" + Math.floor(10000 + Math.random() * 90000);
  
  sheet.appendRow([
    newId, 
    s.name, 
    s.contact, 
    s.email || "", 
    s.state || "", 
    s.city || "", 
    s.address || "", 
    0, 0, 0 // Financials
  ]);
  
  return newId;
}

/**
 * STOCK SYNC
 */
function _syncPurchaseStock(itemId, sizeName, qtyPurchased) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('InventoryItems');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const rowIdx = data.findIndex(r => r[headers.indexOf('Item ID')] === itemId);
  if (rowIdx === -1) return;

  const row = rowIdx + 1;
  const sizeCol = headers.indexOf('Size') + 1;
  const purCol = headers.indexOf('QTY Purchased') + 1;
  const remCol = headers.indexOf('Remaining QTY') + 1;

  const currentStr = data[rowIdx][headers.indexOf('Size')] || "";
  let sizeMap = {};
  
  // Parse "S:10, M:5"
  if(currentStr) {
    currentStr.split(',').forEach(p => {
      let [k, v] = p.split(':');
      if(k) sizeMap[k.trim()] = Number(v || 0);
    });
  }

  // Add Stock
  if(sizeName) {
    sizeMap[sizeName] = (sizeMap[sizeName] || 0) + Number(qtyPurchased);
  }

  // Rebuild String
  const newStr = Object.entries(sizeMap).map(([k,v]) => `${k}:${v}`).join(', ');
  
  sheet.getRange(row, sizeCol).setValue(newStr);
  
  // Update Totals
  const curPur = Number(data[rowIdx][headers.indexOf('QTY Purchased')] || 0);
  const curRem = Number(data[rowIdx][headers.indexOf('Remaining QTY')] || 0);
  
  sheet.getRange(row, purCol).setValue(curPur + Number(qtyPurchased));
  sheet.getRange(row, remCol).setValue(curRem + Number(qtyPurchased));
}

function poCreateInventoryItem(item) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems');
  const newId = "P" + Math.floor(Math.random() * 100000);
  sheet.appendRow([newId, "", item.category, "", "", item.brand, item.name, 0, 0, 0, 5, "No", ""]);
  return { id: newId, name: item.name };
}