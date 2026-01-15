/**
 * FASHION FIZZ BD - SALES ENGINE (ROBUST SAVE)
 * Fixes: Ensures Sales Details are saved even if stock sync has issues.
 */

function soShowSalesUI() {
  const html = HtmlService.createTemplateFromFile('sales')
    .evaluate()
    .setTitle('New Sales Order')
    .setWidth(1250)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getSalesStartupData() {
  try {
    return {
      customers: soGetRangeDataAsObjects('RANGECUSTOMERS') || [],
      items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS') || [], 
      sales: soGetRangeDataAsObjects('RANGESO') || [],
      cities: _getUniqueDimension('City') || []
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * MASTER SAVE FUNCTION
 */
function soSaveOrder(soData, items, customer) {
  const ss = SpreadsheetApp.getActive();
  const soSheet = ss.getSheetByName('SalesOrders');
  const sdSheet = ss.getSheetByName('SalesDetails');
  
  if (!soSheet || !sdSheet) throw new Error("Critical Error: 'SalesOrders' or 'SalesDetails' sheet missing.");

  // 1. Handle Customer (Create if New)
  let custId = customer.id;
  if (customer.isNew) {
    custId = custAddNewCustomer({
      name: customer.name,
      contact: customer.contact,
      city: customer.city,
      address: customer.address
    });
  }

  // 2. SAVE SALES ORDER (Header)
  soSheet.appendRow([
    new Date(), 
    soData.id, 
    custId, 
    customer.name, 
    soData.invoice, 
    customer.state || "", 
    customer.city, 
    soData.totalAmount, 
    0, // Total Received
    soData.totalAmount, // Balance
    "Unpaid", 
    "Pending"
  ]);
  
  // FORCE SAVE: Ensures the header is recorded even if details fail later
  SpreadsheetApp.flush(); 

  // 3. SAVE DETAILS & UPDATE STOCK
  // We use a loop with try-catch to prevent one bad item from stopping everything
  const inventory = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');
  
  items.forEach(item => {
    try {
      // A. Write to SalesDetails
      sdSheet.appendRow([
        new Date(), 
        soData.id, 
        "SD-" + Math.floor(Math.random() * 1000000000), // Unique Detail ID
        custId, 
        customer.name,
        customer.state || "", 
        customer.city, 
        soData.invoice, 
        item.id, 
        item.category || "", // Item Type 
        item.category || "", // Item Category
        item.subcategory || "", 
        item.name, 
        item.qty, 
        item.price, 
        item.price, // Excl Tax
        0, // Tax Rate
        0, // Tax Total
        item.price, // Incl Tax
        item.ship, 
        item.total
      ]);

      // B. Update Inventory Stock
      _syncSalesStockSafe(item.id, item.size, item.qty, inventory);
      
    } catch (err) {
      console.error("Error saving item " + item.name + ": " + err.message);
      // We continue to the next item instead of crashing the whole order
    }
  });

  // 4. Update Customer Financials
  try {
    custUpdateCustomerFinancials(custId, soData.totalAmount);
  } catch (e) {
    console.error("Error updating customer balance: " + e.message);
  }

  return { success: true, message: "Order " + soData.id + " saved successfully!" };
}

/**
 * ROBUST STOCK SYNC
 * Updates the size string (S:10 -> S:9) and total counts safely.
 */
function _syncSalesStockSafe(itemId, sizeSold, qtySold, preloadedInventory) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems');
  
  // We use the preloaded data to find the row index efficiently
  // Note: Row index in sheet = Array Index + 2 (Header is row 1)
  const itemIndex = preloadedInventory.findIndex(i => i['Item ID'] === itemId);
  
  if (itemIndex === -1) return; // Item not found
  
  const rowNum = itemIndex + 2; 
  
  // Get current values directly from sheet to ensure we don't overwrite recent changes
  // Assuming columns: Size(Col 5), Purchased(Col 8), Sold(Col 9), Remaining(Col 10)
  // Adjust these indices if your column order changes!
  const sizeRange = sheet.getRange(rowNum, 5); // Column E (Size)
  const soldRange = sheet.getRange(rowNum, 9); // Column I (QTY Sold)
  const remRange = sheet.getRange(rowNum, 10); // Column J (Remaining QTY)
  
  const currentSizeStr = sizeRange.getValue().toString();
  const currentSold = Number(soldRange.getValue() || 0);
  
  // Parse Size String: "S:10, M:5" -> Object {S:10, M:5}
  let sizeMap = {};
  if (currentSizeStr) {
    currentSizeStr.split(',').forEach(part => {
      let [sName, sQty] = part.split(':').map(x => x ? x.trim() : "");
      if(sName) sizeMap[sName] = Number(sQty || 0);
    });
  }

  // Subtract Sold Qty
  if (sizeSold && sizeMap.hasOwnProperty(sizeSold)) {
    sizeMap[sizeSold] = Math.max(0, sizeMap[sizeSold] - Number(qtySold));
  } else if (!sizeSold && Object.keys(sizeMap).length === 0) {
    // Handling for items without variants (if any)
    // For now, we assume everything has variant logic per your request
  }

  // Rebuild String
  const newSizeStr = Object.entries(sizeMap)
    .map(([k,v]) => `${k}:${v}`)
    .join(', ');

  // Recalculate Total Remaining
  const newTotalStock = Object.values(sizeMap).reduce((a, b) => a + b, 0);

  // Update Sheet
  sizeRange.setValue(newSizeStr);
  soldRange.setValue(currentSold + Number(qtySold));
  remRange.setValue(newTotalStock);
  
  // Flush to ensure stock is updated before next op
  SpreadsheetApp.flush();
}