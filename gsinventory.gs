/**
 * FASHION FIZZ BD - INVENTORY MANAGEMENT (FIXED)
 * Supports Variants (Sizes), Opening Stock, and Cloud Images
 */

function itemShowInventoryUI() {
  const html = HtmlService.createTemplateFromFile('inventory')
    .evaluate()
    .setTitle('Inventory Management');
  SpreadsheetApp.getUi().showSidebar(html);
}

function itemGetInventoryData() {
  return {
    items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS'),
    brands: _getUniqueDimension('Brands'), 
    categories: _getUniqueDimension('Item Category'),
    subcategories: _getUniqueDimension('Item Subcategory')
  };
}

function itemGenerateInventoryId() {
  return 'P' + Math.floor(10000 + Math.random() * 90000);
}

function itemDeleteItem(itemId) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEINVENTORYITEMS');
  const sheet = range.getSheet();
  const data = range.getValues();
  const idCol = data[0].indexOf('Item ID');
  const rowIdx = data.findIndex(r => r[idCol] === itemId);
  
  if (rowIdx > 0) { 
    sheet.deleteRow(range.getRow() + rowIdx);
    return { success: true, message: "Item deleted successfully" };
  } else {
    throw new Error("Item ID not found");
  }
}

/**
 * MASTER SAVE FUNCTION - UPDATED
 * Now saves Size + Quantity (e.g. "S:5, M:0") so we can view it later.
 */
function itemSaveProductWithVariants(data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems'); 
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let itemId = data.id === "AUTO" || !data.id ? itemGenerateInventoryId() : data.id;
  let isNew = data.id === "AUTO" || !data.id;

  // --- CHANGED HERE: Combine Size AND Stock (e.g., "S:10, L:5") ---
  const sizes = data.variants.map(v => `${v.size}:${v.openingStock || 0}`).join(', ');
  
  const totalOpeningStock = data.variants.reduce((sum, v) => sum + (parseFloat(v.openingStock) || 0), 0);

  const rowMap = {
    "Item ID": itemId,
    "Item Name": data.name,
    "Brands": data.brand,
    "Item Category": data.category,
    "Item Subcategory": data.subcategory,
    "Size": sizes, // Now stores "S:10, L:0" 
    "Reorder Level": data.reorderLevel,
    "Image URL": data.imageUrl, 
    "QTY Purchased": totalOpeningStock,
    "Remaining QTY": totalOpeningStock,
    "QTY Sold": 0,
    "Reorder Required": "No"
  };

  if (isNew) {
    const newRow = headers.map(h => {
      const val = rowMap[h.trim()];
      return val !== undefined ? val : "";
    });
    sheet.appendRow(newRow);
  } else {
    const allData = sheet.getDataRange().getValues();
    const idColIdx = headers.indexOf("Item ID");
    const rowIdx = allData.findIndex(r => r[idColIdx] === itemId);
    
    if (rowIdx > -1) {
      const sheetRowNum = rowIdx + 1;
      headers.forEach((h, colIdx) => {
        const trimmedH = h.trim();
        // Update all fields except dynamic stock columns logic if needed
        if (rowMap[trimmedH] !== undefined && !["QTY Purchased", "Remaining QTY", "QTY Sold"].includes(trimmedH)) {
          sheet.getRange(sheetRowNum, colIdx + 1).setValue(rowMap[trimmedH]);
        }
      });
    }
  }

  return { success: true, message: "Product saved successfully under ID: " + itemId };
}

// ... (Rest of the file: itemAddNewDimension, itemUploadImage, etc. remain unchanged) ...
function itemAddNewDimension(type, value) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEDIMENSIONS');
  const sheet = range.getSheet();
  const headers = range.getValues()[0];
  const colIdx = headers.indexOf(type);
  if (colIdx === -1) throw new Error("Dimension type not found: " + type);
  const lastRow = sheet.getLastRow();
  const colData = sheet.getRange(1, colIdx + 1, lastRow).getValues();
  let emptyRow = lastRow + 1;
  for(let i = 1; i < colData.length; i++) {
    if(colData[i][0] === "" || colData[i][0] === null) {
      emptyRow = i + 1;
      break;
    }
  }
  sheet.getRange(emptyRow, colIdx + 1).setValue(value);
  SpreadsheetApp.flush();
  return { success: true };
}

function itemUploadImage(base64Data, fileName) {
  try {
    const folderId = '1usgkVjV4Q7oLQ7leBQQk2FABoPxDeed5';
    const folder = DriveApp.getFolderById(folderId);
    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?sz=s1000&id=" + file.getId();
  } catch (e) {
    return "Error: " + e.toString();
  }
}