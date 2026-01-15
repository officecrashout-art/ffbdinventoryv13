/**
 * FASHION FIZZ BD - PAYMENTS (INTEGRATED)
 */

function ptShowUI() {
  const html = HtmlService.createTemplateFromFile('payments')
    .evaluate().setTitle('Vendor Payments').setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getPaymentStartupData() {
  return {
    suppliers: soGetRangeDataAsObjects('RANGESUPPLIERS'),
    pos: soGetRangeDataAsObjects('RANGEPO'),
    payments: soGetRangeDataAsObjects('RANGEPAYMENTS')
  };
}

function ptSaveNewPayment(pay) {
  const ss = SpreadsheetApp.getActive();
  const ptSheet = ss.getSheetByName('Payments');
  const poSheet = ss.getSheetByName('PurchaseOrders');
  
  // 1. Log Payment
  ptSheet.appendRow([
    new Date(pay.date), pay.id, pay.supId, pay.supName,
    "", "", pay.poId, pay.bill, pay.mode, pay.amount
  ]);
  
  // 2. Update PO Balance & Status
  const poData = poSheet.getDataRange().getValues();
  const h = poData[0];
  const rIdx = poData.findIndex(r => r[h.indexOf('PO ID')] === pay.poId);
  
  if(rIdx > 0) {
    const r = rIdx + 1;
    const paidCol = h.indexOf('Total Paid') + 1;
    const balCol = h.indexOf('PO Balance') + 1;
    const statCol = h.indexOf('PMT Status') + 1;
    
    const curPaid = Number(poData[rIdx][h.indexOf('Total Paid')] || 0);
    const total = Number(poData[rIdx][h.indexOf('Total Amount')] || 0);
    
    const newPaid = curPaid + Number(pay.amount);
    const newBal = total - newPaid;
    
    poSheet.getRange(r, paidCol).setValue(newPaid);
    poSheet.getRange(r, balCol).setValue(newBal > 0 ? newBal : 0);
    poSheet.getRange(r, statCol).setValue(newBal <= 0 ? 'Paid' : 'Partial');
  }

  // 3. Update Supplier Financials
  supUpdateFinancials(pay.supId, 0, pay.amount);

  return { success: true };
}