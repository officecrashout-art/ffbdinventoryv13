/**
 * FASHION FIZZ BD - PREMIUM BACKEND
 */

function getDashboardData() {
  const ss = SpreadsheetApp.getActive();
  const soData = soGetRangeDataAsObjects('RANGESO');
  const invData = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');

  // 1. KPI Calculations
  const totalSales = soData.reduce((sum, r) => sum + (Number(r['Total SO Amount']) || 0), 0);
  const totalReceived = soData.reduce((sum, r) => sum + (Number(r['Total Received']) || 0), 0);
  const totalDue = soData.reduce((sum, r) => sum + (Number(r['SO Balance']) || 0), 0);

  // 2. Inventory Logic (Valuation & Low Stock)
  let stockValue = 0;
  const categoryMap = {};
  const lowStockItems = [];

  invData.forEach(item => {
    const qty = Number(item['Remaining QTY'] || 0);
    // Safe Cost Access: Check 'Unit Cost', 'Purchase Price', or default to 0
    const cost = Number(item['Unit Cost'] || item['Purchase Price'] || item['Cost Price'] || 0);
    
    // Calculate Value
    stockValue += (qty * cost);

    // Pie Chart Data
    if (qty > 0) {
      const cat = item['Item Category'] || 'Uncategorized';
      categoryMap[cat] = (categoryMap[cat] || 0) + qty;
    }

    // Low Stock Alert (Handling 0 Stock)
    const reorder = Number(item['Reorder Level'] || 0);
    // Trigger if Qty is 0 OR (Reorder Level is set AND Qty <= Level)
    if (qty === 0 || (reorder > 0 && qty <= reorder)) {
      lowStockItems.push({ 
        name: item['Item Name'], 
        qty: qty, 
        level: reorder > 0 ? reorder : 'N/A',
        id: item['Item ID']
      });
    }
  });

  // 3. Top Modules (Cities)
  const cityMap = {};
  soData.forEach(r => {
    const amt = Number(r['Total SO Amount'] || 0);
    if(r['City']) cityMap[r['City']] = (cityMap[r['City']] || 0) + amt;
  });

  const getTop5 = (map, key) => Object.entries(map)
    .map(([k,v]) => ({ [key]: k, total: v }))
    .sort((a,b) => b.total - a.total).slice(0, 5);

  // 4. Sales Chart
  const salesByMonth = {};
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  soData.forEach(r => {
    const d = new Date(r['SO Date']);
    if (d instanceof Date && !isNaN(d)) {
      const key = months[d.getMonth()] + " " + d.getFullYear().toString().substr(-2);
      salesByMonth[key] = (salesByMonth[key] || 0) + Number(r['Total SO Amount']);
    }
  });

  return {
    kpi: { sales: totalSales, received: totalReceived, due: totalDue, stock: stockValue },
    chart: { labels: Object.keys(salesByMonth), values: Object.values(salesByMonth) },
    pie: { labels: Object.keys(categoryMap), values: Object.values(categoryMap) },
    lowStock: lowStockItems.slice(0, 10), // Limit to top 10 alerts
    topCities: getTop5(cityMap, 'city'),
    recent: soData.reverse().slice(0, 6).map(r => ({
      id: r['SO ID'], name: r['Customer Name'], amount: Number(r['Total SO Amount']), status: r['Receipt Status'] || 'Unpaid'
    }))
  };
}