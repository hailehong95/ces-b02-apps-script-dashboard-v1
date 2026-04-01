// ============================================================
// Code.gs — Server-side logic for Sales Dashboard
// Spreadsheet columns:
//   A: Transaction_ID | B: Date | C: Region | D: Product_Category
//   E: Product_SKU    | F: Quantity | G: Unit_Price | H: Discount_Applied
//   I: Customer_Rating | J: Is_Member | K: Total_Sales
// ============================================================

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sales Analytics Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Utility: include partial HTML files ──────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Load & normalise raw data ─────────────────────────────────
function getRawData_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var header = data[0];
  var rows   = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];

    // Normalise date → epoch ms (handles Date objects & strings)
    var rawDate = r[1];
    var dateMs  = null;
    if (rawDate instanceof Date) {
      dateMs = rawDate.getTime();
    } else if (typeof rawDate === 'string' && rawDate.length > 0) {
      var parsed = new Date(rawDate);
      dateMs = isNaN(parsed.getTime()) ? null : parsed.getTime();
    }

    rows.push({
      id          : parseInt(r[0])          || 0,
      dateMs      : dateMs,
      region      : String(r[2] || '').trim(),
      category    : String(r[3] || '').trim(),
      sku         : String(r[4] || '').trim(),
      quantity    : parseFloat(r[5])        || 0,
      unitPrice   : parseFloat(r[6])        || 0,
      discount    : parseFloat(r[7])        || 0,  // 0.0 – 0.2
      rating      : parseInt(r[8])          || 0,
      isMember    : String(r[9] || '').trim().toLowerCase() === 'yes',
      totalSales  : parseFloat(r[10])       || 0
    });
  }
  return rows;
}

// ============================================================
//  API FUNCTIONS (called via google.script.run from client)
// ============================================================

// 1. KPI Summary Cards
function getKpiData() {
  var rows = getRawData_();
  var totalRevenue = 0, totalQty = 0, totalRating = 0;
  rows.forEach(function(r) {
    totalRevenue += r.totalSales;
    totalQty     += r.quantity;
    totalRating  += r.rating;
  });
  return {
    totalTransactions : rows.length,
    totalRevenue      : Math.round(totalRevenue * 100) / 100,
    totalQuantity     : totalQty,
    avgRating         : Math.round((totalRating / rows.length) * 100) / 100
  };
}

// 2. Monthly Revenue Trend (LineChart)
function getMonthlySales() {
  var rows = getRawData_();
  var MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun',
                     'Jul','Aug','Sep','Oct','Nov','Dec'];
  var map = {};
  for (var m = 0; m < 12; m++) map[m] = 0;

  rows.forEach(function(r) {
    if (r.dateMs === null) return;
    var month = new Date(r.dateMs).getMonth(); // 0-based
    map[month] += r.totalSales;
  });

  // Google Charts DataTable format: [[label, value], ...]
  var result = [['Month', 'Revenue (USD)']];
  for (var m = 0; m < 12; m++) {
    result.push([MONTH_NAMES[m], Math.round(map[m] * 100) / 100]);
  }
  return result;
}

// 3. Sales by Product Category (PieChart)
function getSalesByCategory() {
  var rows = getRawData_();
  var map = {};
  rows.forEach(function(r) {
    map[r.category] = (map[r.category] || 0) + r.totalSales;
  });
  var result = [['Category', 'Revenue (USD)']];
  Object.keys(map).sort().forEach(function(k) {
    result.push([k, Math.round(map[k] * 100) / 100]);
  });
  return result;
}

// 4. Sales by Region (BarChart)
function getSalesByRegion() {
  var rows = getRawData_();
  var map = {};
  rows.forEach(function(r) {
    map[r.region] = (map[r.region] || 0) + r.totalSales;
  });
  var result = [['Region', 'Revenue (USD)']];
  // Sort descending by revenue
  var sorted = Object.keys(map).sort(function(a,b){ return map[b] - map[a]; });
  sorted.forEach(function(k) {
    result.push([k, Math.round(map[k] * 100) / 100]);
  });
  return result;
}

// 5. Customer Rating Distribution (ColumnChart)
function getRatingDistribution() {
  var rows = getRawData_();
  var map = {1:0, 2:0, 3:0, 4:0, 5:0};
  rows.forEach(function(r) {
    if (r.rating >= 1 && r.rating <= 5) map[r.rating]++;
  });
  var result = [['Rating', 'Number of Orders', {role:'style'}]];
  var colors = ['#e74c3c','#e67e22','#f1c40f','#2ecc71','#27ae60'];
  for (var i = 1; i <= 5; i++) {
    result.push(['★'.repeat(i), map[i], colors[i-1]]);
  }
  return result;
}

// 6. Discount Level vs Avg Revenue (ColumnChart)
function getDiscountImpact() {
  var rows = getRawData_();
  var map = {};
  rows.forEach(function(r) {
    var key = (r.discount * 100).toFixed(0) + '%';
    if (!map[key]) map[key] = {sum:0, count:0};
    map[key].sum   += r.totalSales;
    map[key].count += 1;
  });
  var result = [['Discount', 'Avg Revenue (USD)']];
  ['0%','5%','10%','15%','20%'].forEach(function(k) {
    var d = map[k] || {sum:0, count:1};
    result.push([k, Math.round((d.sum / d.count) * 100) / 100]);
  });
  return result;
}

// 7. Member vs Non-Member per Category (StackedBarChart)
function getMemberBreakdown() {
  var rows = getRawData_();
  var categories = ['Clothing','Electronics','Garden','Toys'];
  // map[category] = {member: 0, nonMember: 0}
  var map = {};
  categories.forEach(function(c){ map[c] = {member:0, nonMember:0}; });
  rows.forEach(function(r) {
    if (!map[r.category]) return;
    if (r.isMember) map[r.category].member++;
    else            map[r.category].nonMember++;
  });
  var result = [['Category', 'Member', 'Non-Member']];
  categories.forEach(function(c) {
    result.push([c, map[c].member, map[c].nonMember]);
  });
  return result;
}
