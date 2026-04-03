/**
 * Google Apps Script — Pricing Simulation Data Receiver
 * 
 * SETUP:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click Deploy → New deployment → Web app
 *    - Execute as: Me | Who has access: Anyone
 * 5. Copy the Web App URL into the simulation app
 */

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Teams sheet ──────────────────────────────────────────
    var teamsSheet = ss.getSheetByName("Teams");
    if (!teamsSheet) {
      teamsSheet = ss.insertSheet("Teams");
      teamsSheet.appendRow([
        "Timestamp","Team Name","Section",
        "Member 1","Member 2","Member 3","Member 4",
        "Generic Strategy","Year 1 Objectives","Year 2 Objectives","Year 3 Objectives",
        "Total Revenue","Total Profit","Total Meals Sold",
        "Revenue Comments","Profit Comments"
      ]);
      teamsSheet.getRange(1, 1, 1, 16).setFontWeight("bold");
    }

    var members = data.members || [];
    var strategy = data.strategy || {};
    var quarters = data.quarters || [];
    var conclusions = data.conclusions || {};

    var totalRevenue = 0, totalProfit = 0, totalSales = 0;
    for (var i = 0; i < quarters.length; i++) {
      totalRevenue += (quarters[i].revenue || 0);
      totalProfit += (quarters[i].profit || 0);
      totalSales += (quarters[i].totalSales || 0);
    }

    teamsSheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.teamName || "", data.section || "",
      members[0] ? members[0].name : "", members[1] ? members[1].name : "",
      members[2] ? members[2].name : "", members[3] ? members[3].name : "",
      strategy.generic || "", strategy.year1 || "",
      strategy.year2 || "", strategy.year3 || "",
      totalRevenue, totalProfit, totalSales,
      conclusions.revenueComments || "", conclusions.profitComments || ""
    ]);

    // ── Quarters sheet ───────────────────────────────────────
    var qSheet = ss.getSheetByName("Quarters");
    if (!qSheet) {
      qSheet = ss.insertSheet("Quarters");
      qSheet.appendRow([
        "Timestamp","Team Name","Section","Quarter","Phase",
        "Own Price","Avg Industry Price","Promotion Expense",
        "New Customers","Retained Customers","Total Sales",
        "Revenue","Profit","Observations","Next Quarter Strategy"
      ]);
      qSheet.getRange(1, 1, 1, 15).setFontWeight("bold");
    }

    for (var j = 0; j < quarters.length; j++) {
      var q = quarters[j];
      qSheet.appendRow([
        data.timestamp || new Date().toISOString(),
        data.teamName || "", data.section || "",
        q.quarter, q.phase, q.ownPrice, q.avgCompPrice,
        q.promoExpense || 0, q.newCustomers || 0, q.retainedCustomers || 0,
        q.totalSales || 0, q.revenue || 0, q.profit || 0,
        q.observations || "", q.nextStrategy || ""
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status:"ok",message:"Data saved!"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:"error",message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status:"ok",message:"API running"}))
    .setMimeType(ContentService.MimeType.JSON);
}
