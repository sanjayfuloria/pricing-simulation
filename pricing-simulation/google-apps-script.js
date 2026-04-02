/**
 * Google Apps Script — Pricing Simulation Data Receiver
 * 
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click Deploy → New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy and copy the Web App URL
 * 6. Paste the URL in the simulation app's Google Sheets field
 * 
 * The script creates two sheets:
 *   "Teams"    — one row per submission (team info + totals)
 *   "Quarters" — one row per quarter per submission (detailed data)
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Teams sheet ──────────────────────────────────────────────
    let teamsSheet = ss.getSheetByName("Teams");
    if (!teamsSheet) {
      teamsSheet = ss.insertSheet("Teams");
      teamsSheet.appendRow([
        "Timestamp",
        "Team Name",
        "Section",
        "Member 1", "Member 2", "Member 3", "Member 4",
        "Generic Strategy",
        "Year 1 Objectives",
        "Year 2 Objectives",
        "Year 3 Objectives",
        "Total Revenue (All Quarters)",
        "Total Profit (All Quarters)",
        "Total Meals Sold",
        "Revenue Comments",
        "Profit Comments"
      ]);
      teamsSheet.getRange(1, 1, 1, 16).setFontWeight("bold");
    }

    const members = data.members || [];
    const memberNames = members.map(m => m.name || "").join(", ");
    const strategy = data.strategy || {};
    const quarters = data.quarters || [];
    const conclusions = data.conclusions || {};

    const totalRevenue = quarters.reduce((s, q) => s + (q.revenue || 0), 0);
    const totalProfit = quarters.reduce((s, q) => s + (q.profit || 0), 0);
    const totalSales = quarters.reduce((s, q) => s + (q.totalSales || 0), 0);

    teamsSheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.teamName || "",
      data.section || "",
      members[0] ? members[0].name : "",
      members[1] ? members[1].name : "",
      members[2] ? members[2].name : "",
      members[3] ? members[3].name : "",
      strategy.generic || "",
      strategy.year1 || "",
      strategy.year2 || "",
      strategy.year3 || "",
      totalRevenue,
      totalProfit,
      totalSales,
      conclusions.revenueComments || "",
      conclusions.profitComments || ""
    ]);

    // ── Quarters sheet ───────────────────────────────────────────
    let qSheet = ss.getSheetByName("Quarters");
    if (!qSheet) {
      qSheet = ss.insertSheet("Quarters");
      qSheet.appendRow([
        "Timestamp",
        "Team Name",
        "Section",
        "Quarter",
        "Phase",
        "Own Price",
        "Avg Industry Price",
        "Promotion Expense",
        "New Customers",
        "Retained Customers",
        "Total Sales",
        "Revenue",
        "Profit",
        "Observations",
        "Next Quarter Strategy"
      ]);
      qSheet.getRange(1, 1, 1, 15).setFontWeight("bold");
    }

    quarters.forEach(function(q) {
      qSheet.appendRow([
        data.timestamp || new Date().toISOString(),
        data.teamName || "",
        data.section || "",
        q.quarter,
        q.phase,
        q.ownPrice,
        q.avgCompPrice,
        q.promoExpense || 0,
        q.newCustomers || 0,
        q.retainedCustomers || 0,
        q.totalSales || 0,
        q.revenue || 0,
        q.profit || 0,
        q.observations || "",
        q.nextStrategy || ""
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", message: "Data saved successfully!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Allow CORS preflight
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Pricing Simulation API is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}
