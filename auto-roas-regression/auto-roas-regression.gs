/**
 * ==============================================================================
 * AUTOMATED ROAS OPTIMISATION SCRIPT (QUADRATIC REGRESSION) - V3.2
 * ==============================================================================
 *
 * WHAT DOES THIS SCRIPT DO?
 * -------------------------
 * This script optimises your campaigns for Maximum Profit (rather than just Revenue).
 * 1. Pulls historical data from Google's Traffic Simulator (All Campaign Types).
 * 2. Uses quadratic regression to model the relationship between ROAS and Profit.
 * 3. Identifies the specific ROAS target where profit is maximized.
 * 4. Logs the analysis to a spreadsheet and generates a visual Profit Curve.
 * 5. Updates the campaign or portfolio (supports Search, Shopping, PMax).
 *
 * ==============================================================================
 */

/**
 * ==============================================================================
 * CONFIGURATION (Settings to customise your script)
 * ==============================================================================
 */

const CONFIG = {
  // 1. WHICH CAMPAIGNS SHOULD WE LOOK AT?
  // Type the exact names of the campaigns you want the script to manage. 
  // Put each name in quote marks, inside the brackets, separated by commas.
  // For example: ["UK_Search_Brand", "PMax_Shoes"]. 
  // If you leave it empty like [], the script will look at every eligible campaign.
  campaignNames: [], 
  
  // 2. LIFETIME VALUE (LTV) MULTIPLIER
  // This setting adjusts the value of a sale to account for future repeat 
  // purchases. It tells the script what a customer is actually worth in 
  // the long run, rather than just their one-off checkout total.
  //
  // - 1.0 = DEFAULT. Use this if you only want to track the initial sale.
  // - 1.4 = LTV BOOST. For this client, every £100 spent at checkout is 
  //         worth £140 in total lifetime value. We use 1.4 to optimise 
  //         for this "True Value".
  //
  // Note: This is a multiplier, not a margin calculation. To increase 
  // the value by 40%, use 1.4. (Using 0.4 would reduce the value by 60%).
  conversionValueMultiplier: 1.0,

  // 3. CONFIDENCE SCORE (DATA QUALITY)
  // This is a trust score for Google's data, graded from 0 to 1.
  // A score of 1 means the data forms a perfect, reliable pattern.
  // A score of 0 means the data is completely random and unreliable.
  // The number you type here is the minimum score you are willing to accept.
  // If the data scores lower than your chosen number, the script will stop.
  // This stops the script from making changes based on bad guesses.
  minRSquared: 0.5,
  
  // 4. SAFETY GUARDRAILS
  // These settings act as safety nets to prevent the script from making massive or risky changes.
  guardrails: {
    // Limits how much the script can alter your ROAS target in a single step.
    // For instance, if you set this to 0.2, a current target of 2.0 can only move 
    // down to 1.8 or up to 2.2 in one go.
    maxRoasChange: 0.2, 
    // The absolute lowest ROAS target you are willing to accept for any campaign.
    minRoasLimit: 0.8,  
    // The absolute highest ROAS target you want the script to set for any campaign.
    maxRoasLimit: 4.0   
  },

  // 5. YOUR DASHBOARD (THE SPREADSHEET)
  // Create a new, empty Google Sheet and paste its full web address right here. 
  // Ensure the address is placed between the quote marks.
  // The script will automatically build your dashboard and store all its findings here.
  spreadsheetUrl: "", 
  
  // 6. WHO GETS THE EMAIL?
  // Enter your email address here between the quote marks.
  // The script will send you a full report every time it finishes a run.
  // To send reports to multiple people, separate their email addresses with commas.
  emailAddresses: "",

  // 7. THE MASTER SWITCH (ACTION MODE)
  // This is the main control button for making live changes.
  // Keep this set to 'false' at first. The script will run as a test, filling 
  // your spreadsheet but not touching your live Google Ads account.
  // Once you trust the numbers in the spreadsheet, change this word to 'true' 
  // to let the script automatically update your campaigns.
  updateCampaigns: false
};

/**
 * ==============================================================================
 * MAIN SCRIPT LOGIC
 * ==============================================================================
 */

/**
 * MAIN EXECUTION BLOCK
 * Handles global logging, spreadsheet initialisation, and campaign iteration.
 */
function main() {
  const logBuffer = [];
  const originalLogger = Logger.log;
  
  // Intercepting Logger to store messages for the final email report.
  Logger.log = function(msg) {
    originalLogger(msg);
    logBuffer.push(msg);
  };

  logHeader("🚀 IGNITING PROFIT ENGINE (OMNICHANNEL ENABLED)...");

  if (!CONFIG.updateCampaigns) {
    Logger.log("\n⚠️ NOTE: Script is in READ-ONLY mode. No changes will be pushed to Google Ads.\n");
  }

  const sheetObj = ensureSpreadsheet();
  if (!sheetObj) {
    Logger.log("💥 Critical Failure: Unable to spin up the spreadsheet.");
    sendFinalEmail(logBuffer.join("\n"), "CRITICAL FAILURE");
    return;
  }
  Logger.log(`📜 Mission Report: ${sheetObj.url}`);

  // Querying GAQL for simulation data. We require specifically TARGET_ROAS type simulations.
  const query = `
    SELECT 
      campaign.id, 
      campaign.name,
      campaign_simulation.target_roas_point_list.points 
    FROM campaign_simulation 
    WHERE 
      campaign_simulation.type = 'TARGET_ROAS'
      AND campaign.status = 'ENABLED'
  `;

  const report = AdsApp.search(query);
  let processedCount = 0;

  // Iterate through campaigns returned by the API that have valid simulator data.
  while (report.hasNext()) {
    const row = report.next();
    const currentName = row.campaign.name;

    // Filter by name if the user populated the campaignNames array in CONFIG.
    if (CONFIG.campaignNames.length > 0 && CONFIG.campaignNames.indexOf(currentName) === -1) {
      continue; 
    }

    // --- FAULT TOLERANCE ADDED HERE ---
    try {
      processSingleCampaign(row, sheetObj);
      processedCount++;
    } catch (e) {
      Logger.log(`⚠️ SKIPPING "${currentName}": ${e.message}`);
      skippedCount++;
    }
  }

  // Formatting and rebuilding the interactive dashboard.
  polishSpreadsheet(sheetObj);

  if (processedCount === 0) {
    Logger.log("\n🤷‍♂️ No eligible campaigns found. The simulator is silent.");
  } else {
    logHeader(`🏁 MISSION COMPLETE. Processed ${processedCount} Campaign(s).`);
  }

  sendFinalEmail(logBuffer.join("\n"), `Processed ${processedCount} Campaigns`);
}

/**
 * CORE LOGIC
 * Extracts data, performs mathematical analysis, applies constraints, and logs.
 */
function processSingleCampaign(row, sheetObj) {
  const campaignName = row.campaign.name;
  const campaignId = row.campaign.id;
  
  logHeader(`🕵️ ANALYSING: "${campaignName}"`);

  const currentSettings = getCurrentRoasSettings(campaignId);
  if (!currentSettings) {
    Logger.log(`🚫 Abort: Couldn't find current bidding settings.`);
    return;
  }
  
  const currentRoas = currentSettings.roas;
  Logger.log(`    📍 Current Target: ${currentRoas.toFixed(2)}`);

  const points = row.campaignSimulation.targetRoasPointList.points;
  let regressionData = [];
  let spendData = []; 

  Logger.log(`    🔮 Gazing into the Traffic Simulator (${points.length} scenarios found)...`);

  // Transform raw simulation points into workable datasets for regression.
  for (const point of points) {
    const roas = point.targetRoas;
    const cost = point.costMicros / 1000000;
    const value = point.biddableConversionsValue;
    // Calculate Predicted Net Profit using the multiplier in CONFIG.
    const profit = (value * CONFIG.conversionValueMultiplier) - cost;
    
    // Dataset for ROAS vs Profit (our main optimization goal).
    regressionData.push({ x: roas, y: profit });
    // Dataset for ROAS vs Spend (required for back-calculating metrics).
    spendData.push({ x: roas, y: cost });
  }

  // Require minimum 3 data points to fit a quadratic curve (parabola).
  if (regressionData.length < 3) {
    Logger.log("    👻 Ghost Town: Not enough data points to build a model.");
    return;
  }

  // --- REGRESSION ENGINE ---
  // Fit standard form: y = ax^2 + bx + c
  const coeffs = fitQuadratic(regressionData);
  const spendCoeffs = fitQuadratic(spendData);
  const rSquared = calculateRSquared(regressionData, coeffs);
  
  if (rSquared < CONFIG.minRSquared) {
    Logger.log(`    🎲 Too Chaotic: Data correlation is weak (R² ${rSquared.toFixed(2)}). Skipping for safety.`);
    return;
  }

  // Safety Check: A positive 'a' coefficient means the parabola opens upwards.
  // This implies infinite profit at high or low spend, indicating faulty data or logic. Aborting.
  if (coeffs.a >= 0) {
    Logger.log("    🎢 U-Curve Detected: Google thinks profit increases infinitely. Skipping.");
    return;
  }

  // --- 1. OPTIMAL MATHEMATICS ---
  // Finding the vertex of the parabola using x = -b / (2a)
  const optimalRoas = -coeffs.b / (2 * coeffs.a);
  let optimalProfit = (coeffs.a * (optimalRoas * optimalRoas)) + (coeffs.b * optimalRoas) + coeffs.c;
  
  let optimalSpend = (spendCoeffs.a * (optimalRoas * optimalRoas)) + (spendCoeffs.b * optimalRoas) + spendCoeffs.c;
  optimalSpend = Math.max(0, optimalSpend); 
  
  // Back-calculating revenue from profit and spend.
  let optimalRevenue = (optimalProfit + optimalSpend) / CONFIG.conversionValueMultiplier;
  optimalRevenue = Math.max(0, optimalRevenue);
  
  // Re-calculating profit to ensure consistency in spreadsheet logging.
  optimalProfit = (optimalRevenue * CONFIG.conversionValueMultiplier) - optimalSpend;

  // --- 2. GUARDED MATHEMATICS ---
  // Apply constraints (CONFIG.guardrails) to the theoretical mathematical optimum.
  const guardedRoas = applyGuardrails(optimalRoas, currentRoas);
  // Re-calculate all predicted performance metrics based on this constrained ROAS.
  let guardedProfit = (coeffs.a * (guardedRoas * guardedRoas)) + (coeffs.b * guardedRoas) + coeffs.c;
  
  let guardedSpend = (spendCoeffs.a * (guardedRoas * guardedRoas)) + (spendCoeffs.b * guardedRoas) + spendCoeffs.c;
  guardedSpend = Math.max(0, guardedSpend); 
  
  let guardedRevenue = (guardedProfit + guardedSpend) / CONFIG.conversionValueMultiplier;
  guardedRevenue = Math.max(0, guardedRevenue);
  
  guardedProfit = (guardedRevenue * CONFIG.conversionValueMultiplier) - guardedSpend;


  Logger.log(`    💎 Sweet Spot Found: ${optimalRoas.toFixed(2)}`);
  Logger.log(`    🛡️ Safety Shields:   ${guardedRoas.toFixed(2)} (Guarded)`);
  
  let actionTaken = "READ ONLY";

  // Prevent updates if the predicted maximum profit is negative.
  if (optimalProfit <= 0) {
    Logger.log(`    🛑 WARNING: Campaign intrinsically unprofitable (Peak Profit is £0 or less). Skipping to prevent bleeding.`);
    actionTaken = "SKIPPED (UNPROFITABLE)";
  } else if (CONFIG.updateCampaigns) {
    // Only flag as updated if the required change is greater than 0.01 (trivial difference).
    if (Math.abs(guardedRoas - currentRoas) > 0.01) {
      actionTaken = "UPDATED";
    } else {
      actionTaken = "NO CHANGE";
    }
  }

  const stats = getLastWeekStats(campaignId);

  // Send all data (theoretical and constrained) to spreadsheet tabs.
  logToSpreadsheet(sheetObj, {
    timestamp: new Date(),
    campaign: campaignName,
    rawPoints: points, 
    coeffs: coeffs,
    rSquared: rSquared,
    optimalRoas: optimalRoas,
    optimalProfit: optimalProfit,
    optimalSpend: optimalSpend,
    optimalRevenue: optimalRevenue,
    guardedRoas: guardedRoas,
    guardedProfit: guardedProfit,
    guardedSpend: guardedSpend,
    guardedRevenue: guardedRevenue,
    currentRoas: currentRoas, 
    stats: stats,
    action: actionTaken        
  });

  // Execute actual API updates via Mutate operations.
  if (actionTaken === "UPDATED") {
       applyTargetRoas(campaignId, currentSettings, guardedRoas);
  } else if (actionTaken === "NO CHANGE") {
       Logger.log(`    ✅ Status: We are already perfect. No change needed.`);
  } else if (actionTaken === "SKIPPED (UNPROFITABLE)") {
       Logger.log(`    👀 Status: Target update aborted due to negative maximum profit.`);
  } else {
       Logger.log(`    👀 Status: Skipped (Read-Only Mode)`);
  }
}

/**
 * EMAIL TOOLS
 * Sends caught logs at execution end.
 */
function sendFinalEmail(fullLog, statusSummary) {
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const recipient = CONFIG.emailAddresses;

  if (!recipient || recipient === "") return;

  const subject = `Google Ads Script Log: ${accountName} (${statusSummary})`;
  const body = `Automated ROAS Optimisation Script has finished running.\n\n` +
               `Account: ${accountName} (${accountId})\n` +
               `Spreadsheet: ${CONFIG.spreadsheetUrl}\n\n` +
               `--- EXECUTION LOGS ---\n\n${fullLog}`;

  try {
    MailApp.sendEmail(recipient, subject, body);
  } catch (e) {}
}

/**
 * SPREADSHEET & REPORTING TOOLS
 * Heavy manipulation of Spreadsheet API to build data logs and charts.
 */
function ensureSpreadsheet() {
  let ss;
  let isNew = false;

  if (CONFIG.spreadsheetUrl && CONFIG.spreadsheetUrl !== "") {
    try {
      ss = SpreadsheetApp.openByUrl(CONFIG.spreadsheetUrl);
    } catch (e) {
      Logger.log("⚠️ Config URL failed. Forging a new spreadsheet...");
      isNew = true;
    }
  } else {
    isNew = true;
  }

  if (isNew) {
    const accountName = AdsApp.currentAccount().getName();
    const accountId = AdsApp.currentAccount().getCustomerId();
    const fileName = `Auto ROAS Regression Logs | ${accountName} | ${accountId}`;
    
    ss = SpreadsheetApp.create(fileName);
    
    const emailStr = CONFIG.emailAddresses;
    if (emailStr && emailStr.length > 0) {
      const editors = emailStr.split(",").map(function(e) { return e.trim(); });
      try {
        ss.addEditors(editors);
        Logger.log(`    🤝 Granted access to: ${editors.join(", ")}`);
      } catch (e) {
        Logger.log(`    ⚠️ Failed to grant access: ${e.toString()}`);
      }
      MailApp.sendEmail(CONFIG.emailAddresses, "New ROAS Regression Log", 
        `Fresh hot spreadsheet created for ${accountName}.\n\nAccess it here: ${ss.getUrl()}`
      );
    } else {
        Logger.log("    ⚠️ No email provided. You must find the sheet URL in these logs.");
    }
  }

  let dashSheet = ss.getSheetByName("Dashboard");
  if (!dashSheet) dashSheet = ss.insertSheet("Dashboard", 0);

  const dataSheet = getOrCreateSheet(ss, "Historical Simulation Data", [
    "Timestamp", "Campaign", "Start Date", "End Date", 
    "Current Target ROAS", "Raw Points (Hidden)", 
    "Coeff A", "Coeff B", "Coeff C", "R-Squared", 
    "Optimal ROAS", "Profit (Optimal)"
  ]);
  
  const perfSheet = getOrCreateSheet(ss, "Weekly Performance Log", [
    "Timestamp", "Campaign", "Start Date", "End Date", 
    "Action Taken", 
    "Spend (Actual)", "Spend (Guarded)", "Spend (Optimal)",
    "Revenue (Actual)", "Revenue (Guarded)", "Revenue (Optimal)",
    "ROAS (Actual)", "ROAS (Guarded)", "ROAS (Optimal)", 
    "Profit (Actual)", "Profit (Guarded)", "Profit (Optimal)", "Profit Diff"
  ]);

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  return { ss: ss, url: ss.getUrl(), dataSheet: dataSheet, perfSheet: perfSheet, dashSheet: dashSheet };
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2); 
  }
  return sheet;
}

function polishSpreadsheet(sheetObj) {
  const ss = sheetObj.ss;
  const sheets = [sheetObj.dataSheet, sheetObj.perfSheet];
  const currencyCode = AdsApp.currentAccount().getCurrencyCode();
  const symbols = { 'USD': '$', 'GBP': '£', 'EUR': '€', 'AUD': '$' };
  const symbol = symbols[currencyCode] || '';
  const currencyFormat = symbol ? `${symbol}#,##0.00` : `#,##0.00 [${currencyCode}]`;

  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const maxCols = sheet.getMaxColumns();

    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground("#C33B48").setFontColor("white").setFontWeight("bold").setFontFamily("Manrope");
    
    sheet.setFrozenColumns(2);

    if (lastRow > 1) {
      const fullRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      fullRange.setFontFamily("Manrope");
      
      sheet.setColumnWidth(1, 150); 
      sheet.autoResizeColumn(2); 
      for (let i = 2; i < lastCol; i++) {
         sheet.setColumnWidth(i + 1, 140);
      }

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (let i = 0; i < headers.length; i++) {
        const headerName = headers[i].toLowerCase();
        const colRange = sheet.getRange(2, i + 1, lastRow - 1, 1);

        if (headerName.includes("roas") || headerName.includes("coeff") || headerName.includes("r-squared")) {
          colRange.setNumberFormat("0.00");
        } else if (headerName.includes("profit") || headerName.includes("spend") || headerName.includes("revenue") || headerName.includes("diff")) {
          colRange.setNumberFormat(currencyFormat);
        }
      }
      if (sheet.getFilter()) sheet.getFilter().remove();
      sheet.getDataRange().createFilter();
    }

    if (sheet.getName() === "Historical Simulation Data") {
       sheet.hideColumns(6);
       try { sheet.hideColumns(7, 4); } catch(e){}
    }

    if (maxCols > lastCol) {
      try { sheet.deleteColumns(lastCol + 1, maxCols - lastCol); } catch(e){}
    }
  }
  updateInteractiveDashboard(sheetObj);
}

function updateInteractiveDashboard(sheetObj) {
  const dash = sheetObj.dashSheet;
  const perfSheet = sheetObj.perfSheet;
  if (perfSheet.getLastRow() < 2) return;

  dash.clear();
  const charts = dash.getCharts();
  for (let i = 0; i < charts.length; i++) { dash.removeChart(charts[i]); }

  // Header and Dropdown setup
  dash.getRange("A1").setValue("Select Campaign:").setFontWeight("bold").setFontSize(16).setFontColor("#C33B48").setFontFamily("Manrope");
  dash.setColumnWidth(1, 200);

  const dropdownCell = dash.getRange("B1");
  dropdownCell.setFontFamily("Manrope").setFontSize(12).setBackground("#f3f3f3");
  dash.setColumnWidth(2, 450); 

  const campaignsRange = perfSheet.getRange(2, 2, perfSheet.getLastRow() - 1, 1);
  // Setting up Data Validation to list campaign names for selection.
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(campaignsRange).setAllowInvalid(false).build();
  dropdownCell.setDataValidation(rule);

  if (dropdownCell.getValue() === "") {
    const firstCampaign = campaignsRange.getValues()[0][0];
    dropdownCell.setValue(firstCampaign);
  }

  // --- DIAGNOSTICS PANEL ---
  // Using native Sheets XLOOKUP formulas in the background to pull latest execution diagnostics.
  dash.getRange("A3").setValue("Optimal ROAS:").setFontWeight("bold").setFontFamily("Manrope").setFontColor("#555555");
  dash.getRange("B3").setFormula(`=IFERROR(XLOOKUP(B1, 'Weekly Performance Log'!B:B, 'Weekly Performance Log'!N:N, "", 0, -1), "")`).setNumberFormat("0.00").setFontFamily("Manrope").setHorizontalAlignment("right");

  dash.getRange("A4").setValue("Guarded Target ROAS:").setFontWeight("bold").setFontFamily("Manrope").setFontColor("#555555");
  dash.getRange("B4").setFormula(`=IFERROR(XLOOKUP(B1, 'Weekly Performance Log'!B:B, 'Weekly Performance Log'!M:M, "", 0, -1), "")`).setNumberFormat("0.00").setFontFamily("Manrope").setHorizontalAlignment("right");

  dash.getRange("A5").setValue("Script Action:").setFontWeight("bold").setFontFamily("Manrope").setFontColor("#555555");
  dash.getRange("B5").setFormula(`=IFERROR(XLOOKUP(B1, 'Weekly Performance Log'!B:B, 'Weekly Performance Log'!E:E, "", 0, -1), "")`).setFontFamily("Manrope").setHorizontalAlignment("right");

  dash.getRange("A6").setValue("Guardrail Status:").setFontWeight("bold").setFontFamily("Manrope").setFontColor("#555555");
  dash.getRange("B6").setFormula(`=IF(ISBLANK(B1), "", IF(ROUND(B3, 2)=ROUND(B4, 2), "✅ Matches Optimal Target", "🛡️ Target Limited by Guardrails"))`).setFontWeight("bold").setFontFamily("Manrope").setHorizontalAlignment("right");

  // DATA TABLE FOR HISTORICAL TRACKER
  dash.getRange("K3:M3").setValues([["Date", "Profit (Actual)", "Profit (Guarded)"]])
      .setFontWeight("bold").setFontFamily("Manrope").setBackground("#C33B48").setFontColor("white"); 
  dash.setColumnWidth(11, 150); 
  dash.setColumnWidth(12, 140); 
  dash.setColumnWidth(13, 140); 
  
  // Natively querying the performance log to pull data for the historical chart.
  const queryFormula = `=QUERY('Weekly Performance Log'!A:R, "SELECT A, O, P WHERE B = '"&SUBSTITUTE(B1, "'", "''")&"' ORDER BY A LABEL A '', O '', P ''", 0)`;
  dash.getRange("K4").setFormula(queryFormula); 
  
  // PERFORMANCE CHART (Actual vs Predicted)
  const perfChart = dash.newChart()
    .asLineChart()
    .addRange(dash.getRange("K3:M"))
    .setPosition(9, 1, 0, 0)
    .setTitle('Profit Tracker: Actual vs. Predicted (Guarded)')
    .setColors(['#333333', '#C33B48']) 
    .setXAxisTitle('Date')
    .setYAxisTitle('Profit')
    .setNumHeaders(1)
    .setLegendPosition(Charts.Position.TOP)
    .setOption('width', 600)
    .setOption('height', 400)
    .setOption('vAxis.format', '£#,##0.00') 
    .build();
  dash.insertChart(perfChart);

  // CURVE CHART DATA GENERATION
  dash.getRange("O1:P1").setValues([["Simulated ROAS", "Predicted Profit"]]); 
  
  // ARRAYFORMULA generates the quadratic sequence based on coefficients stored in historical logs.
  const curveFormula = `=ARRAYFORMULA(IF(ISBLANK(B1), "", LET(
    opt, XLOOKUP(B1, 'Historical Simulation Data'!B:B, 'Historical Simulation Data'!K:K, "", 0, -1),
    a, XLOOKUP(B1, 'Historical Simulation Data'!B:B, 'Historical Simulation Data'!G:G, "", 0, -1),
    b, XLOOKUP(B1, 'Historical Simulation Data'!B:B, 'Historical Simulation Data'!H:H, "", 0, -1),
    c, XLOOKUP(B1, 'Historical Simulation Data'!B:B, 'Historical Simulation Data'!I:I, "", 0, -1),
    steps, SEQUENCE(21, 1, -10, 1),
    roas_seq, ROUND(IF(opt + (steps * 0.1) < 0.1, 0.1, opt + (steps * 0.1)), 2),
    profit_seq, (a * roas_seq^2) + (b * roas_seq) + c,
    CHOOSECOLS(HSTACK(roas_seq, profit_seq), 1, 2)
  )))`;
  
  dash.getRange("O2").setFormula(curveFormula);
  dash.getRange("O2:O22").setNumberFormat("0.00");
  dash.getRange("P2:P22").setNumberFormat("£#,##0.00");
  dash.hideColumns(15, 2); // Hide data table, show only the visual chart.

  // CURVE CHART (Parabola Visualisation)
  const curveChart = dash.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(dash.getRange("O1:P22"))
    .setPosition(9, 3, 0, 0) 
    .setOption('title', 'Current Projection: Finding the Best Target ROAS') 
    .setOption('colors', ['#000000']) 
    .setOption('hAxis.title', 'Target ROAS')
    .setOption('vAxis.title', 'Predicted Profit') 
    .setOption('vAxis.format', '£#,##0.00') 
    // Uses standard Google Sheets trendline to smooth the parabola.
    .setOption('trendlines', { 0: { type: 'polynomial', degree: 2, color: '#C33B48', visibleInLegend: false, opacity: 0.8 } })
    .setOption('width', 500)
    .setOption('height', 400)
    .setOption('legend', 'none')
    .build();
  
  dash.insertChart(curveChart);
}

function logToSpreadsheet(sheetObj, metrics) {
  const today = new Date();
  const format = "dd-MMM-yyyy"; 
  const tz = AdsApp.currentAccount().getTimeZone();

  const simEnd = new Date(today); simEnd.setDate(today.getDate() - 1);
  const simStart = new Date(today); simStart.setDate(today.getDate() - 7);
  
  const dayOfWeek = today.getDay(); 
  const diffToLastSun = dayOfWeek === 0 ? 7 : dayOfWeek; 
  const perfEnd = new Date(today); perfEnd.setDate(today.getDate() - diffToLastSun);
  const perfStart = new Date(perfEnd); perfStart.setDate(perfEnd.getDate() - 6);

  const rawString = metrics.rawPoints.map(p => Object.keys(p).map(key => `${key}=${p[key]}`).join(", ")).join(" | ");

  const actualProfit = (metrics.stats.revenue * CONFIG.conversionValueMultiplier) - metrics.stats.cost;
  const profitDiff = metrics.guardedProfit - actualProfit; 

  sheetObj.dataSheet.appendRow([
    Utilities.formatDate(new Date(), tz, "dd-MMM-yyyy HH:mm"),
    metrics.campaign,
    Utilities.formatDate(simStart, tz, format),
    Utilities.formatDate(simEnd, tz, format),
    metrics.currentRoas, 
    rawString, 
    metrics.coeffs.a,
    metrics.coeffs.b,
    metrics.coeffs.c,
    metrics.rSquared,
    metrics.optimalRoas,
    metrics.optimalProfit
  ]);

  sheetObj.perfSheet.appendRow([
    Utilities.formatDate(new Date(), tz, "dd-MMM-yyyy HH:mm"), // A
    metrics.campaign, // B
    Utilities.formatDate(perfStart, tz, format), // C
    Utilities.formatDate(perfEnd, tz, format), // D
    metrics.action, // E                
    metrics.stats.cost, // F        
    metrics.guardedSpend, // G            
    metrics.optimalSpend, // H            
    metrics.stats.revenue, // I      
    metrics.guardedRevenue, // J            
    metrics.optimalRevenue, // K            
    metrics.stats.roas, // L
    metrics.guardedRoas, // M 
    metrics.optimalRoas, // N        
    actualProfit, // O             
    metrics.guardedProfit, // P            
    metrics.optimalProfit, // Q            
    profitDiff // R                
  ]);
}

function getLastWeekStats(campaignId) {
  const query = `SELECT metrics.cost_micros, metrics.conversions_value FROM campaign WHERE campaign.id = ${campaignId} AND segments.date DURING LAST_WEEK_MON_SUN`;
  const rows = AdsApp.search(query);
  if (rows.hasNext()) {
    const row = rows.next();
    const cost = row.metrics.costMicros / 1000000;
    const revenue = row.metrics.conversionsValue;
    const roas = cost > 0 ? (revenue / cost) : 0;
    return { cost, revenue, roas };
  }
  return { cost: 0, revenue: 0, roas: 0 };
}

function getCurrentRoasSettings(campaignId) {
  // Query to check campaign-level settings and portfolio strategy links.
  const query = `
    SELECT 
      campaign.bidding_strategy_type, 
      campaign.target_roas.target_roas, 
      campaign.bidding_strategy 
    FROM campaign 
    WHERE campaign.id = ${campaignId}
  `;
  const rows = AdsApp.search(query);
  if (!rows.hasNext()) return null;
  const row = rows.next();

  let roas = row.campaign.targetRoas ? row.campaign.targetRoas.targetRoas : 0;
  let type = row.campaign.biddingStrategyType;
  let isPortfolio = !!row.campaign.biddingStrategy;
  let portfolioId = isPortfolio ? row.campaign.biddingStrategy.split('/').pop() : null;

  // If a portfolio strategy exists, query it specifically to get correct target ROAS.
  if (isPortfolio) {
    const pQuery = `SELECT bidding_strategy.type, bidding_strategy.target_roas.target_roas FROM bidding_strategy WHERE bidding_strategy.id = ${portfolioId}`;
    const pRows = AdsApp.search(pQuery);
    if (pRows.hasNext()) {
      const pRow = pRows.next();
      type = pRow.biddingStrategy.type;
      if (pRow.biddingStrategy.targetRoas) roas = pRow.biddingStrategy.targetRoas.targetRoas;
    }
  }
  return { roas, type, isPortfolio, portfolioId };
}

function applyGuardrails(optimal, current) {
  let target = optimal;
  // Apply absolute business limits.
  target = Math.max(CONFIG.guardrails.minRoasLimit, Math.min(CONFIG.guardrails.maxRoasLimit, target));
  // Apply limit on change velocity (soft limit).
  if (current > 0) {
    const maxChange = CONFIG.guardrails.maxRoasChange;
    target = Math.max(current - maxChange, Math.min(current + maxChange, target));
  }
  return parseFloat(target.toFixed(2));
}

/**
 * MATHS ENGINE - LEAST SQUARES METHOD
 * Fits data to y = ax^2 + bx + c by constructing and solving a system of linear equations.
 */
function fitQuadratic(data) {
  let s4 = 0, s3 = 0, s2 = 0, s1 = 0, s0 = 0, sy = 0, sxy = 0, sx2y = 0;
  for (const p of data) {
    const x = p.x; const y = p.y; const x2 = x*x;
    s4+=x2*x2; s3+=x2*x; s2+=x2; s1+=x; s0+=1; sy+=y; sxy+=x*y; sx2y+=x2*y;
  }
  // Construct matrix [A] and vector [B] for linear system Ax=B.
  return solve3x3([[s4,s3,s2],[s3,s2,s1],[s2,s1,s0]], [sx2y,sxy,sy]);
}

/**
 * Calculates Coeff of Determination (R²) for standard quadratic form.
 * PredY = ax^2 + bx + c
 */
function calculateRSquared(data, coeffs) {
  let ssTot = 0, ssRes = 0, sumY = 0;
  for (const p of data) sumY += p.y;
  const yMean = sumY / data.length;
  for (const p of data) {
    const pred = (coeffs.a*p.x*p.x) + (coeffs.b*p.x) + coeffs.c;
    ssRes += Math.pow(p.y - pred, 2);
    ssTot += Math.pow(p.y - yMean, 2);
  }
  // Handle edge case where data is a horizontal line.
  return ssTot === 0 ? 0 : 1 - (ssRes / ssTot);
}

/**
 * MATRIX SOLVER
 * Solves standard 3x3 linear system using Cramer's Rule via determinants.
 */
function solve3x3(A, B) {
  // 3x3 Determinant Helper
  const det = m => m[0][0]*(m[1][1]*m[2][2]-m[1][2]*m[2][1]) - m[0][1]*(m[1][0]*m[2][2]-m[1][2]*m[2][0]) + m[0][2]*(m[1][0]*m[2][1]-m[1][1]*m[2][0]);
  const D = det(A);
  // Fail safe for singular/ill-conditioned matrices.
  if (Math.abs(D) < 1e-9) return {a:0, b:0, c:0};
  // Replacer function for Cramer's rule columns.
  const rep = (c,v) => { let m = JSON.parse(JSON.stringify(A)); for(let i=0; i<3; i++) m[i][c] = v[i]; return m; };
  return { a: det(rep(0,B))/D, b: det(rep(1,B))/D, c: det(rep(2,B))/D };
}

function logHeader(t) { Logger.log("\n" + "=".repeat(40) + "\n" + t + "\n" + "=".repeat(40)); }

/**
 * GOOGLE ADS API UPDATER (MUTATE)
 * Modern way to update bidding strategies via bulk mutate operation.
 */
function applyTargetRoas(campaignId, settings, roas) {
  // Format ROAS to standard Number type with 2 decimal places. GA expects format as standard 'Number' not object.
  const formattedRoas = Number(roas.toFixed(2)); 
  
  if (settings.isPortfolio) {
    Logger.log(`    🏗️ Action: Updating Portfolio Strategy...`);
    updatePortfolioViaMutate(settings.portfolioId, settings.type, formattedRoas);
  } else {
    Logger.log(`    🏗️ Action: Updating Campaign (Mutate)...`);
    updateCampaignViaMutate(campaignId, settings.type, formattedRoas);
  }
}

/**
 * Modifies campaign-level standard bidding strategies.
 */
function updateCampaignViaMutate(campaignId, type, roas) {
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  const resourceName = `customers/${customerId}/campaigns/${campaignId}`;
  
  // Constructing the inner payload operation.
  let innerOp = { "update": { "resourceName": resourceName }, "update_mask": "" };

  if (type === 'TARGET_ROAS') {
    innerOp.update.targetRoas = { "targetRoas": roas };
    innerOp.update_mask = "target_roas.target_roas";
  } else if (type === 'MAXIMIZE_CONVERSION_VALUE') {
    innerOp.update.maximizeConversionValue = { "targetRoas": roas };
    innerOp.update_mask = "maximize_conversion_value.target_roas";
  } else {
    // Other strategy types cannot be mutated directly in this manner.
    Logger.log(`    ⚠️ Error: Strategy '${type}' cannot be updated directly.`);
    return;
  }

  try {
    const response = AdsApp.mutate({ "campaign_operation": innerOp });
    if (response.isSuccessful()) {
       Logger.log(`    ✅ Success: Campaign updated to ${roas.toFixed(2)}`);
    } else {
       Logger.log(`    ❌ API Failure: ${response.getErrorMessages().join(", ")}`);
    }
  } catch (e) {
    Logger.log(`    ❌ Critical Error: ${e.toString()}`);
  }
}

/**
 * Modifies central portfolio bidding strategies.
 */
function updatePortfolioViaMutate(strategyId, type, roas) {
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  const resourceName = `customers/${customerId}/biddingStrategies/${strategyId}`;
  let innerOp = { "update": { "resourceName": resourceName }, "update_mask": "" };

  if (type === 'TARGET_ROAS') {
    innerOp.update.targetRoas = { "targetRoas": roas };
    innerOp.update_mask = "target_roas.target_roas";
  } else if (type === 'MAXIMIZE_CONVERSION_VALUE') {
    innerOp.update.maximizeConversionValue = { "targetRoas": roas };
    innerOp.update_mask = "maximize_conversion_value.target_roas";
  } else {
    Logger.log(`    ⚠️ Error: Unsupported Portfolio Type: ${type}`);
    return;
  }

  try {
    const response = AdsApp.mutate({ "bidding_strategy_operation": innerOp });
    if (response.isSuccessful()) {
       Logger.log(`    ✅ Success: Portfolio updated to ${roas.toFixed(2)}`);
    } else {
       Logger.log(`    ❌ API Failure: ${response.getErrorMessages().join(", ")}`);
    }
  } catch (e) {
    Logger.log(`    ❌ Critical Error: ${e.toString()}`);
  }
}
