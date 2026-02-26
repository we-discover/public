/**
 * ==============================================================================
 * WEDISCOVER AUTOMATED ROAS OPTIMISATION SCRIPT (QUADRATIC REGRESSION)
 * ==============================================================================
 *
 * WHAT DOES THIS SCRIPT DO?
 * -------------------------
 * This script optimises your campaigns for Maximum Profit (rather than just Revenue).
 * 1. It pulls historical data from Google's "Traffic Simulator".
 * 2. It uses quadratic regression to model the relationship between ROAS and Profit.
 * 3. It identifies the specific ROAS target where profit is maximised.
 * 4. It logs the analysis to a spreadsheet and updates the campaign (if enabled).
 * 5. It emails the full execution log to the specified users.
 * 
 * AUTHOR: Nathan Ifill (@nathanifill)
 *
 * ==============================================================================
 * CONFIGURATION
 * ==============================================================================
 */

const CONFIG = {
  // 1. CAMPAIGN SELECTION
  // Leave this as [] to analyse ALL campaigns using Target ROAS. 
  // To limit the script to specific campaigns, add their names in quotes like this: ["Brand_Search", "Shopping_UK"].
  campaignNames: [], 
  
  // 2. VALUE ADJUSTMENT
  // If your conversion value doesn't account for profit margins, you can adjust it here.
  // For example, 0.5 would mean you keep 50% of the revenue as gross profit. 1.0 means use the value as-is.
  conversionValueMultiplier: 1.0,

  // 3. DATA QUALITY CHECK (R-Squared)
  // This measures how "reliable" the data is on a scale of 0 to 1. 
  // 0.5 is the minimum recommended; if the data is too messy or unpredictable, the script will skip the campaign.
  minRSquared: 0.5,
  
  // 4. SAFETY GUARDRAILS
  // These settings prevent the script from making drastic or risky changes.
  guardrails: {
    maxRoasChange: 0.2,  // The maximum amount the ROAS target can shift in one run (e.g., 2.0 to 2.2).
    minRoasLimit: 0.8,   // The absolute lowest ROAS target the script is allowed to set.
    maxRoasLimit: 4.0    // The absolute highest ROAS target the script is allowed to set.
  },

  // 5. REPORTING (The Spreadsheet)
  // Paste the full URL of a Google Sheet here to log results. 
  // If left blank, the script will create a brand new sheet for you and email you the link.
  spreadsheetUrl: "", 
  
  // Enter the email addresses (separated by commas) that should receive the results and access to the sheet.
  emailAddresses: "",

  // 6. ACTION MODE
  // If 'false', the script will only "pretend" to work, calculating everything and logging it to the sheet without changing your account.
  // Set this to 'true' only when you are happy with the recommendations and want the script to update your live bids.
  updateCampaigns: false
};

/**
 * ==============================================================================
 * MAIN SCRIPT LOGIC
 * ==============================================================================
 */

function main() {
  // --- LOG CAPTURE SETUP ---
  const logBuffer = [];
  const originalLogger = Logger.log;
  // This interceptor ensures everything logged to the console is saved for the email
  Logger.log = function(msg) {
    originalLogger(msg);
    logBuffer.push(msg);
  };

  logHeader("🚀 IGNITING PROFIT ENGINE...");

  // --- STEP 0: SETUP SPREADSHEET ---
  const sheetObj = ensureSpreadsheet();
  if (!sheetObj) {
    Logger.log("💥 Critical Failure: Unable to spin up the spreadsheet.");
    sendFinalEmail(logBuffer.join("\n"), "CRITICAL FAILURE");
    return;
  }
  Logger.log(`📜 Mission Report: ${sheetObj.url}`);

  // --- STEP 1: FETCH DATA ---
  const query = `
    SELECT 
      campaign.id, 
      campaign.name,
      campaign_simulation.target_roas_point_list.points 
    FROM campaign_simulation 
    WHERE 
      campaign_simulation.type = 'TARGET_ROAS'
  `;

  const report = AdsApp.search(query);
  let processedCount = 0;

  while (report.hasNext()) {
    const row = report.next();
    const currentName = row.campaign.name;

    if (CONFIG.campaignNames.length > 0 && CONFIG.campaignNames.indexOf(currentName) === -1) {
      continue; 
    }

    processSingleCampaign(row, sheetObj);
    processedCount++;
  }

  // --- STEP 2: TIDY UP ---
  polishSpreadsheet(sheetObj);

  if (processedCount === 0) {
    Logger.log("\n🤷‍♂️ No eligible campaigns found. The simulator is silent.");
  } else {
    logHeader(`🏁 MISSION COMPLETE. Processed ${processedCount} Campaign(s).`);
  }

  // --- STEP 3: SEND LOGS VIA EMAIL ---
  sendFinalEmail(logBuffer.join("\n"), `Processed ${processedCount} Campaigns`);
}

/**
 * CORE LOGIC
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

  Logger.log(`    🔮 Gazing into the Traffic Simulator (${points.length} scenarios found)...`);

  for (const point of points) {
    const roas = point.targetRoas;
    const cost = point.costMicros / 1000000;
    const value = point.biddableConversionsValue;
    const profit = (value * CONFIG.conversionValueMultiplier) - cost;
    
    regressionData.push({ x: roas, y: profit });
  }

  if (regressionData.length < 3) {
    Logger.log("    👻 Ghost Town: Not enough data points to build a model.");
    return;
  }

  const coeffs = fitQuadratic(regressionData);
  const rSquared = calculateRSquared(regressionData, coeffs);
  
  if (rSquared < CONFIG.minRSquared) {
    Logger.log(`    🎲 Too Chaotic: Data correlation is weak (R² ${rSquared.toFixed(2)}). Skipping for safety.`);
    return;
  }

  if (coeffs.a >= 0) {
    Logger.log("    🎢 U-Curve Detected: Google thinks profit increases infinitely. I doubt that. Skipping.");
    return;
  }

  const optimalRoasRaw = -coeffs.b / (2 * coeffs.a);
  const expProfit = (coeffs.a * (optimalRoasRaw * optimalRoasRaw)) + (coeffs.b * optimalRoasRaw) + coeffs.c;
  const finalRoas = applyGuardrails(optimalRoasRaw, currentRoas);

  Logger.log(`    💎 Sweet Spot Found: ${optimalRoasRaw.toFixed(2)}`);
  Logger.log(`    🛡️ Safety Shields:   ${finalRoas.toFixed(2)} (Guarded)`);
  
  let actionTaken = "READ ONLY";
  if (CONFIG.updateCampaigns) {
    if (Math.abs(finalRoas - currentRoas) > 0.01) {
      actionTaken = "UPDATED";
    } else {
      actionTaken = "NO CHANGE";
    }
  }

  const stats = getLastWeekStats(campaignId);

  logToSpreadsheet(sheetObj, {
    timestamp: new Date(),
    campaign: campaignName,
    rawPoints: points, 
    coeffs: coeffs,
    rSquared: rSquared,
    optimalRoas: optimalRoasRaw,
    expProfit: expProfit,
    finalRoas: finalRoas,
    currentRoas: currentRoas, 
    stats: stats,
    action: actionTaken       
  });

  if (actionTaken === "UPDATED") {
       applyTargetRoas(campaignId, currentSettings, finalRoas);
  } else if (actionTaken === "NO CHANGE") {
       Logger.log(`    ✅ Status: We are already perfect. No change needed.`);
  } else {
       Logger.log(`    👀 Status: Read Only Mode.`);
  }
}

/**
 * EMAIL TOOLS
 */

function sendFinalEmail(fullLog, statusSummary) {
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const recipient = CONFIG.emailAddresses;

  if (!recipient || recipient === "") {
    return; // No email configured
  }

  const subject = `Google Ads Script Log: ${accountName} (${statusSummary})`;
  const body = `Automated ROAS Optimisation Script has finished running.\n\n` +
               `Account: ${accountName} (${accountId})\n` +
               `Spreadsheet: ${CONFIG.spreadsheetUrl}\n\n` +
               `--- EXECUTION LOGS ---\n\n${fullLog}`;

  try {
    MailApp.sendEmail(recipient, subject, body);
    // Use the stored original logger to avoid infinite loops
    // Note: Logger was redefined in main(), so we use original logic if needed
  } catch (e) {
    // Fail silently in logs to avoid disrupting the main flow
  }
}

/**
 * SPREADSHEET & REPORTING TOOLS
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
  if (!dashSheet) {
    dashSheet = ss.insertSheet("Dashboard", 0);
  }

  const dataSheet = getOrCreateSheet(ss, "Historical Simulation Data", [
    "Timestamp", "Campaign", "Start Date", "End Date", 
    "Current Target ROAS", 
    "Raw Points (Hidden)", 
    "Coeff A", "Coeff B", "Coeff C", "R-Squared", 
    "Optimal ROAS", "Predicted Profit"
  ]);
  
  const perfSheet = getOrCreateSheet(ss, "Weekly Performance Log", [
    "Timestamp", "Campaign", "Start Date", "End Date", 
    "Action Taken", 
    "Actual Spend", "Predicted Spend", 
    "Actual Revenue", "Predicted Revenue", 
    "Actual ROAS", "Optimal ROAS", 
    "Actual Profit", "Predicted Profit", "Profit Diff"
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

    if (lastRow > 1) {
      const fullRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      fullRange.setFontFamily("Manrope");
      sheet.setColumnWidth(1, 200); 
      sheet.setColumnWidth(2, 500); 

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (let i = 0; i < headers.length; i++) {
        const headerName = headers[i].toLowerCase();
        const colRange = sheet.getRange(2, i + 1, lastRow - 1, 1);

        if (headerName.includes("roas") || headerName.includes("coeff") || headerName.includes("r-squared")) {
          colRange.setNumberFormat("0.00");
        } else if (headerName.includes("profit") || headerName.includes("spend") || headerName.includes("revenue")) {
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

  dash.getRange("A1").setValue("Select Campaign:").setFontWeight("bold").setFontSize(16).setFontColor("#C33B48").setFontFamily("Manrope");
  dash.setColumnWidth(1, 200);

  const dropdownCell = dash.getRange("B1");
  dropdownCell.setFontFamily("Manrope").setFontSize(12).setBackground("#f3f3f3");
  dash.setColumnWidth(2, 450); 

  const campaignsRange = perfSheet.getRange(2, 2, perfSheet.getLastRow() - 1, 1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(campaignsRange).setAllowInvalid(false).build();
  dropdownCell.setDataValidation(rule);

  if (dropdownCell.getValue() === "") {
    const firstCampaign = campaignsRange.getValues()[0][0];
    dropdownCell.setValue(firstCampaign);
  }

  const queryFormula = `=QUERY('Weekly Performance Log'!A:N, "SELECT A, L, M WHERE B = '"&B1&"' ORDER BY A LABEL L 'Actual Profit', M 'Predicted Profit'", 1)`;
  dash.getRange("A20").setFormula(queryFormula); 
  
  const charts = dash.getCharts();
  for (let i = 0; i < charts.length; i++) { dash.removeChart(charts[i]); }

  const chartBuilder = dash.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dash.getRange("A20:C"))
    .setPosition(3, 1, 0, 0)
    .setOption('title', 'Profit Impact: Actual vs. Predicted Model')
    .setOption('colors', ['#333333', '#C33B48']) 
    .setOption('hAxis.title', 'Date')
    .setOption('vAxis.title', 'Profit')
    .setOption('width', 1000)
    .setOption('height', 500)
    .setOption('legend', {position: 'top'});

  dash.insertChart(chartBuilder.build());
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
  const predictedProfit = metrics.expProfit;
  const profitDiff = predictedProfit - actualProfit;

  let predictedSpend = 0;
  let predictedRevenue = 0;
  const effectiveMargin = (metrics.optimalRoas * CONFIG.conversionValueMultiplier) - 1;
  if (effectiveMargin > 0) {
      predictedSpend = predictedProfit / effectiveMargin;
      predictedRevenue = predictedSpend * metrics.optimalRoas;
  }

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
    metrics.expProfit
  ]);

  sheetObj.perfSheet.appendRow([
    Utilities.formatDate(new Date(), tz, "dd-MMM-yyyy HH:mm"),
    metrics.campaign,
    Utilities.formatDate(perfStart, tz, format),
    Utilities.formatDate(perfEnd, tz, format),
    metrics.action,                  
    metrics.stats.cost,          
    predictedSpend,              
    metrics.stats.revenue,       
    predictedRevenue,            
    metrics.stats.roas,          
    metrics.finalRoas,           
    actualProfit,                
    predictedProfit,             
    profitDiff                   
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
  const campaignIter = AdsApp.campaigns().withIds([campaignId]).get();
  if (!campaignIter.hasNext()) return null;
  const campaign = campaignIter.next();
  let roas = 0;
  let type = campaign.getBiddingStrategyType();
  let isPortfolio = false;
  let portfolioId = null;
  const portfolioStrategy = campaign.bidding().getStrategy();

  if (portfolioStrategy) {
    isPortfolio = true;
    portfolioId = portfolioStrategy.getId();
    const query = `SELECT bidding_strategy.type, bidding_strategy.target_roas.target_roas FROM bidding_strategy WHERE bidding_strategy.id = ${portfolioId}`;
    const rows = AdsApp.search(query);
    if (rows.hasNext()) {
      const row = rows.next();
      type = row.biddingStrategy.type;
      if (row.biddingStrategy.targetRoas) roas = row.biddingStrategy.targetRoas.targetRoas;
    }
  } else {
    roas = campaign.bidding().getTargetRoas() || 0;
  }
  return { roas, type, isPortfolio, portfolioId };
}

function applyGuardrails(optimal, current) {
  let target = optimal;
  target = Math.max(CONFIG.guardrails.minRoasLimit, Math.min(CONFIG.guardrails.maxRoasLimit, target));
  if (current > 0) {
    const maxChange = CONFIG.guardrails.maxRoasChange;
    target = Math.max(current - maxChange, Math.min(current + maxChange, target));
  }
  return parseFloat(target.toFixed(2));
}

function fitQuadratic(data) {
  let s4 = 0, s3 = 0, s2 = 0, s1 = 0, s0 = 0, sy = 0, sxy = 0, sx2y = 0;
  for (const p of data) {
    const x = p.x; const y = p.y; const x2 = x*x;
    s4+=x2*x2; s3+=x2*x; s2+=x2; s1+=x; s0+=1; sy+=y; sxy+=x*y; sx2y+=x2*y;
  }
  return solve3x3([[s4,s3,s2],[s3,s2,s1],[s2,s1,s0]], [sx2y,sxy,sy]);
}

function calculateRSquared(data, coeffs) {
  let ssTot = 0, ssRes = 0, sumY = 0;
  for (const p of data) sumY += p.y;
  const yMean = sumY / data.length;
  for (const p of data) {
    const pred = (coeffs.a*p.x*p.x) + (coeffs.b*p.x) + coeffs.c;
    ssRes += Math.pow(p.y - pred, 2);
    ssTot += Math.pow(p.y - yMean, 2);
  }
  return ssTot === 0 ? 0 : 1 - (ssRes / ssTot);
}

function solve3x3(A, B) {
  const det = m => m[0][0]*(m[1][1]*m[2][2]-m[1][2]*m[2][1]) - m[0][1]*(m[1][0]*m[2][2]-m[1][2]*m[2][0]) + m[0][2]*(m[1][0]*m[2][1]-m[1][1]*m[2][0]);
  const D = det(A);
  if (Math.abs(D) < 1e-9) return {a:0, b:0, c:0};
  const rep = (c,v) => { let m = JSON.parse(JSON.stringify(A)); for(let i=0; i<3; i++) m[i][c] = v[i]; return m; };
  return { a: det(rep(0,B))/D, b: det(rep(1,B))/D, c: det(rep(2,B))/D };
}

function formatCurrency(n) { return n.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,'); }
function logHeader(t) { Logger.log("\n" + "=".repeat(40) + "\n" + t + "\n" + "=".repeat(40)); }

function applyTargetRoas(campaignId, settings, roas) {
  if (settings.isPortfolio) {
    Logger.log(`    🏗️ Action: Updating Portfolio Strategy...`);
    updatePortfolioViaMutate(settings.portfolioId, settings.type, roas);
  } else {
    const campaign = AdsApp.campaigns().withIds([campaignId]).get().next();
    if (['TARGET_ROAS', 'MAXIMIZE_CONVERSION_VALUE'].indexOf(settings.type) > -1) {
      campaign.bidding().setTargetRoas(roas);
      Logger.log(`    🏗️ Action: Standard Campaign updated to ${roas}`);
    } else {
      Logger.log(`    ⚠️ Error: Strategy '${settings.type}' cannot be updated directly.`);
    }
  }
}

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
       Logger.log(`    ✅ Success: Portfolio updated to ${roas}`);
    } else {
       Logger.log(`    ❌ API Failure: ${response.getErrorMessages().join(", ")}`);
    }
  } catch (e) {
    Logger.log(`    ❌ Critical Error: ${e.toString()}`);
  }
}
