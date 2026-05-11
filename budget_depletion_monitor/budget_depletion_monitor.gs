const SPREADSHEET_URL = '';

function main() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  // 1. Rename the Spreadsheet
  const accountName = AdsApp.currentAccount().getName();
  const customerId = AdsApp.currentAccount().getCustomerId();
  const newName = `${accountName} (${customerId}) Budget Depletion Monitor | WeDiscover`;
  spreadsheet.rename(newName);
  
  const timeZone = AdsApp.currentAccount().getTimeZone();
  
  // 2. Define Dates (Calculated explicitly to control API behavior)
  
  // TODAY
  const todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  
  // YESTERDAY
  const yesterdayDate = new Date();
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterdayDate, timeZone, 'yyyy-MM-dd');

  // LAST 7 DAYS (EXCLUDING TODAY)
  // Range: Today-7 to Yesterday
  const sevenDaysAgoDate = new Date();
  sevenDaysAgoDate.setDate(sevenDaysAgoDate.getDate() - 7); 
  const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgoDate, timeZone, 'yyyy-MM-dd');

  // 3. GENERATE BUDGET HISTORY
  const budgetMap = processRecentBudgets(spreadsheet, timeZone);

  // 4. Process Reports
  processReport(spreadsheet, 'Today Data', `${todayStr},${todayStr}`, true, budgetMap);
  processReport(spreadsheet, 'Yesterday Data', `${yesterdayStr},${yesterdayStr}`, false, budgetMap);
  
  // Note: We pass the range [Today-7, Yesterday]
  processReport(spreadsheet, 'Last 7 Days Data', `${sevenDaysAgoStr},${yesterdayStr}`, false, budgetMap);
}

/**
 * Generates the "Recent Budgets" sheet and returns a lookup map for other functions.
 */
function processRecentBudgets(ss, timeZone) {
  let sheet = ss.getSheetByName('Recent Budgets Data');
  sheet.clear();

  // 1. Get Current Budgets & Enabled Campaigns
  const campaignQuery = `
    SELECT campaign.id, campaign.name, campaign_budget.amount_micros 
    FROM campaign 
    WHERE campaign.status = 'ENABLED' 
    AND campaign.serving_status != 'ENDED'
    AND campaign.advertising_channel_type != 'DISPLAY'
  `;
  
  const campIter = AdsApp.search(campaignQuery);
  const campaignData = {}; 
  
  while (campIter.hasNext()) {
    const row = campIter.next();
    campaignData[row.campaign.id] = {
      name: row.campaign.name,
      currentBudget: (row.campaignBudget.amountMicros || 0) / 1000000
    };
  }

  // 2. Get Change History (Calculated Safe Range)
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - 29); // Safe 29-day lookback
  
  const startStr = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  const endStr = Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');

  const changeQuery = `
    SELECT 
      change_event.change_date_time, 
      change_event.campaign, 
      change_event.old_resource, 
      change_event.new_resource 
    FROM change_event 
    WHERE 
      change_event.change_resource_type = 'CAMPAIGN_BUDGET' 
      AND change_event.change_date_time BETWEEN '${startStr}' AND '${endStr}'
    ORDER BY change_event.change_date_time DESC
    LIMIT 10000
  `;

  const changeIter = AdsApp.search(changeQuery);
  const changes = {}; 

  while (changeIter.hasNext()) {
    const row = changeIter.next();
    
    // Safety Check for Shared Budgets (prevents crash)
    if (!row.changeEvent.campaign) continue; 

    const campResource = row.changeEvent.campaign;
    const campId = campResource.split('/').pop();

    if (!campaignData[campId]) continue; 

    const dateStr = row.changeEvent.changeDateTime.substring(0, 10); // YYYY-MM-DD
    const oldBudget = (row.changeEvent.oldResource.campaignBudget.amountMicros || 0) / 1000000;
    const newBudget = (row.changeEvent.newResource.campaignBudget.amountMicros || 0) / 1000000;

    if (!changes[campId]) changes[campId] = {};
    if (!changes[campId][dateStr]) changes[campId][dateStr] = [];
    
    changes[campId][dateStr].push({ old: oldBudget, new: newBudget });
  }

  // 3. Build the Grid
  const budgetLookup = {}; 
  const headers = ["Campaign Name", "ID"];
  const dates = [];
  
  // Headers for the last 30 days
  for (let i = 0; i < 30; i++) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const dStr = Utilities.formatDate(d, timeZone, 'yyyy-MM-dd');
    dates.push(dStr);
    headers.push(dStr);
  }

  const output = [headers];

  for (const campId in campaignData) {
    const c = campaignData[campId];
    const row = [c.name, campId];
    
    // Start with current budget and walk backwards
    let runningBudget = c.currentBudget;

    for (const dateStr of dates) {
      const dayChanges = (changes[campId] && changes[campId][dateStr]) ? changes[campId][dateStr] : null;
      let finalBudgetForDay = 0;

      if (dayChanges) {
        let sum = 0;
        let count = 0;
        dayChanges.forEach(ch => { sum += ch.new; count++; });
        const earliestChange = dayChanges[dayChanges.length - 1];
        sum += earliestChange.old; 
        count++;
        finalBudgetForDay = sum / count;
        runningBudget = earliestChange.old;
      } else {
        finalBudgetForDay = runningBudget;
      }

      row.push(finalBudgetForDay);
      budgetLookup[`${campId}_${dateStr}`] = finalBudgetForDay;
    }
    output.push(row);
  }

  // Write to Sheet
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold").setBackground("#e0e0e0");
  sheet.getRange(2, 3, output.length - 1, output[0].length - 2).setNumberFormat("#,##0.00");
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);

  return budgetLookup;
}

function processReport(ss, sheetName, dateRangeInput, isToday, budgetMap) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found. Skipping.`);
    return;
  }
  sheet.clear(); 

  const timeZone = AdsApp.currentAccount().getTimeZone();
  const currentHour = parseInt(Utilities.formatDate(new Date(), timeZone, "H")); 

  const campaigns = getCampaignData(dateRangeInput);
  const segments = getSegmentedData(dateRangeInput);

  const output = [];
  
  const headers = [
    "Customer ID", "Account", "Campaign", "Status", "Budget", 
    "Search IS", "Budget Lost IS", "Conversions", "CPA", "Revenue", "ROAS", 
    "CTR", "Conv. Rate", "Lost Conversions", "Lost Revenue"
  ];

  let pacingHeaders = [];
  
  // Determine if it's the "Weekly" view by checking the sheet name passed
  const isWeekly = (sheetName === 'Last 7 Days Data');

  if (isWeekly) {
    // Generate dates: Start from Yesterday (1) back to 7 days ago (7)
    // This excludes Today (0)
    for (let i = 1; i <= 7; i++) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      pacingHeaders.push(Utilities.formatDate(d, timeZone, 'yyyy-MM-dd'));
    }
  } else {
    // Hours 0-23
    for (let i = 0; i < 24; i++) pacingHeaders.push(i.toString());
  }

  output.push([...headers, ...pacingHeaders]);

  const sortedCampaignIds = Object.keys(campaigns).sort((a, b) => {
    return campaigns[a].name.toUpperCase().localeCompare(campaigns[b].name.toUpperCase());
  });

  for (const campId of sortedCampaignIds) {
    const c = campaigns[campId];
    const campSegments = segments[campId] || {};
    
    let displayBudget = 0;

    if (isWeekly) {
      displayBudget = c.currentBudget; 
    } else {
      const dateKey = dateRangeInput.split(',')[0]; 
      const lookupKey = `${campId}_${dateKey}`;
      
      if (budgetMap[lookupKey]) {
        displayBudget = budgetMap[lookupKey];
      } else {
        displayBudget = c.currentBudget; 
      }
    }

    const conversions = c.conversions;
    const value = c.convValue;
    const cost = c.cost;
    const clicks = c.clicks;
    const impressions = c.impressions;
    
    const cpa = conversions > 0 ? cost / conversions : 0;
    const roas = cost > 0 ? value / cost : 0;
    const ctr = impressions > 0 ? clicks / impressions : 0;
    const cvr = clicks > 0 ? conversions / clicks : 0;
    const searchIS = c.searchIS; 
    const lostISBudget = c.lostISBudget;

    let lostConversions = 0;
    let lostRevenue = 0;
    
    if (searchIS > 0 && lostISBudget > 0) {
      lostConversions = (conversions / searchIS) * lostISBudget;
      lostRevenue = (value / searchIS) * lostISBudget;
    }

    const row = [
      AdsApp.currentAccount().getCustomerId(),
      AdsApp.currentAccount().getName(),
      c.name,
      c.status,
      displayBudget, 
      searchIS,
      lostISBudget,
      conversions,
      cpa,
      value, 
      roas,
      ctr,
      cvr,
      lostConversions,
      lostRevenue
    ];

    let pacingData = [];
    
    if (isWeekly) {
      pacingHeaders.forEach(dateKey => { 
        const dayCost = campSegments[dateKey] || 0;
        const lookupKey = `${campId}_${dateKey}`;
        const historicBudget = budgetMap[lookupKey] || displayBudget; 

        const percent = historicBudget > 0 ? (dayCost / historicBudget) : 0;
        pacingData.push(percent);
      });
    } else {
      let cumulativeCost = 0;
      for (let i = 0; i < 24; i++) {
        const hourCost = campSegments[i] || 0;
        cumulativeCost += hourCost;
        
        if (isToday && i > currentHour) {
          pacingData.push(""); 
        } else {
          const percent = displayBudget > 0 ? (cumulativeCost / displayBudget) : 0;
          pacingData.push(percent);
        }
      }
    }

    output.push([...row, ...pacingData]);
  }

  if (output.length > 0) {
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    formatSheet(sheet, output.length, output[0].length);
  }
}

function buildDateCondition(dateRangeInput) {
  // All inputs are now comma separated strings
  const parts = dateRangeInput.split(',');
  if (parts.length === 2) {
    return `BETWEEN '${parts[0]}' AND '${parts[1]}'`;
  }
  return `DURING TODAY`;
}

function getCampaignData(dateRangeInput) {
  const dateCondition = buildDateCondition(dateRangeInput);
  
  const query = `
    SELECT 
      campaign.id, 
      campaign.name, 
      campaign.status, 
      campaign_budget.amount_micros,
      metrics.cost_micros, 
      metrics.conversions, 
      metrics.conversions_value, 
      metrics.clicks, 
      metrics.impressions,
      metrics.search_impression_share,
      metrics.search_budget_lost_impression_share
    FROM campaign 
    WHERE campaign.status = 'ENABLED' 
    AND campaign.serving_status != 'ENDED'
    AND campaign.advertising_channel_type != 'DISPLAY'
    AND segments.date ${dateCondition}
  `;

  const rows = AdsApp.search(query);
  const data = {};

  while (rows.hasNext()) {
    const row = rows.next();
    const c = row.campaign;
    const m = row.metrics;
    const b = row.campaignBudget;

    data[c.id] = {
      name: c.name,
      status: c.status,
      currentBudget: (b.amountMicros || 0) / 1000000, 
      cost: (m.costMicros || 0) / 1000000,
      conversions: m.conversions || 0,
      convValue: m.conversionsValue || 0,
      clicks: m.clicks || 0,
      impressions: m.impressions || 0,
      searchIS: m.searchImpressionShare || 0,
      lostISBudget: m.searchBudgetLostImpressionShare || 0
    };
  }
  return data;
}

function getSegmentedData(dateRangeInput) {
  const dateCondition = buildDateCondition(dateRangeInput);
  
  const parts = dateRangeInput.split(',');
  const isWeekly = (parts[0] !== parts[1]); // If start != end, it's a range (Last 7 Days)
  
  const segmentField = isWeekly ? 'segments.date' : 'segments.hour';
  
  const query = `
    SELECT 
      campaign.id, 
      metrics.cost_micros, 
      ${segmentField}
    FROM campaign 
    WHERE campaign.status = 'ENABLED'
    AND campaign.serving_status != 'ENDED'
    AND campaign.advertising_channel_type != 'DISPLAY'
    AND segments.date ${dateCondition}
  `;

  const rows = AdsApp.search(query);
  const data = {}; 

  while (rows.hasNext()) {
    const row = rows.next();
    const cId = row.campaign.id;
    const cost = (row.metrics.costMicros || 0) / 1000000;
    let key;

    if (isWeekly) {
      key = row.segments.date; 
    } else {
      key = row.segments.hour; 
    }

    if (!data[cId]) data[cId] = {};
    if (!data[cId][key]) data[cId][key] = 0;
    
    data[cId][key] += cost;
  }
  return data;
}

function formatSheet(sheet, rows, cols) {
  sheet.getRange(1, 1, 1, cols).setFontWeight("bold").setBackground("#e0e0e0");
  
  const currencyCols = [5, 9, 10, 15];
  currencyCols.forEach(c => {
    sheet.getRange(2, c, rows-1, 1).setNumberFormat("#,##0.00");
  });

  const percentCols = [6, 7, 11, 12, 13];
  percentCols.forEach(c => {
    sheet.getRange(2, c, rows-1, 1).setNumberFormat("0.00%");
  });

  if (cols >= 16) {
    sheet.getRange(2, 16, rows-1, cols-15).setNumberFormat("0%");
  }
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
}
