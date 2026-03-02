/**
 * ============================================================================
 * 🕵️‍♂️ Change History Auditor - WeDiscover 
 * ============================================================================
 *
 * WHAT THIS SCRIPT DOES:
 * This programme audits the Google Ads change history for the last 29 days. 
 * It automatically renames the target spreadsheet to match the account, 
 * translates the raw API JSON into beautiful, plain English, executes an 
 * intelligent secondary lookup for deleted items, and applies bespoke formatting 
 * to the spreadsheet (including universal word wrap, data filters, and date formats).
 *
 * @author Nathan Ifill, WeDiscover
 * ============================================================================
 */

// ============================================================================
// ⚙️ CONFIGURATION SETTINGS (User Inputs Required Here)
// ============================================================================

// STEP 1: Insert the full URL of the Google Sheet where you want the report.
const SPREADSHEET_URL = ""; 

// STEP 2: Define the exact name of the tab (sheet) to overwrite. 
const SHEET_NAME = ""; 

// STEP 3: List any system users' email address or names of automated tools 
// you wish to ignore, written in speech marks, separated by commas.
//
// We recommend excluding "Bulk Actions" and "Low activity system bulk change"
// For example:
//
// const IGNORE_USERS = [
//   'lucy@example.com',
//   'Bulk Actions', 
//   'Low activity system bulk change'
// ];

const IGNORE_USERS = [
  'Bulk Actions', 
  'Low activity system bulk change'
];

// ============================================================================
// 🧠 CORE LOGIC & ORCHESTRATION
// ============================================================================

/**
 * The main entry point for the Google Ads Script.
 */
function main() { 
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  console.log(`🚀 Firing up the WeDiscover Change History Auditor! Analysing account: ${accountName} (${accountId})`);
  
  const changeAlerts = getChangeAlerts();

  if (changeAlerts.length > 0) {
    console.log(`🎯 Bingo! Found ${changeAlerts.length} actionable changes. Prepping the spreadsheet...`);
    reportResults(changeAlerts);
  } else {
    console.log(`😴 All quiet on the western front! No unrecognised changes found.`);
  }
  
  console.log(`🏁 Audit complete. Have a brilliant day!`);
}

/**
 * Generates a safe 29-day date range to perfectly abide by API limitations.
 * @returns {Object} { startDate: "YYYY-MM-DD", endDate: "YYYY-MM-DD" }
 */
function getSafeDateRange() {
  const endDateObj = new Date();
  const startDateObj = new Date();
  
  startDateObj.setDate(endDateObj.getDate() - 29); 

  const formatDate = (date) => {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };

  return {
    startDate: formatDate(startDateObj),
    endDate: formatDate(endDateObj)
  };
}

/**
 * 🕵️‍♂️ Advanced Secondary Lookup (The Universal Safety Net)
 * Bounces orphaned IDs against Google's global constants to recover deleted names.
 */
function fetchDeepCriterionNames(rawChanges) {
  const map = {};
  const idsToFetch = new Set();
  
  const campCriteria = [];
  const adgCriteria = [];
  
  const getVal = (obj, ...keys) => keys.reduce((a, c) => (a && a[c] !== undefined) ? a[c] : null, obj);

  for (const raw of rawChanges) {
    const resName = raw.resourceName;
    if (resName) {
      if (resName.includes('/campaignCriteria/')) campCriteria.push(resName);
      if (resName.includes('/adGroupCriteria/')) adgCriteria.push(resName);
    }
    
    const critId = getVal(raw.row.changeEvent?.newResource, 'campaignCriterion', 'criterionId') || 
                   getVal(raw.row.changeEvent?.oldResource, 'campaignCriterion', 'criterionId') ||
                   getVal(raw.row.changeEvent?.newResource, 'adGroupCriterion', 'criterionId') || 
                   getVal(raw.row.changeEvent?.oldResource, 'adGroupCriterion', 'criterionId');
                   
    if (critId && !isNaN(critId)) idsToFetch.add(critId);
  }

  const chunkArray = (arr, size) => Array.from({ length: Math.ceil(arr.length / size) }, (v, i) => arr.slice(i * size, i * size + size));

  if (campCriteria.length > 0) {
    const unique = [...new Set(campCriteria)];
    for (const chunk of chunkArray(unique, 200)) {
      const query = `
        SELECT campaign_criterion.resource_name, campaign_criterion.display_name, campaign_criterion.type,
               campaign_criterion.location.geo_target_constant, campaign_criterion.keyword.text
        FROM campaign_criterion WHERE campaign_criterion.resource_name IN ('${chunk.join("', '")}')
      `;
      try {
        const result = AdsApp.search(query);
        while (result.hasNext()) {
          const row = result.next();
          const cc = row.campaignCriterion;
          let name = cc.displayName;
          
          if (cc.type === 'LOCATION' && cc.location && cc.location.geoTargetConstant) {
             const id = cc.location.geoTargetConstant.split('/')[1];
             idsToFetch.add(id);
             name = `LOCATION_ID_${id}`; 
          } else if (!name && cc.type === 'KEYWORD' && cc.keyword) {
             name = `keyword '${cc.keyword.text}'`;
          }
          if (name) map[cc.resourceName] = { type: cc.type, name: name };
        }
      } catch (e) {}
    }
  }

  if (adgCriteria.length > 0) {
    const unique = [...new Set(adgCriteria)];
    for (const chunk of chunkArray(unique, 200)) {
      const query = `
        SELECT ad_group_criterion.resource_name, ad_group_criterion.display_name, ad_group_criterion.type,
               ad_group_criterion.keyword.text
        FROM ad_group_criterion WHERE ad_group_criterion.resource_name IN ('${chunk.join("', '")}')
      `;
      try {
        const result = AdsApp.search(query);
        while (result.hasNext()) {
          const row = result.next();
          const ac = row.adGroupCriterion;
          let name = ac.displayName;
          
          if (!name && ac.type === 'KEYWORD' && ac.keyword) name = `keyword '${ac.keyword.text}'`;
          if (name) map[ac.resourceName] = { type: ac.type, name: name };
        }
      } catch (e) {}
    }
  }

  const fallbackMap = {};
  if (idsToFetch.size > 0) {
    const idArray = [...idsToFetch];
    
    idArray.forEach(id => {
      if (id === '30000') fallbackMap[id] = "device 'Desktop'";
      if (id === '30001') fallbackMap[id] = "device 'Mobile'";
      if (id === '30002') fallbackMap[id] = "device 'Tablet'";
      if (id === '30004') fallbackMap[id] = "device 'Connected TV'";
    });

    for (const chunk of chunkArray(idArray, 200)) {
      const idList = chunk.join(",");
      
      try {
        const geoQuery = `SELECT geo_target_constant.id, geo_target_constant.name, geo_target_constant.country_code FROM geo_target_constant WHERE geo_target_constant.id IN (${idList})`;
        const geoResult = AdsApp.search(geoQuery);
        while (geoResult.hasNext()) {
           const row = geoResult.next();
           fallbackMap[row.geoTargetConstant.id] = `location '${row.geoTargetConstant.name} (${row.geoTargetConstant.countryCode})'`;
        }
      } catch(e) {}

      try {
        const langQuery = `SELECT language_constant.id, language_constant.name FROM language_constant WHERE language_constant.id IN (${idList})`;
        const langResult = AdsApp.search(langQuery);
        while (langResult.hasNext()) {
           const row = langResult.next();
           fallbackMap[row.languageConstant.id] = `language '${row.languageConstant.name}'`;
        }
      } catch(e) {}

      try {
        const audQuery = `SELECT user_list.id, user_list.name FROM user_list WHERE user_list.id IN (${idList})`;
        const audResult = AdsApp.search(audQuery);
        while (audResult.hasNext()) {
           const row = audResult.next();
           fallbackMap[row.userList.id] = `audience '${row.userList.name}'`;
        }
      } catch(e) {}
    }
  }

  for (const resName in map) {
     if (map[resName].name && map[resName].name.startsWith('LOCATION_ID_')) {
        const id = map[resName].name.split('_')[2];
        map[resName].name = fallbackMap[id] ? fallbackMap[id] : `location (ID: ${id})`;
     } else if (map[resName].name && !map[resName].name.startsWith('keyword') && !map[resName].name.startsWith('location')) {
        map[resName].name = `criterion '${map[resName].name}'`;
     }
  }

  return { criteriaMap: map, fallbackMap: fallbackMap };
}

/**
 * 🗣️ The Plain English Translator
 */
function generatePlainEnglishDescription(operation, resourceType, changedFieldsObj, email, campaignName, adGroupName, newRes, oldRes, criterionDisplayName) {
  
  let firstName = "Someone";
  let company = "an unknown company";
  
  if (email && email.includes('@')) {
    const parts = email.split('@');
    const localPart = parts[0]; 
    const domainPart = parts[1].toLowerCase();

    const namePart = localPart.split('.')[0]; 
    firstName = namePart.charAt(0).toUpperCase() + namePart.slice(1).toLowerCase();

    if (domainPart.includes('we-discover')) {
      company = "WeDiscover"; 
    } else {
      const domainName = domainPart.split('.')[0];
      company = domainName.split('-').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join('-');
    }
  }

  const userContext = `${firstName} from ${company}`;

  let targetName = '';
  if (resourceType === 'CAMPAIGN' && campaignName) {
    targetName = ` '${campaignName}'`;
  } else if (resourceType === 'AD_GROUP' && adGroupName) {
    targetName = ` '${adGroupName}'`;
  } else if ((resourceType === 'AD_GROUP_AD' || resourceType === 'AD_GROUP_CRITERION') && adGroupName) {
    targetName = ` in ad group '${adGroupName}'`;
  } else if (campaignName) {
    targetName = ` '${campaignName}'`; 
  }
  
  const getVal = (obj, ...keys) => keys.reduce((a, c) => (a && a[c] !== undefined) ? a[c] : null, obj);

  let fields = [];
  if (changedFieldsObj) {
    let parsedFields = changedFieldsObj;
    if (typeof changedFieldsObj === 'string') {
      try { parsedFields = JSON.parse(changedFieldsObj); } catch(e) {}
    }
    
    if (parsedFields && Array.isArray(parsedFields.paths)) {
      fields = parsedFields.paths;
    } else if (parsedFields && Array.isArray(parsedFields)) {
      fields = parsedFields;
    } else if (typeof changedFieldsObj === 'string') {
      fields = changedFieldsObj.split(',').map(f => f.trim());
    }
  }

  if (resourceType === 'CAMPAIGN_CRITERION') {
    const ip = getVal(newRes, 'campaignCriterion', 'ipBlock', 'ipAddress') || getVal(oldRes, 'campaignCriterion', 'ipBlock', 'ipAddress');
    if (ip) {
      if (operation === 'CREATE') return `${userContext} excluded IP address '${ip}' from the campaign${targetName}.`;
      if (operation === 'REMOVE') return `${userContext} removed the IP address exclusion '${ip}' from the campaign${targetName}.`;
    }
  }

  if (resourceType === 'CAMPAIGN_CRITERION' || resourceType === 'AD_GROUP_CRITERION') {
    const isIpBlock = getVal(newRes, 'campaignCriterion', 'ipBlock') || getVal(oldRes, 'campaignCriterion', 'ipBlock');
    if (!isIpBlock) {
       const critName = criterionDisplayName ? criterionDisplayName : 'a criterion';
       let targetScope = "the item";
       
       if (resourceType === 'CAMPAIGN_CRITERION') targetScope = campaignName ? `campaign '${campaignName}'` : 'the campaign';
       else if (resourceType === 'AD_GROUP_CRITERION') targetScope = adGroupName ? `ad group '${adGroupName}'` : 'the ad group';
       
       if (operation === 'CREATE') return `${userContext} added ${critName} to the ${targetScope}.`;
       if (operation === 'REMOVE') return `${userContext} removed ${critName} from the ${targetScope}.`;
       if (operation === 'UPDATE') return `${userContext} updated ${critName} in the ${targetScope}.`;
    }
  }

  const isStatusChange = fields.some(f => f === 'status' || f.endsWith('.status'));
  if (isStatusChange && operation === 'UPDATE') {
    const resourceKeys = Object.keys(newRes || {});
    for (let key of resourceKeys) {
      if (newRes[key] && newRes[key].status) {
        const newStatus = newRes[key].status;
        const oldStatus = oldRes && oldRes[key] ? oldRes[key].status : null;
        let type = (resourceType || 'resource').replace(/_/g, ' ').toLowerCase();
        if (type === 'ad group ad') type = 'ad';
        
        if (newStatus !== oldStatus) {
           if (newStatus === 'PAUSED') return `${userContext} paused the ${type}${targetName}.`;
           if (newStatus === 'ENABLED') return `${userContext} enabled the ${type}${targetName}.`;
           if (newStatus === 'REMOVED') return `${userContext} removed the ${type}${targetName}.`;
        }
      }
    }
  }

  if (resourceType === 'CAMPAIGN_BUDGET' && operation === 'UPDATE') {
    const oldAmount = getVal(oldRes, 'campaignBudget', 'amountMicros');
    const newAmount = getVal(newRes, 'campaignBudget', 'amountMicros');
    if (oldAmount && newAmount) {
      return `${userContext} updated the budget amount from £${(oldAmount/1000000).toFixed(2)} to £${(newAmount/1000000).toFixed(2)}.`;
    }
  }

  const opMap = {
    'CREATE': 'created a new',
    'UPDATE': 'updated the',
    'REMOVE': 'removed the',
    'UNSPECIFIED': 'modified the',
    'UNKNOWN': 'modified the'
  };
  const action = opMap[operation] || 'modified the';

  let fallbackType = (resourceType || 'resource').replace(/_/g, ' ').toLowerCase();
  if (fallbackType === 'ad group ad') fallbackType = 'ad';
  if (fallbackType === 'campaign criterion' || fallbackType === 'ad group criterion') fallbackType = 'criterion';

  if (operation === 'UPDATE' && fields.length > 0) {
    return `${userContext} ${action} ${fallbackType}${targetName} (Changed: ${fields.join(', ')}).`;
  }

  return `${userContext} ${action} ${fallbackType}${targetName}.`;
}

/**
 * 📡 The Main Extraction Engine
 */
function getChangeAlerts() {
  const accountName = AdsApp.currentAccount().getName();
  
  const { startDate, endDate } = getSafeDateRange();
  console.log(`🕵️‍♂️ Snooping for changes between ${startDate} and ${endDate}...`);
    
  const query = `
    SELECT 
      campaign.name, 
      ad_group.name,
      change_event.change_date_time, 
      change_event.change_resource_name, 
      change_event.change_resource_type, 
      change_event.changed_fields, 
      change_event.client_type, 
      change_event.new_resource, 
      change_event.old_resource, 
      change_event.resource_change_operation, 
      change_event.user_email 
    FROM change_event 
    WHERE change_event.change_date_time >= '${startDate}' 
      AND change_event.change_date_time <= '${endDate}'
      AND change_event.user_email NOT IN ('${IGNORE_USERS.join("', '")}') 
    ORDER BY change_event.change_date_time DESC 
    LIMIT 9999
  `;
              
  let result;
  try {
    result = AdsApp.search(query);
  } catch (e) {
    console.error(`💥 Yikes! Issue retrieving results from the search API: ${e}`);
    return []; 
  } 
  
  const rawChanges = [];
  while (result.hasNext()) {
    const row = result.next();
    rawChanges.push({ row: row, resourceName: row.changeEvent?.changeResourceName || "" });
  }
  
  console.log(`🔍 Found ${rawChanges.length} raw events. Running the deep resolver to fetch human-readable names...`);
  
  const resolverMap = fetchDeepCriterionNames(rawChanges);

  const formatJSON = (obj) => {
    if (!obj) return "";
    if (typeof obj === 'object' && Object.keys(obj).length === 0) return "";
    try { return JSON.stringify(obj, null, 2); } catch (e) { return String(obj); }
  };
  
  const finalChanges = [];

  for (const raw of rawChanges) {
    try {
      const row = raw.row;
      const campaignName = row.campaign?.name ?? "";
      const adGroupName = row.adGroup?.name ?? "";
      const { changeEvent } = row;
      
      const operation = changeEvent.resourceChangeOperation || "";
      const resourceType = changeEvent.changeResourceType || "";
      const userEmail = changeEvent.userEmail || "";
      
      const changedFields = changeEvent.changedFields || {};
      const newRes = changeEvent.newResource || {};
      const oldRes = changeEvent.oldResource || {};

      // 🕒 FORMATTING: Clean the timestamp by trimming off the microseconds
      let rawDate = changeEvent.changeDateTime || "";
      let cleanDate = rawDate.includes('.') ? rawDate.split('.')[0] : rawDate;

      const resNameObj = resolverMap.criteriaMap[raw.resourceName];
      let criterionDisplayName = "";
      
      if (resNameObj && resNameObj.name) {
          criterionDisplayName = resNameObj.name;
      } else {
          const getVal = (obj, ...keys) => keys.reduce((a, c) => (a && a[c] !== undefined) ? a[c] : null, obj);
          const critId = getVal(newRes, 'campaignCriterion', 'criterionId') || 
                         getVal(oldRes, 'campaignCriterion', 'criterionId') ||
                         getVal(newRes, 'adGroupCriterion', 'criterionId') || 
                         getVal(oldRes, 'adGroupCriterion', 'criterionId');
                         
          if (critId && resolverMap.fallbackMap[critId]) {
              criterionDisplayName = resolverMap.fallbackMap[critId]; 
          } else if (critId) {
              criterionDisplayName = `criterion (ID: ${critId})`; 
          }
      }

      const plainEnglishDesc = generatePlainEnglishDescription(
        operation, resourceType, changedFields, userEmail, campaignName, adGroupName, newRes, oldRes, criterionDisplayName
      );

      finalChanges.push([
        cleanDate,                  // Cleaned 'YYYY-MM-DD HH:mm:ss'
        plainEnglishDesc,           // Column B
        accountName,
        userEmail,
        changeEvent.clientType,
        campaignName,
        resourceType,
        operation,
        changeEvent.changeResourceName, 
        formatJSON(changedFields),
        formatJSON(newRes),
        formatJSON(oldRes)
      ]);
      
    } catch (e) {
      console.error(`⚠️ Hiccup parsing a result row: ${e}`);
    }
  }
 
  return finalChanges;
}

// ============================================================================
// 📊 REPORTING & OUTPUT (Bespoke Overwrite Mode)
// ============================================================================

/**
 * Orchestrates the spreadsheet preparation and data injection.
 */
function reportResults(changes) {
  const sheet = prepareOutputSheet();
  if (sheet) {
    addOutputToSheet(changes, sheet);  
  }
}

/**
 * Hooks into the Google Sheet, automatically renames the spreadsheet to match 
 * the account (only if it is currently untitled), explicitly removes old filters, 
 * wipes old data, and sizes columns.
 */
function prepareOutputSheet() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  if (!spreadsheet) {
    console.error("❌ CRITICAL: Cannot open the designated reporting spreadsheet. Check your URL!");
    return null;
  }

  // --- AUTO-RENAME SPREADSHEET (Only if Untitled) ---
  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const expectedTitle = `${accountName} (${accountId}) Change History | WeDiscover`;
  
  if (spreadsheet.getName() === "Untitled spreadsheet") {
    spreadsheet.rename(expectedTitle);
    console.log(`📝 Auto-renamed the spreadsheet from "Untitled spreadsheet" to: "${expectedTitle}"`);
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    console.log(`🛠️ Sheet named '${SHEET_NAME}' not found. Constructing a fresh tab now.`);
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }  

  // --- FILTER CLEANUP ---
  // If a filter is active from a previous run, gracefully remove it before the wipe
  if (sheet.getFilter()) {
     sheet.getFilter().remove();
  }

  // Nuclear wipe of the sheet to ensure no residual data or formatting remains
  sheet.clear();

  const columnWidths = [
    160, // Change Date & Time
    350, // Plain English Description (Moved to Column B)
    150, // Account Name
    200, // User Email
    150, // Client Type
    200, // Campaign Name
    150, // Resource Type
    150, // Change Operation
    250, // Changed Resource Name
    250, // Changed Fields
    300, // New Resource
    300  // Old Resource
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  return sheet;
}

/**
 * Dumps the data into the sheet, applies universal word-wrap, formats the date, 
 * injects WeDiscover's row banding, and sets up a fresh filter view!
 */
function addOutputToSheet(output, sheet) {
  const headers = [
    'Change Date & Time', 
    'Plain English Description', 
    'Account Name', 
    'User Email', 
    'Client Type', 
    'Campaign Name', 
    'Resource Type', 
    'Change Operation',
    'Changed Resource Name',
    'Changed Fields', 
    'New Resource', 
    'Old Resource'
  ];
  
  const dataToWrite = [headers, ...output];
  const dataRange = sheet.getRange(1, 1, dataToWrite.length, headers.length);
  dataRange.setValues(dataToWrite);

  // Apply typography, alignment, and UNIVERSAL WORD WRAP to all cells
  dataRange.setFontFamily("Manrope");
  dataRange.setVerticalAlignment("middle");
  dataRange.setWrap(true);

  // --- STRICT DATE FORMATTING ---
  // Ensures Google Sheets interprets Column A as a pure Date/Time string
  const dateColumnRange = sheet.getRange(2, 1, output.length, 1);
  dateColumnRange.setNumberFormat('ddd dd mmm yyyy "at" hh:mm');

  // Apply beautiful alternating Row Banding
  const banding = dataRange.applyRowBanding();
  banding.setHeaderRowColor('#B12A32');   // WeDiscover Brand Red
  banding.setFirstRowColor('#FFFFFF');    // Clean White
  banding.setSecondRowColor('#FDF4F5');   // Very Subtle Red Tint

  // Make the header pop
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Freeze the top row so the header follows the user as they scroll
  sheet.setFrozenRows(1);

  // --- AUTO-FILTER ---
  // Construct a fresh, clean filter over the brand new data range
  dataRange.createFilter();

  console.log(`✨ Magic! Successfully transformed, filtered, and beamed ${output.length} rows into your spreadsheet.`);
}
