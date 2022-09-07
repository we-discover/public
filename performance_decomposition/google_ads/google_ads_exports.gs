/*
    Name:        WeDiscover - Data Export for Performance Decomposition, Google Ads Script
    Description: A script to build export data for use with WeDiscover's
                 Performance Decomposition tool
    License:     https://github.com/we-discover/public/blob/master/LICENSE
    Version:     1.0.0
    Released:    2022-09-08
    Contact:     scripts@we-discover.com
*/


/************* GENERAL CONFIGS **************************************************************************************************************/

// URL of spreadsheet
const SPREADSHEET_URL = "XXXX";

// Number of days of data to include in export
const REPORTING_WINDOW = 180;

/************* MCC CONFIGS ******************************************************************************************************************/

// To filter which child accounts to pull data from, enter them below in ONE of the below
// To only pull data from specific accounts, add them to ACCOUNTS_TO_INCLUDE
// To pull data from all accounts except specific ones, add them to ACCOUNTS_TO_EXCLUDE
// Lists should be comma separated account IDs with dashses, enclosed in quotes
// E.g. ['123-456-7890', '345-6789-0123']

const ACCOUNTS_TO_INCLUDE = [];
const ACCOUNTS_TO_EXCLUDE = [];

/********************************************************************************************************************************************/
/************* DO NOT EDIT BELOW THIS LINE **************************************************************************************************/
/********************************************************************************************************************************************/

/************* SCRIPT ENTRY POINT ***********************************************************************************************************/

function main() {
  const isMcc = typeof AdsManagerApp == "undefined" ? false : true;
  
  const reportStart = getDateXDaysAgo(REPORTING_WINDOW);
  const yesterday = getDateXDaysAgo(1);
  const queriesConfig = [
    {
      name: "BASE_METRICS_PERFORMANCE_QUERY",
      query: `${BASE_METRICS_PERFORMANCE_QUERY_TEMPLATE} '${reportStart}' AND '${yesterday}'`,
      sheetName: "Google Ads Import: Google Ads Campaign Stats"
    },
    {
      name: "CONVERSION_PERFORMANCE_QUERY",
      query: `${CONVERSION_PERFORMANCE_QUERY_TEMPLATE} '${reportStart}' AND '${yesterday}'`,
      sheetName: "Google Ads Import: Google Ads Campaign Conv. Stats"
    },
  ];

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  Logger.log(`Loaded config and connected to spreadsheet ${spreadsheet.getName()}.`);

  const reportsBySheetName = initialiseReports(queriesConfig);
  
  if (isMcc) {
    const accountIterator = getAccountsToReportOn(ACCOUNTS_TO_INCLUDE, ACCOUNTS_TO_EXCLUDE);
    while (accountIterator.hasNext()) {
      let account = accountIterator.next();
      AdsManagerApp.select(account);
      accountMain(account, queriesConfig, reportsBySheetName);
    }

    Logger.log('Finished iterating.')
  }

  if (!isMcc) {
    accountMain(AdsApp.currentAccount(), queriesConfig, reportsBySheetName);
  }

  Logger.log('Exporting data to Google Sheet...')
  exportToSheet(reportsBySheetName, spreadsheet);
  Logger.log('Terminating.');

  return null;
}

/************* UTILITY FUNCTIONS ************************************************************************************************************/

// Run each query in a given account and export the results to a G Sheet
function accountMain(account, queriesConfig, reportsBySheetName) {
  Logger.log(`* Iterating through queries in account: ${account.getName()} (${account.getCustomerId()})...`);

  for (let i = 0; i < queriesConfig.length; i++) {    
    const config = queriesConfig[i];
    Logger.log(`  * Processing query ${config.name}...`);

    const query = config.query;
    const reportRows = AdsApp.report(query).rows();
    const headers = reportsBySheetName[config.sheetName]['headerOrder'];
    
    while (reportRows.hasNext()) {
      let row = reportRows.next();

      // Store query results by sheet and add each metric under its header
      for (let i = 0; i < headers.length; i++) {
         const header = headers[i];
         reportsBySheetName[config.sheetName][header].push(row[header]);
      }
    }
  }
  
  return null;
}

// Export reports to sheet from objects of reports ordered by sheet name
function exportToSheet(reportsBySheetName, spreadsheet) {
      
  for (let sheetName in reportsBySheetName) {
      const sheet = getOrCreateSheet(spreadsheet, sheetName);
      sheet.clear();
      const reports = reportsBySheetName[sheetName];

      // Write to sheet column by column due to format of reportsBySheetName object
      for (let i = 0; i < reports['headerOrder'].length; i++) {
        const currentHeader = reports['headerOrder'][i];
        let column = [[currentHeader]];
        const columnValues = reports[currentHeader].map(el => [el]);                // Store each value within an Array so it can be output to G Sheet
        column = column.concat(columnValues);
        sheet.getRange(1, i + 1, column.length, 1).setValues(column);
      }

      sheet.hideSheet();
  }
  
  return null;
}

// Get a sheet by name, or create and colour-code the sheet if it does not exist
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet !== null) {
    Logger.log(`  * Found sheet: ${sheetName}.`);

    return sheet;
  }

  else if (sheet === null) {
    Logger.log(`Creating sheet: ${sheetName}.`);

    sheet = spreadsheet.insertSheet(sheetName)
    sheet.setTabColor('red');

    return sheet;
  }
}

// Return a date of the form YYYY-MM-DD for a given number of days ago
function getDateXDaysAgo(lookbackWindow) {
  const date = new Date();
  const dateXDaysAgo = new Date(date.getTime() - 1000 * 60 * 60 * 24 * lookbackWindow);

  return Utilities.formatDate(dateXDaysAgo, AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");
}

// Validate MCC configs and return iterator of accounts to report on
function getAccountsToReportOn(accountsToInclude, accountsToExclude) {
  if (accountsToInclude.length > 0 && accountsToExclude.length > 0) {
    throw new Error("Only one of ACCOUNTS_TO_INCLUDE, ACCOUNTS_TO_EXCLUDE should be filled. The other must be empty.");
  }

  else if (accountsToInclude.length === 0 && accountsToExclude.length === 0) {
    return AdsManagerApp.accounts().get();
  }

  else if (accountsToInclude.length > 0 && accountsToExclude.length === 0) {
    validateAccountLists(accountsToInclude, 'ACCOUNTS_TO_INCLUDE');
    return AdsManagerApp.accounts().withIds(accountsToInclude).get();
  }

  else if (accountsToInclude.length === 0 && accountsToExclude.length > 0) {
    validateAccountLists(accountsToExclude, 'ACCOUNTS_TO_EXCLUDE');
    const invertedIds = changeAccountExclusionToInclusion(accountsToExclude);

    return AdsManagerApp.accounts().withIds(invertedIds).get();
  }

}

// Validate that account IDs in an array are of the form XXX-XXX-XXXX
function validateAccountLists(ids, listName) {
  const checkIndividualElement = (el) => /^\d{3}\-\d{3}\-\d{4}$/.test(el);
  const allCorrectForm = ids.every(checkIndividualElement);

  if (allCorrectForm) {
    Logger.log(`Account IDs in list ${listName} of the correct form.`);
  }

  else if (!allCorrectForm) {
    const malformedIds = ids.filter((el) => !checkIndividualElement(el));
    throw new Error(`The following account IDs in the list ${listName} are incorrectly formatted: ${malformedIds}. \
Please ensure IDs are entered as a comma-separated list of the form XXX-XXX-XXXX`);
  }


}

// Get the column headers from a raw GAQL query as an array
function getHeaders(query) {
  const formattedQuery = query.replace(/\s+/g, " ").replace(/^ /, "").replace(/ $/, "");
  const headerString = /^SELECT (.*) FROM.*$/g.exec(formattedQuery)[1];
  
  return headerString.split(', ');
}

// Initialise form of report as {sheetName1: {header1: [], header2: [], ...}, ...}
function initialiseReports(queriesConfig) {
  const reportsBySheetName = {};
  
  for (let i = 0; i < queriesConfig.length; i++) {
    const config = queriesConfig[i];
    const queryHeaders = getHeaders(config.query);
    reportsBySheetName[config.sheetName] = {};
    reportsBySheetName[config.sheetName]['headerOrder'] = queryHeaders;
    queryHeaders.forEach(h => reportsBySheetName[config.sheetName][h] = []);
  }
  
  return reportsBySheetName;
}

// Get IDs of accounts which are in MCC but not in accountsToExclude
// For use with .withIds() function when selecting child accounts
function changeAccountExclusionToInclusion(accountsToExclude) {
  const accountsToInclude = [];

  const accountsIterator = AdsManagerApp.accounts().get();

  while(accountsIterator.hasNext()) {
    let customerId = accountsIterator.next().getCustomerId();

    if (accountsToExclude.indexOf(customerId) === -1) {
      accountsToInclude.push(customerId);
    }
  }

  return accountsToInclude;
}

/************* GAQL QUERIES *****************************************************************************************************************/

// NB: The WHERE clauses must be completed with dates before using

const BASE_METRICS_PERFORMANCE_QUERY_TEMPLATE = `
  SELECT 
    segments.date, 
    customer.id,
    customer.descriptive_name,
    customer.currency_code,
    campaign.id,
    campaign.name, 
    metrics.cost_micros, 
    metrics.search_impression_share,
    metrics.impressions, 
    metrics.clicks,
    metrics.conversions,
    metrics.conversions_value
  FROM 
    campaign 
  WHERE 
    segments.date BETWEEN 
`;

const CONVERSION_PERFORMANCE_QUERY_TEMPLATE = `
  SELECT 
    segments.date, 
    customer.id,
    customer.descriptive_name,
    customer.currency_code,
    campaign.id,
    campaign.name, 
    segments.conversion_action_name,
    metrics.all_conversions, 
    metrics.all_conversions_value
  FROM 
    campaign 
  WHERE 
    segments.date BETWEEN 
`;
