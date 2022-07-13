/************* USER CONFIGS *********************************************************************************************/

// URL of spreadsheet
const SPREADSHEET_URL = "XXXXXX";

// Number of days of data to include in export
const REPORTING_WINDOW = 180;

/************************************************************************************************************************/
/************* DO NOT EDIT BELOW THIS LINE ******************************************************************************/
/************************************************************************************************************************/

/************* MAIN FUNCTION ********************************************************************************************/

function main() {
  const reportStart = getDateXDaysAgo(REPORTING_WINDOW);
  const yesterday = getDateXDaysAgo(1)
  const queriesConfig = [
    {
      name: "BASE_METRICS_PERFORMANCE_QUERY",
      queryTemplate: BASE_METRICS_PERFORMANCE_QUERY_TEMPLATE,
      sheetName: "Google Ads Import: Google Ads Campaign Stats"
    },
    {
      name: "CONVERSION_PERFORMANCE_QUERY",
      queryTemplate: CONVERSION_PERFORMANCE_QUERY_TEMPLATE,
      sheetName: "Google Ads Import: Google Ads Campaign Conv. Stats"
    },
  ];

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  Logger.log(`Loaded config and connected to spreadsheet ${spreadsheet.getName()}.`);

  Logger.log("Iterating through queries...");

  for (let i = 0; i < queriesConfig.length; i++) {    
    const config = queriesConfig[i];
    Logger.log(`Processing query ${config.name}...`);

    const query = `${config.queryTemplate} '${reportStart}' AND '${yesterday}'`;

    const sheet = getOrCreateSheet(spreadsheet, config.sheetName);
    getReportToSheet(query, sheet);
    sheet.hideSheet();

    Logger.log(`Exported ${config.name}.`);
  }

  Logger.log("Finished iterating.");
  Logger.log("Terminating.");

}

/************* UTILITY FUNCTIONS ******************************************************************************************/

// Pull a GAQL report and export it to a tab in a Google Sheet
function getReportToSheet(query, sheet) {
  const report = AdsApp.report(query);

  sheet.clear();
  report.exportToSheet(sheet);
}

// Get a sheet by name, or create and colour-code the sheet if it does not exist
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet !== null) {
    Logger.log(`Found sheet: ${sheetName}.`);

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

/************* GAQL QUERIES *********************************************************************************************/

// NB: The WHERE clauses must be completed with dates before using

const BASE_METRICS_PERFORMANCE_QUERY_TEMPLATE = `
  SELECT 
    segments.date, 
    customer.id,
    customer.descriptive_name,
    campaign.id,
    campaign.name, 
    metrics.cost_micros, 
    metrics.search_impression_share,
    metrics.impressions, 
    metrics.clicks
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
