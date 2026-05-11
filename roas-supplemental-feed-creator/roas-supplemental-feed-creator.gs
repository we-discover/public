/**
 * Google Ads Script to extract product item IDs and ROAS,
 * then categorise them into a Google Sheet.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================
const SPREADSHEET_URL = 'INSERT SPREADSHEET URL HERE';
const SHEET_NAME = 'INSERT SHEET NAME HERE'; // e.g. 'Sheet1'

// ROAS thresholds for categorisation
const THRESHOLD_HIGH = 2.0;
const THRESHOLD_LOW = 1.0;

// Labels to apply based on the thresholds (Column C)
const LABEL_HIGH = 'high';
const LABEL_MED = 'med';
const LABEL_LOW = 'low';

// ============================================================================
// MAIN PROGRAMME
// ============================================================================
function main() {
  const sheet = getSheet(SPREADSHEET_URL, SHEET_NAME);
  const productData = fetchProductData();
  writeToSheet(sheet, productData);
}

/**
 * Retrieves the target Google Sheet.
 * @param {string} url The URL of the spreadsheet.
 * @param {string} name The name of the specific sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(url, name) {
  const spreadsheet = SpreadsheetApp.openByUrl(url);
  const sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    throw new Error(`Sheet '${name}' not found. Please verify the name is correct.`);
  }

  return sheet;
}

/**
 * Fetches product performance data for the last 30 days, calculates ROAS, and sorts by ID.
 * Uses a two-query method to catch products that had zero activity in the last 30 days,
 * with added safety checks for undefined/ghost products.
 * @returns {Array<Array<string|number>>} A 2D array representing rows for the sheet.
 */
function fetchProductData() {
  // 1. Fetch 30-day performance data and store it in a dictionary
  const thirtyDayQuery = `
    SELECT
      segments.product_item_id,
      metrics.conversions_value,
      metrics.cost_micros
    FROM shopping_performance_view
    WHERE segments.date DURING LAST_30_DAYS
  `;
  
  const thirtyDayReport = AdsApp.search(thirtyDayQuery);
  const statsMap = {};

  while (thirtyDayReport.hasNext()) {
    const row = thirtyDayReport.next();
    
    // SAFETY CHECK: Ensure the product ID actually exists before processing
    if (row.segments && row.segments.productItemId) {
      const itemId = row.segments.productItemId;
      statsMap[itemId] = {
        cost: row.metrics.costMicros / 1000000,
        value: row.metrics.conversionsValue || 0
      };
    }
  }

  // 2. Fetch ALL known product IDs (all-time) to build our master list
  const allTimeQuery = `
    SELECT segments.product_item_id
    FROM shopping_performance_view
  `;
  
  const allTimeReport = AdsApp.search(allTimeQuery);
  const allItemIds = new Set();
  
  while (allTimeReport.hasNext()) {
    const row = allTimeReport.next();
    
    // SAFETY CHECK: Prevents the "Cannot read properties of undefined" error
    if (row.segments && row.segments.productItemId) {
      allItemIds.add(row.segments.productItemId);
    }
  }

  // 3. Compile the final rows
  const rows = [];
  rows.push(['id', 'roas', 'custom_label_2']); // Headers

  // Sort the IDs alphabetically/numerically for a clean sheet output
  const sortedIds = Array.from(allItemIds).sort();

  for (const itemId of sortedIds) {
    let cost = 0;
    let conversionValue = 0;

    // If the product had activity in the last 30 days, grab those stats
    if (statsMap[itemId]) {
      cost = statsMap[itemId].cost;
      conversionValue = statsMap[itemId].value;
    }

    // Calculate Return on Ad Spend (ROAS)
    const rawRoas = cost > 0 ? (conversionValue / cost) : 0;
    
    // Format to 2 decimal places and ensure it remains a number type
    const roas = Number(rawRoas.toFixed(2));

    // Categorise the product based on ROAS performance
    const label = determineLabel(roas);

    // Push the compiled row into our data array
    rows.push([itemId, roas, label]);
  }

  return rows;
}

/**
 * Determines the correct label based on ROAS thresholds.
 * @param {number} roas The calculated ROAS.
 * @returns {string} The assigned categorical label.
 */
function determineLabel(roas) {
  if (roas > THRESHOLD_HIGH) {
    return LABEL_HIGH;
  }

  if (roas < THRESHOLD_LOW) {
    return LABEL_LOW;
  }

  // If it falls between the thresholds (inclusive)
  return LABEL_MED;
}

/**
 * Writes the processed data to the Google Sheet and applies formatting.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The target sheet.
 * @param {Array<Array<string|number>>} data The 2D array of data to output.
 */
function writeToSheet(sheet, data) {
  // Completely clear contents AND formatting to ensure a clean slate
  sheet.clear();

  if (data.length > 1) { 
    // Write the new data starting from row 1, column 1
    const range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    
    // Make the header row (row 1, column 1 to the end of the columns) bold
    sheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold");
    
    Logger.log(`Successfully wrote ${data.length - 1} products to the sheet in ascending order.`);
  } else {
    Logger.log('No product data found to process.');
  }
}
