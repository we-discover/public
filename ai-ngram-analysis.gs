/*
 *  Name:        WeDiscover - Search Query N-Gram Analysis, Google Ads Script
 *
 *  Description: Mines the search query report for words and phrases (n-grams),
 *               aggregates click, impression, cost and conversion performance
 *               per term, and writes segmented results (account, campaign and
 *               ad group level) to a Google Sheet. Optimised for high-volume
 *               accounts (1M+ clicks/month, 50M+ impressions/month).
 *
 *  License:     https://github.com/we-discover/public/blob/master/LICENSE
 *  Version:     3.10.5
 *  Released:    2026-05-06
 *  Contact:     scripts@we-discover.com
 *
 *  Credits:     Original n-gram concept and structure by Brainlabs Digital
 *               (https://github.com/Brainlabs-Digital/Google-Ads-Scripts).
 *               GAQL migration by Nils Rooijmans (2022, 2025).
 *               Shared set access fix by Arjan Schoorl / Flowboost (2025).
 *               Rewritten for ES6, GAQL and scale by WeDiscover (2026).
 *
 */


// =============================================================================
// CONFIGURATION
// This is the only section you need to edit. Everything below it runs itself.
// =============================================================================
const CONFIG = {

  // -- Date range --------------------------------------------------------------
  // Enter the start and end dates for the data you want to analyse.
  // Dates go in as DD/MM/YYYY because that is the correct way to write
  // a date and we will not be taking questions. Yes, Americans put the
  // month first (June 23rd becomes 06/23). Nobody knows why. We've asked.
  startDate: '01/01/2026',   // dd/mm/yyyy -- change this to your start date
  endDate:   '01/02/2026',   // dd/mm/yyyy -- change this to your end date

  // -- Currency ----------------------------------------------------------------
  // Type your currency symbol directly. £ for GBP, $ for USD, € for EUR.
  // Whatever you paste in here will appear in the Cost columns in the sheet.
  currencySymbol: '£',

  // -- Campaign filter ---------------------------------------------------------
  // If you only want to look at certain campaigns, type part of their name here.
  // For example, typing 'Brand' will only include campaigns with 'Brand' in the
  // name. Leave both fields empty to include all campaigns.
  campaignNameContains:       '',  // only include campaigns whose name contains this
  campaignNameDoesNotContain: '',  // exclude campaigns whose name contains this

  // -- Paused campaigns and ad groups ------------------------------------------
  // Set these to true to skip anything that is currently paused.
  // Set to false if you want to include paused campaigns or ad groups too.
  ignorePausedCampaigns: true,
  ignorePausedAdGroups:  true,

  // -- Negative keywords -------------------------------------------------------
  // Set to true to filter out search queries that are already blocked by your
  // negative keywords. This gives you a cleaner picture of what is actually
  // reaching your ads. Recommended: true.
  checkNegatives: true,

  // -- Spreadsheet -------------------------------------------------------------
  // Paste the full URL of the Google Sheet you want the results written to.
  // The sheet must already exist and the account running this script must
  // have edit access to it.
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/INSERT-SPREADSHEET-URL-HERE',

  // -- N-gram length -----------------------------------------------------------
  // An n-gram is a sequence of words. A 1-gram is a single word like "shoes".
  // A 2-gram is two words like "running shoes". A 3-gram is "womens running shoes".
  // minNGramLength: the shortest phrase to look for (1 is a good starting point).
  // maxNGramLength: the longest phrase to look for.
  // On high-traffic accounts, keep maxNGramLength at 4 or below. Going higher
  // uses a lot more memory and makes the script slower.
  minNGramLength: 1,
  maxNGramLength: 4,

  // -- Clear the spreadsheet on each run ---------------------------------------
  // Set to true to wipe the sheet clean before writing new results.
  // Set to false to keep old results and add new ones on top (not recommended).
  clearSpreadsheet: true,

  // -- Thresholds --------------------------------------------------------------
  // Any phrase that does not meet all of these minimums will be left out of
  // the results. This stops the sheet from filling up with hundreds of phrases
  // that got one click and are not worth your attention.
  // On big accounts, a clicks threshold of 50 or higher is sensible.
  thresholds: {
    queryCount:   0,   // minimum number of search queries containing the phrase
    impressions:  0,   // minimum impressions
    clicks:       50,  // minimum clicks -- raise this on high-volume accounts
    cost:         0,   // minimum spend
    conversions:  0,   // minimum conversions
  },

  // -- Time limit --------------------------------------------------------------
  // Google Ads Scripts stop automatically after 30 minutes. This setting tells
  // the script to stop early and save whatever it has found so far, rather than
  // crashing with no output. 25 minutes is a safe value.
  maxExecutionMinutes: 25,
};
// =============================================================================
// END OF CONFIGURATION -- do not edit anything below this line
// =============================================================================


// =============================================================================
// AI CONFIGURATION (OPTIONAL)
//
// This section controls the optional AI summary feature. If you just want
// the data sheets and no AI analysis, leave apiKey blank and ignore this.
//
// ---- HOW TO GET A FREE API KEY (takes about two minutes) -------------------
//   1. Go to https://aistudio.google.com in any browser.
//   2. Sign in with any Google account -- the same one you use for Ads is fine.
//   3. Click "Get API key" in the left-hand navigation panel.
//   4. Click "Create API key", select any project, then click "Create".
//   5. Copy the key that appears and paste it into the apiKey field below.
//
//   The free tier gives you 1,500 requests per day at no cost. No credit card
//   is required. For reference, even the most data-heavy run of this script
//   in 'full' mode (see below) uses around 40-50 requests, so the free tier
//   is more than sufficient for regular monthly use.
//
// ---- WHAT THE AI DOES -------------------------------------------------------
//   The script sends your n-gram data to Google's Gemini Flash model -- a fast,
//   capable AI that is specifically designed for analytical tasks like this. It
//   reads the phrase data, understands which phrases are driving conversions and
//   which are wasting budget, and writes a plain-English summary with specific
//   recommendations.
//
//   Your data is sent to Google's API servers to generate the summary, in the
//   same way that you would paste data into a chat window. It is not used to
//   train any AI model. See https://ai.google.dev/terms for Google's full terms.
//
// ---- TWO MODES: 'account' vs 'full' ----------------------------------------
//   mode: 'account'
//     One API call. Sends the top N phrases by clicks from each of the four
//     account-level sheets (words, 2-grams, 3-grams, 4-grams). Fast, cheap,
//     good for a quick overall read of the account. This is the default.
//
//   mode: 'full'
//     One API call per campaign, using ALL phrase data at both campaign and
//     ad group level. Produces specific, actionable recommendations for every
//     campaign and its ad groups. The number of API calls depends on how many
//     campaigns you have and how large they are.
//
//     For large campaigns (lots of ad groups or very high search volume), the
//     script automatically splits the data into chunks that each fit within the
//     token budget, sends them separately, and then combines the responses into
//     a single campaign section. You do not need to configure anything for this
//     -- it is entirely automatic and works on any account.
//
//     On a typical account with 10-30 campaigns this takes 2-3 minutes of API
//     time on top of the usual mining time, and costs well under $0.20 even on
//     the paid tier. On the free tier it costs nothing.
// =============================================================================
const AI_CONFIG = {

  // -- API key -----------------------------------------------------------------
  // Paste your Google AI Studio key here. Leave blank ('') to skip the AI
  // summary entirely. The rest of the script runs exactly as normal either way.
  apiKey: '',

  // -- Mode --------------------------------------------------------------------
  // 'account' -- one call, top N phrases per account-level sheet. Fast.
  // 'full'    -- one call per campaign, all data at campaign and ad group level.
  mode: 'account',

  // -- Token budget (full mode only) -------------------------------------------
  // In 'full' mode, each API call is limited to this many tokens of input data.
  // The script measures each campaign's data size automatically and splits it
  // into chunks if needed -- you do not need to do anything manually.
  //
  // 60,000 is a good balance between giving the AI enough context to write
  // specific recommendations and keeping each call fast and reliable. There is
  // no strong reason to change this unless you are hitting timeout issues, in
  // which case reducing it to 40,000 will produce more but smaller calls.
  tokenBudget: 60000,

  // -- Number of phrases per sheet (account mode only) -------------------------
  // How many of the top phrases by clicks to include from each account-level
  // sheet. 50 is a good default. In 'full' mode this setting is ignored --
  // all rows are sent.
  topN: 50,

  // -- Output sheet name -------------------------------------------------------
  // The AI summary is written to a sheet with this name, placed at the front
  // of the workbook. In 'full' mode this sheet contains the account summary
  // followed by a section for each campaign.
  sheetName: 'AI Summary',

  // -- Account summary prompt --------------------------------------------------
  // The instruction sent to the AI for the account-level summary. The phrase
  // data is appended automatically -- you only need to describe what you want
  // the AI to do with it. You can rewrite this freely. The only rule is to
  // keep it under a few hundred words; longer instructions do not improve the
  // output and just waste your token allowance.
  accountPrompt: [
    'You are a senior Google Ads strategist at WeDiscover in London.',
    'You are sharp, no-nonsense, and have a dry sense of humour.',
    'You are reviewing search query n-gram data to find clean, incremental growth.',
    'Below is data from a Google Ads account at the word, 2-gram, 3-gram and 4-gram level.',
    'Columns are: phrase, clicks, cost, conversions, conversion_value, roas.',
    '',
    '1. DATA INTEGRITY RULES -- apply these before any analysis:',
    '- Ignore phrases that are fragments of a brand name (e.g. function words',
    '  like "on the", "for a", "in the" that only appear because they are part',
    '  of a branded query). Focus on phrases with clear, standalone search intent.',
    '- Do not suggest phrases as opportunities if they lack clear commercial intent.',
    '- When citing any n-gram phrase in your response, ALWAYS wrap it in single',
    '  quotes. Examples: \'gifts for her\', \'60th birthday gifts\', \'for him\'.',
    '  Every single phrase you mention must have single quotes around it.',
    '  This is the most important formatting rule.',
    '',
    '2. FORMAT RULES -- follow these exactly:',
    '- Start each section with a numbered heading on its own line, like: 1. TOP THEMES',
    '- Write bullet points starting with "- " (a hyphen then a space)',
    '- Use plain sentences for any introductory or closing text',
    '- Do not use markdown, asterisks, hash symbols, or any other markup',
    '- Do not use ALL CAPS except in section headings',
    '- Use British English throughout (e.g. "optimise" not "optimize")',
    '- Write at a Grade 8 reading level -- clear, direct, no jargon',
    '- Make it engaging and slightly entertaining where the data supports it',
    '',
    '3. ANALYSIS SECTIONS:',
    '',
    '1. TOP THEMES',
    'Identify 3 high-intent search categories (excluding core brand traffic).',
    'Name the supporting phrases and their ROAS.',
    '',
    '2. INCREMENTAL OPPORTUNITIES',
    'Name specific non-brand phrases with strong ROAS and meaningful volume.',
    'For each one, say whether it is best suited to Exact match (to protect',
    'high-margin volume and isolate performance) or Phrase match (to capture',
    'variations and scale). Be specific -- name the phrases and the numbers.',
    '',
    '3. WASTAGE',
    'Identify non-brand phrases with high spend and ROAS below 0.8.',
    'Name them, state their ROAS, and explain briefly why they are a problem.',
    '',
    '4. STRATEGIC NOTE',
    'One sharp observation on how the account is positioned -- for example,',
    'how brand vs. generic performance is shifting, or a structural pattern',
    'worth the account manager knowing about.',
    '',
    'Keep the whole summary under 500 words.',
    'Final reminder: every phrase you mention must be in single quotes.',
    'Write ROAS in capitals throughout -- never write "roas" in lowercase.',
  ].join('\n'),

  // -- Campaign prompt (full mode only) ----------------------------------------
  // The instruction sent to the AI for each per-campaign analysis. Three
  // placeholders are filled in automatically before the call is made:
  //   {CAMPAIGN}     -- the campaign name
  //   {CHUNK_NOTE}   -- blank for single-chunk campaigns; for large campaigns
  //                     that are split across multiple calls, this becomes
  //                     something like "(part 2 of 4)" so the AI knows it is
  //                     seeing a partial view and should not draw final
  //                     conclusions from this chunk alone.
  //   {AVG_ROAS}     -- the average ROAS for this campaign across all its data,
  //                     giving the AI a benchmark to compare ad groups against.
  //
  // The phrase data (campaign rows and ad group rows) is appended automatically.
  // Campaign columns: phrase, clicks, cost, conversions, conv_value, roas.
  // Ad group columns: ad_group, phrase, clicks, cost, conversions, conv_value, roas.
  campaignPrompt: [
    'You are a senior Google Ads strategist at WeDiscover in London.',
    'You are sharp, no-nonsense, and have a dry sense of humour.',
    'You are reviewing n-gram data for the campaign "{CAMPAIGN}" {CHUNK_NOTE}.',
    'The average ROAS for this campaign is {AVG_ROAS}.',
    'Columns for campaign-level data: phrase, clicks, cost, conversions, conv_value, roas.',
    'Columns for ad group data: ad_group, phrase, clicks, cost, conversions, conv_value, roas.',
    '',
    '1. DATA INTEGRITY RULES -- apply these before any analysis:',
    '- Ignore phrases that are fragments of a brand name or function words with',
    '  no standalone commercial intent (e.g. "on the", "for a", "in the").',
    '- When citing any n-gram phrase in your response, ALWAYS wrap it in single',
    '  quotes. Examples: \'birthday gifts for her\', \'for him\', \'sterling silver\'.',
    '  Every single phrase you mention must have single quotes around it.',
    '  This is the most important formatting rule.',
    '- When suggesting opportunities, distinguish clearly between:',
    '  Exact match -- use for high-margin phrases where you want to isolate and',
    '  protect performance (specific, high-intent phrases with strong ROAS).',
    '  Phrase match -- use for modifiers and intent signals where you want to',
    '  capture variations and scale (good ROAS but benefits from broader reach).',
    '',
    '2. FORMAT RULES -- follow these exactly:',
    '- Start each section with a numbered heading on its own line, like: 1. CAMPAIGN OVERVIEW',
    '- Write bullet points starting with "- " (a hyphen then a space)',
    '- Use plain sentences for any introductory or closing text',
    '- Do not use markdown, asterisks, hash symbols, or any other markup',
    '- Do not use ALL CAPS except in section headings',
    '- Use British English throughout (e.g. "optimise" not "optimize")',
    '- Write at a Grade 8 reading level -- clear, direct, no jargon',
    '- Make it engaging and slightly entertaining where the data supports it',
    '',
    '3. ANALYSIS SECTIONS:',
    '',
    '1. CAMPAIGN OVERVIEW',
    'What is this campaign about and how is it performing overall?',
    'Reference the average ROAS and call out any standout phrases.',
    '',
    '2. AD GROUP PERFORMANCE',
    'Which ad groups are punching above their weight, and which are dragging',
    'the campaign down? Compare against the average ROAS of {AVG_ROAS}.',
    'Name the ad groups and the specific phrases driving performance.',
    '',
    '3. WASTAGE',
    'Identify non-brand phrases with high spend and ROAS below 0.8.',
    'Name them with their ROAS and explain why they are a problem.',
    '',
    '4. ACTIONS',
    'Two or three specific things a PPC manager could act on this week.',
    'For each action, say whether it involves Exact match or Phrase match,',
    'and name the actual phrases involved.',
    '',
    'Keep the whole analysis under 400 words.',
    'Final reminder: every phrase you mention must be in single quotes.',
    'Write ROAS in capitals throughout -- never write "roas" in lowercase.',
  ].join('\n'),

};
// =============================================================================
// END OF AI CONFIGURATION
// =============================================================================


// -----------------------------------------------------------------------------
// Constants
// These are derived from CONFIG and define the columns and formatting used
// throughout the script. Changing these will break things.
// -----------------------------------------------------------------------------
const STAT_COLS   = ['clicks', 'impressions', 'cost', 'conversions', 'conversionsValue'];
const STAT_LABELS = ['Clicks', 'Impressions', 'Cost', 'Conversions', 'Conv. Value'];

// Each entry is [column label, numerator field, denominator field].
// These are calculated from the raw stats when building each output row.
const CALC_STATS = [
  ['CTR',            'clicks',           'impressions'],
  ['CPC',            'cost',             'clicks'],
  ['Conv. Rate',     'conversions',      'clicks'],
  ['Cost / Conv.',   'cost',             'conversions'],
  ['Conv. Val/Cost', 'conversionsValue', 'cost'],
];

const CURRENCY_FMT = CONFIG.currencySymbol + '#,##0.00';

// Number formats for each column, matching STAT_COLS then CALC_STATS order.
const COL_FORMATS = ['#,##0', '#,##0', CURRENCY_FMT, '#,##0.00', CURRENCY_FMT,
                     '0.00%', CURRENCY_FMT, '0.00%', CURRENCY_FMT, '0.00%'];

// Recorded at startup so we can track elapsed time throughout the run.
const START_TS = Date.now();

// Column indices (1-based) of the five raw stat columns within a data row,
// relative to the start of the row. These are used by the deduplication pass
// to know which cells to sum and which cells to ignore (label columns and
// calculated columns). The label column count varies by sheet type, so the
// dedup pass calculates the offset at runtime from the sheet definition.
// The order must match STAT_COLS exactly: clicks, impressions, cost,
// conversions, conversionsValue.
const RAW_STAT_COUNT = STAT_COLS.length;   // 5
const CALC_STAT_COUNT = CALC_STATS.length; // 5


// -----------------------------------------------------------------------------
// main()
// Entry point. Orchestrates the full pipeline:
//   1. Open the spreadsheet and fetch campaign IDs.
//   2. Load negative keywords into Maps for fast lookup.
//   3. Initialise all output sheets.
//   4. Stream search query rows, accumulate n-gram stats, flush to sheets.
//   5. Finalise: deduplicate, sort, format, add filters.
// -----------------------------------------------------------------------------
function main() {
  validateConfig();

  const ss          = openSpreadsheet();
  const gaqlRange   = buildDateRange();
  const campaignIds = fetchActiveCampaignIds(gaqlRange);

  if (campaignIds.length === 0) {
    Logger.log('[ WeDiscover | N-Gram ] No active campaigns found with impressions in this date range. Nothing to mine. Exiting.');
    return;
  }
  Logger.log('============================================================');
  Logger.log('  WeDiscover | Search Query N-Gram Analysis');
  Logger.log('  Buckle up. We are mining ' + campaignIds.length + ' campaign(s) for gold.');
  Logger.log('============================================================');

  // Load negative keywords into Maps keyed by ad group ID and campaign ID.
  // Using Maps gives O(1) lookup per query row rather than looping through arrays.
  const { negsByAdGroup, negsByCampaign } = CONFIG.checkNegatives
    ? fetchAllNegatives(campaignIds, gaqlRange)
    : { negsByAdGroup: new Map(), negsByCampaign: new Map() };

  Logger.log('[ 1/4 ] Negative keywords mapped. The blocklist is locked and loaded.');

  // Initialise sheets before processing starts. If the script crashes halfway
  // through, any data already flushed to the sheet will still be there.
  const filterText = buildFilterText() + ' [processing...]';
  const sheets     = initialiseSheets(ss, filterText);
  Logger.log('[ 2/4 ] Sheets initialised. Your spreadsheet is prepped and ready.');

  // N-gram stats are accumulated in flat Maps using compound string keys.
  // Every FLUSH_EVERY rows, the Maps are written to the sheet and then cleared
  // so memory stays flat regardless of how many rows the report contains.
  //
  // The trade-off: flushing separate batches means the same phrase can appear
  // as multiple rows across batches. finaliseSheets() handles this with a
  // deduplication pass that reads the sheet back, merges rows by key, and
  // rewrites clean data. See deduplicateSheet() for details.
  let ngramMaps    = buildEmptyNgramMaps();
  let wordCountMap = new Map();

  const FLUSH_EVERY = 100000;
  let rowsProcessed = 0;
  let rowsSkipped   = 0;
  let flushCount    = 0;
  let stoppedEarly  = false;

  const queryIterator = fetchSearchQueryRows(campaignIds, gaqlRange);

  while (queryIterator.hasNext()) {
    const row = parseQueryRow(queryIterator.next());

    if (CONFIG.checkNegatives && isExcludedByNegative(row, negsByAdGroup, negsByCampaign)) {
      rowsSkipped++;
      continue;
    }

    accumulateNgrams(row, ngramMaps, wordCountMap);
    rowsProcessed++;

    // Every FLUSH_EVERY rows, write the current Maps to the sheet and reset them.
    // The time check runs after the flush because flushing itself takes time.
    if (rowsProcessed % FLUSH_EVERY === 0) {
      const elapsed = (Date.now() - START_TS) / 60000;
      Logger.log('[ 3/4 ] ' + rowsProcessed.toLocaleString() + ' rows processed | ' + elapsed.toFixed(1) + ' min elapsed | Writing batch to sheet...');

      flushToSheets(sheets, ngramMaps, wordCountMap);
      flushCount++;
      ngramMaps    = buildEmptyNgramMaps();
      wordCountMap = new Map();

      Logger.log('         Batch ' + flushCount + ' committed to sheet. Memory cleared. Back to mining...');

      if ((Date.now() - START_TS) / 60000 > CONFIG.maxExecutionMinutes) {
        Logger.log(`[ HEADS UP ] ${CONFIG.maxExecutionMinutes}-minute safety limit reached. Flushing what we have and signing off cleanly.`);
        stoppedEarly = true;
        break;
      }
    }
  }

  // Write whatever is left in the Maps after the loop finishes.
  Logger.log('[ 3/4 ] Final batch -- writing remaining rows to sheet...');
  flushToSheets(sheets, ngramMaps, wordCountMap);

  Logger.log('         Done mining! ' + rowsProcessed.toLocaleString() + ' rows processed, ' + rowsSkipped.toLocaleString() + ' skipped by negatives.');

  const finalFilterText = buildFilterText() + (stoppedEarly ? ' [PARTIAL -- stopped early]' : '');
  finaliseSheets(sheets, finalFilterText);
  deleteEmptySheets(ss);

  // Generate the AI summary sheet if an API key has been supplied.
  // This runs after everything else so a failure here cannot affect the main output.
  if (AI_CONFIG.apiKey) {
    var aiMode = (AI_CONFIG.mode === 'full') ? 'full' : 'account';
    Logger.log('[ AI ] API key found. Mode: ' + aiMode + '. Generating summary...');
    try {
      generateAiSummary(ss, aiMode);
      Logger.log('[ AI ] Summary written successfully.');
    } catch (e) {
      Logger.log('[ AI ] Summary failed (the rest of the data is fine): ' + e);
    }
  }

  Logger.log('[ 4/4 ] Formatting and sorting complete. Your n-gram data is ready to explore.');
  Logger.log('============================================================');
  Logger.log('  WeDiscover | Mission complete in ' + ((Date.now() - START_TS) / 60000).toFixed(2) + ' minutes.');
  Logger.log('  Happy mining. Questions? scripts@we-discover.com');
  Logger.log('============================================================');
}


// -----------------------------------------------------------------------------
// validateConfig()
// Runs basic sanity checks on CONFIG before anything else happens.
// Logs a warning rather than throwing, so the script still runs.
// -----------------------------------------------------------------------------
function validateConfig() {
  if (CONFIG.maxNGramLength > 6) {
    Logger.log('[ WeDiscover | WARNING ] maxNGramLength > 6 on a high-volume account risks timeout and memory errors. Consider keeping it at 4.');
  }
  if (AI_CONFIG.apiKey && AI_CONFIG.apiKey.indexOf('YOUR_') > -1) {
    Logger.log('[ WeDiscover | WARNING ] AI_CONFIG.apiKey looks like a placeholder. Replace it with your actual Google AI Studio key or leave it blank.');
  }
  if (AI_CONFIG.apiKey && AI_CONFIG.mode !== 'account' && AI_CONFIG.mode !== 'full') {
    Logger.log('[ WeDiscover | WARNING ] AI_CONFIG.mode must be "account" or "full". Defaulting to "account".');
  }
}


// -----------------------------------------------------------------------------
// openSpreadsheet()
// Opens the spreadsheet at the URL in CONFIG and removes any blank sheets
// (such as the default "Sheet1" that Google creates automatically).
// Throws a clear error if the URL is missing or the sheet cannot be opened.
// -----------------------------------------------------------------------------
function openSpreadsheet() {
  if (!CONFIG.spreadsheetUrl || CONFIG.spreadsheetUrl.indexOf('YOUR_SPREADSHEET') > -1) {
    throw new Error(
      'No spreadsheet URL set. Create a Google Sheet, paste its URL into ' +
      'CONFIG.spreadsheetUrl, and re-run the script.'
    );
  }

  var ss;
  try {
    ss = SpreadsheetApp.openByUrl(CONFIG.spreadsheetUrl);
  } catch (e) {
    throw new Error(
      'Could not open the spreadsheet. Make sure the sheet exists and the ' +
      'account running this script has edit access. URL: ' +
      CONFIG.spreadsheetUrl + ' Error: ' + e
    );
  }

  return ss;
}


// -----------------------------------------------------------------------------
// deleteEmptySheets(ss)
// Deletes any sheet that has no content (max row = 0 or a single empty cell,
// as Google creates for a brand-new sheet). Called at the end of main() so it
// catches both pre-existing blank sheets and any that Google quietly adds
// between runs. We check ss.getSheets().length before each deletion to avoid
// leaving the spreadsheet with zero sheets, which the API disallows.
// -----------------------------------------------------------------------------
function deleteEmptySheets(ss) {
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getLastRow() <= 1 && ss.getSheets().length > 1) {
      // getLastRow() returns 0 for a truly empty sheet and 1 if there is a
      // single empty row (the state Google leaves a new sheet in). Either way,
      // there is no script-written content and the sheet can go.
      Logger.log('         Deleted empty sheet: ' + allSheets[i].getName());
      ss.deleteSheet(allSheets[i]);
    }
  }
}


// -----------------------------------------------------------------------------
// buildDateRange()
// Converts the DD/MM/YYYY dates in CONFIG into the GAQL date filter string
// that AdsApp.search() expects. Throws if the format is wrong.
// -----------------------------------------------------------------------------
function buildDateRange() {
  function toGaqlDate(ddmmyyyy) {
    var parts = ddmmyyyy.split('/');
    if (parts.length !== 3) {
      throw new Error(
        'Date "' + ddmmyyyy + '" is not in DD/MM/YYYY format. ' +
        'Example: 23/06/2025'
      );
    }
    // Rearrange from DD/MM/YYYY to YYYY-MM-DD for GAQL.
    return parts[2] + '-' + parts[1] + '-' + parts[0];
  }
  return "segments.date BETWEEN '" + toGaqlDate(CONFIG.startDate) + "' AND '" + toGaqlDate(CONFIG.endDate) + "'";
}


// -----------------------------------------------------------------------------
// fetchActiveCampaignIds(dateRange)
// Returns an array of campaign IDs that had impressions in the date range
// and match any name filters set in CONFIG. These IDs are used to scope all
// subsequent queries so we only pull data we actually need.
// -----------------------------------------------------------------------------
function fetchActiveCampaignIds(dateRange) {
  const statusClause = CONFIG.ignorePausedCampaigns
    ? "campaign.status = 'ENABLED'"
    : "campaign.status IN ('ENABLED', 'PAUSED')";

  const nameContains = CONFIG.campaignNameContains
    ? "AND campaign.name LIKE '%" + CONFIG.campaignNameContains + "%'"
    : '';
  const nameExcludes = CONFIG.campaignNameDoesNotContain
    ? "AND campaign.name NOT LIKE '%" + CONFIG.campaignNameDoesNotContain + "%'"
    : '';

  const query = [
    'SELECT campaign.id',
    'FROM   campaign',
    'WHERE  ' + statusClause,
    nameContains,
    nameExcludes,
    'AND metrics.impressions > 0',
    'AND ' + dateRange,
  ].join(' ');

  const ids = [];
  for (const row of AdsApp.search(query)) {
    ids.push(row.campaign.id.toString());
  }
  return ids;
}


// -----------------------------------------------------------------------------
// fetchAllNegatives(campaignIds)
// Builds two Maps of negative keywords: one keyed by ad group ID and one by
// campaign ID. Includes negatives from shared negative keyword lists.
//
// We use Maps rather than arrays so isExcludedByNegative() can look up the
// right list in O(1) instead of scanning everything on every row.
//
// Each Map value is an array of [keywordText, matchType] pairs.
// -----------------------------------------------------------------------------
function fetchAllNegatives(campaignIds, _dateRange) {
  const negsByAdGroup  = new Map();
  const negsByCampaign = new Map();

  const adGroupStatusClause = CONFIG.ignorePausedAdGroups
    ? 'AND ad_group.status = \'ENABLED\''
    : 'AND ad_group.status IN (\'ENABLED\', \'PAUSED\')';

  const campaignIdList = campaignIds.join(',');

  // Ad group level negatives.
  const agNegQuery = [
    'SELECT ad_group.id,',
    '       ad_group_criterion.keyword.text,',
    '       ad_group_criterion.keyword.match_type',
    'FROM   ad_group_criterion',
    'WHERE  ad_group_criterion.negative = TRUE',
    '       AND ad_group_criterion.type = \'KEYWORD\'',
    '       AND campaign.id IN (' + campaignIdList + ')',
    adGroupStatusClause,
  ].join(' ');

  for (const row of AdsApp.search(agNegQuery)) {
    const adGroupId = row.adGroup.id.toString();
    const entry = [
      row.adGroupCriterion.keyword.text.toLowerCase(),
      row.adGroupCriterion.keyword.matchType.toLowerCase(),
    ];
    if (!negsByAdGroup.has(adGroupId)) negsByAdGroup.set(adGroupId, []);
    negsByAdGroup.get(adGroupId).push(entry);
  }

  // Campaign level negatives.
  const campNegQuery = [
    'SELECT campaign.id,',
    '       campaign_criterion.keyword.text,',
    '       campaign_criterion.keyword.match_type',
    'FROM   campaign_criterion',
    'WHERE  campaign_criterion.negative = TRUE',
    '       AND campaign_criterion.type = \'KEYWORD\'',
    '       AND campaign.id IN (' + campaignIdList + ')',
  ].join(' ');

  for (const row of AdsApp.search(campNegQuery)) {
    const campaignId = row.campaign.id.toString();
    const entry = [
      row.campaignCriterion.keyword.text.toLowerCase(),
      row.campaignCriterion.keyword.matchType.toLowerCase(),
    ];
    if (!negsByCampaign.has(campaignId)) negsByCampaign.set(campaignId, []);
    negsByCampaign.get(campaignId).push(entry);
  }

  // Shared negative keyword lists.
  // First find which shared sets each campaign uses, then fetch the keywords
  // from those sets and add them to the campaign-level negatives Map.
  const sharedSetToCampaigns = new Map();

  const sharedSetMemberQuery = [
    'SELECT campaign.id, campaign.name, shared_set.id, shared_set.name',
    'FROM   campaign_shared_set',
    'WHERE  shared_set.type = \'NEGATIVE_KEYWORDS\'',
    '       AND campaign.id IN (' + campaignIdList + ')',
  ].join(' ');

  for (const row of AdsApp.search(sharedSetMemberQuery)) {
    const setId = row.sharedSet.id.toString();
    if (!sharedSetToCampaigns.has(setId)) sharedSetToCampaigns.set(setId, []);
    sharedSetToCampaigns.get(setId).push(row.campaign.id.toString());
  }

  if (sharedSetToCampaigns.size > 0) {
    const sharedCriteriaQuery = [
      'SELECT shared_set.id,',
      '       shared_criterion.keyword.text,',
      '       shared_criterion.keyword.match_type',
      'FROM   shared_criterion',
      'WHERE  shared_set.id IN (' + Array.from(sharedSetToCampaigns.keys()).join(',') + ')',
    ].join(' ');

    for (const row of AdsApp.search(sharedCriteriaQuery)) {
      const setId     = row.sharedSet.id.toString();
      const campaigns = sharedSetToCampaigns.get(setId) || [];
      const entry = [
        row.sharedCriterion.keyword.text.toLowerCase(),
        row.sharedCriterion.keyword.matchType.toLowerCase(),
      ];
      for (const campaignId of campaigns) {
        if (!negsByCampaign.has(campaignId)) negsByCampaign.set(campaignId, []);
        negsByCampaign.get(campaignId).push(entry);
      }
    }
  }

  return { negsByAdGroup, negsByCampaign };
}


// -----------------------------------------------------------------------------
// isExcludedByNegative(row, negsByAdGroup, negsByCampaign)
// Returns true if the search query in this row would be blocked by any of
// the negative keywords at the ad group or campaign level.
//
// Exact match negatives must match the whole query exactly.
// Phrase and broad match negatives just need to appear as whole words
// somewhere in the query. We pad the query with spaces to match whole words
// only rather than substrings (e.g. "run" should not match "running").
// -----------------------------------------------------------------------------
function isExcludedByNegative(row, negsByAdGroup, negsByCampaign) {
  const query      = row.query;
  const adGroupId  = row.adGroupId;
  const campaignId = row.campaignId;

  const isMatch = function(neg) {
    var negText   = neg[0];
    var matchType = neg[1];
    if (matchType === 'exact') return query === negText;
    return (' ' + query + ' ').indexOf(' ' + negText + ' ') > -1;
  };

  if (negsByAdGroup.has(adGroupId)   && negsByAdGroup.get(adGroupId).some(isMatch))   return true;
  if (negsByCampaign.has(campaignId) && negsByCampaign.get(campaignId).some(isMatch)) return true;
  return false;
}


// -----------------------------------------------------------------------------
// fetchSearchQueryRows(campaignIds, dateRange)
// Returns an AdsApp iterator over the search term view for the given campaigns
// and date range. We return the iterator directly rather than collecting rows
// into an array, so the full report is never held in memory at once.
//
// Note: generator functions (function*) are not used because the Google Ads
// Scripts runtime does not support them reliably.
// -----------------------------------------------------------------------------
function fetchSearchQueryRows(campaignIds, dateRange) {
  const adGroupStatusClause = CONFIG.ignorePausedAdGroups
    ? 'AND ad_group.status = \'ENABLED\''
    : 'AND ad_group.status IN (\'ENABLED\', \'PAUSED\')';

  const query = [
    'SELECT campaign.name,',
    '       campaign.id,',
    '       ad_group.name,',
    '       ad_group.id,',
    '       search_term_view.search_term,',
    '       metrics.clicks,',
    '       metrics.impressions,',
    '       metrics.cost_micros,',
    '       metrics.conversions,',
    '       metrics.conversions_value',
    'FROM   search_term_view',
    'WHERE  campaign.id IN (' + campaignIds.join(',') + ')',
    adGroupStatusClause,
    'AND ' + dateRange,
  ].join(' ');

  return AdsApp.search(query);
}


// -----------------------------------------------------------------------------
// parseQueryRow(row)
// Maps a raw AdsApp iterator row to a plain object with named fields.
// Keeping this separate from the iterator loop means the data shape is defined
// in one place and is easy to update.
//
// We use parseInt and parseFloat because the Ads Scripts runtime sometimes
// returns metric fields as strings rather than numbers. Without this, adding
// two "numbers" would concatenate them as strings instead of summing them.
// Cost arrives in micros (millionths of the currency unit), so we divide by
// 1,000,000 to convert to the actual currency amount.
// -----------------------------------------------------------------------------
function parseQueryRow(row) {
  return {
    campaignName:     row.campaign.name,
    campaignId:       row.campaign.id.toString(),
    adGroupName:      row.adGroup.name,
    adGroupId:        row.adGroup.id.toString(),
    query:            row.searchTermView.searchTerm.toLowerCase(),
    clicks:           parseInt(row.metrics.clicks, 10)           || 0,
    impressions:      parseInt(row.metrics.impressions, 10)       || 0,
    cost:             (parseInt(row.metrics.costMicros, 10) || 0) / 1000000,
    conversions:      parseFloat(row.metrics.conversions)         || 0,
    conversionsValue: parseFloat(row.metrics.conversionsValue)    || 0,
  };
}


// -----------------------------------------------------------------------------
// buildEmptyNgramMaps()
// Creates a fresh set of accumulator Maps, one for each n-gram length.
// Each length gets three Maps: total (account level), campaign, and ad group.
// Returns a Map keyed by n (the phrase length).
// -----------------------------------------------------------------------------
function buildEmptyNgramMaps() {
  const maps = new Map();
  for (let n = CONFIG.minNGramLength; n <= CONFIG.maxNGramLength; n++) {
    maps.set(n, {
      total:    new Map(),
      campaign: new Map(),
      adgroup:  new Map(),
    });
  }
  return maps;
}


// -----------------------------------------------------------------------------
// accumulateNgrams(row, ngramMaps, wordCountMap)
// Extracts every n-gram of every configured length from the search query,
// then adds this row's stats to the running total for each one.
//
// Keys are compound strings using a null byte separator (e.g.
// "Campaign Name\x00=\"running shoes\"") so we can store campaign and ad group
// level data in a flat Map without nested objects.
//
// A Set tracks which phrases have already been seen within this query so the
// same phrase in a long query is not counted more than once.
// -----------------------------------------------------------------------------
function accumulateNgrams(row, ngramMaps, wordCountMap) {
  const words = row.query.split(' ');
  const stats = statsFromRow(row);

  // Track how many words the query had (capped at "7+" for very long queries).
  const lenKey = words.length < 7 ? words.length : '7+';
  addStats(wordCountMap, lenKey, stats);

  for (let n = CONFIG.minNGramLength; n <= CONFIG.maxNGramLength; n++) {
    if (n > words.length) break;

    const level = ngramMaps.get(n);
    const seen  = new Set();

    for (let w = 0; w <= words.length - n; w++) {
      const phrase = words.slice(w, w + n).join(' ');
      if (seen.has(phrase)) continue;
      seen.add(phrase);

      // The =" prefix forces Google Sheets to treat the cell as plain text,
      // which stops phrases like "1000" being interpreted as numbers.
      const displayPhrase = '="' + phrase + '"';

      addStats(level.total,    displayPhrase, stats);
      addStats(level.campaign, row.campaignName + '\x00' + displayPhrase, stats);
      addStats(level.adgroup,  row.campaignName + '\x00' + row.adGroupName + '\x00' + displayPhrase, stats);
    }
  }
}


// -----------------------------------------------------------------------------
// statsFromRow(row)
// Pulls the metric fields out of a parsed query row and returns a plain object.
// queryCount starts at 1 because this represents one query row.
// -----------------------------------------------------------------------------
function statsFromRow(row) {
  return {
    clicks:           row.clicks,
    impressions:      row.impressions,
    cost:             row.cost,
    conversions:      row.conversions,
    conversionsValue: row.conversionsValue,
    queryCount:       1,
  };
}


// -----------------------------------------------------------------------------
// addStats(map, key, stats)
// Adds the stats from one query row to the running total stored in the Map
// under the given key. Creates a new entry if the key does not exist yet.
// Object.assign is used on creation so the original stats object is not mutated.
// -----------------------------------------------------------------------------
function addStats(map, key, stats) {
  if (!map.has(key)) {
    map.set(key, Object.assign({}, stats));
    return;
  }
  const existing = map.get(key);
  existing.clicks           += stats.clicks;
  existing.impressions      += stats.impressions;
  existing.cost             += stats.cost;
  existing.conversions      += stats.conversions;
  existing.conversionsValue += stats.conversionsValue;
  existing.queryCount       += 1;
}


// -----------------------------------------------------------------------------
// buildFilterText()
// Builds a human-readable summary of which campaigns and ad groups are included.
// This appears in the teal status bar at the top of each sheet.
// -----------------------------------------------------------------------------
function buildFilterText() {
  let text = CONFIG.ignorePausedAdGroups ? 'Active ad groups' : 'All ad groups';
  text += CONFIG.ignorePausedCampaigns ? ' in active campaigns' : ' in all campaigns';
  if (CONFIG.campaignNameContains)       text += " containing '" + CONFIG.campaignNameContains + "'";
  if (CONFIG.campaignNameDoesNotContain) text += " not containing '" + CONFIG.campaignNameDoesNotContain + "'";
  return text;
}


// =============================================================================
// SPREADSHEET OUTPUT
// Everything below handles creating, writing to, and formatting the sheets.
// =============================================================================


// -----------------------------------------------------------------------------
// WeDiscover brand colours
// Matched from we-discover.com. Used throughout the sheet styling functions.
// Grotesque Medium and Hot Sans (WeDiscover's brand fonts) are not available
// in Google Sheets. Montserrat is the closest available match.
// -----------------------------------------------------------------------------
var WD_CRIMSON = '#C0392B';  // Primary brand red (title bars)
var WD_TEAL    = '#3ECFB2';  // Accent teal (status bar)
var WD_NAVY    = '#1A1F36';  // Dark navy (column headers)
var WD_CREAM   = '#FAF8F5';  // Warm off-white (alternating data rows)
var WD_WHITE   = '#FFFFFF';
var WD_FONT    = 'Montserrat';


// -----------------------------------------------------------------------------
// buildSheetDefs()
// Returns an array describing every sheet the script will write to.
// Each entry contains:
//   name        -- the sheet tab name
//   nLabel      -- short label used in the sheet title ("Words", "2-Grams" etc.)
//   levelName   -- "Account", "Campaign", or "Ad Group"
//   mapKey      -- key used to look up the right accumulator Map in flushToSheets
//   header      -- the column header row as an array of strings
//   keyParser   -- function that splits a compound Map key into label columns
//   labelCount  -- number of label columns before Query Count (used by dedup pass)
//
// This central definition drives initialiseSheets, flushToSheets, and
// finaliseSheets, so adding a new sheet only requires one change here.
// -----------------------------------------------------------------------------
function buildSheetDefs() {
  const defs = [];

  defs.push({
    name:       'Word Count Analysis',
    nLabel:     'Word Count',
    levelName:  'Account',
    mapKey:     'wordcount',
    header:     ['Word Count', 'Query Count'].concat(STAT_LABELS).concat(CALC_STATS.map(function(c) { return c[0]; })),
    keyParser:  function(key) { return ['' + key]; },
    labelCount: 1,
  });

  for (let n = CONFIG.minNGramLength; n <= CONFIG.maxNGramLength; n++) {
    const nLabel = n === 1 ? 'Words' : n + '-Grams';

    defs.push({
      name:       n === 1 ? 'Account Word Analysis'  : 'Account ' + n + '-Gram Analysis',
      nLabel:     nLabel,
      levelName:  'Account',
      mapKey:     'total_' + n,
      header:     ['Phrase', 'Query Count'].concat(STAT_LABELS).concat(CALC_STATS.map(function(c) { return c[0]; })),
      keyParser:  function(key) { return [key]; },
      labelCount: 1,
    });
    defs.push({
      name:       n === 1 ? 'Campaign Word Analysis' : 'Campaign ' + n + '-Gram Analysis',
      nLabel:     nLabel,
      levelName:  'Campaign',
      mapKey:     'campaign_' + n,
      header:     ['Campaign', 'Phrase', 'Query Count'].concat(STAT_LABELS).concat(CALC_STATS.map(function(c) { return c[0]; })),
      keyParser:  function(key) { return key.split('\x00'); },
      labelCount: 2,
    });
    defs.push({
      name:       n === 1 ? 'Ad Group Word Analysis' : 'Ad Group ' + n + '-Gram Analysis',
      nLabel:     nLabel,
      levelName:  'Ad Group',
      mapKey:     'adgroup_' + n,
      header:     ['Campaign', 'Ad Group', 'Phrase', 'Query Count'].concat(STAT_LABELS).concat(CALC_STATS.map(function(c) { return c[0]; })),
      keyParser:  function(key) { return key.split('\x00'); },
      labelCount: 3,
    });
  }

  return defs;
}


// -----------------------------------------------------------------------------
// styleSheet(sheet, def)
// Applies WeDiscover brand styling to the three header rows of a sheet.
//   Row 1: crimson title bar.
//   Row 2: teal status bar (value is set separately in initialiseSheets).
//   Row 3: navy column header row.
// Also sets the tab colour, freezes the header rows, and sets column widths.
// Label columns (phrase, campaign name etc.) are wider than stat columns.
// -----------------------------------------------------------------------------
function styleSheet(sheet, def) {
  var numCols = def.header.length;

  var titleLabel = def.nLabel === 'Word Count'
    ? 'WeDiscover  |  Search Query Performance by Word Count'
    : 'WeDiscover  |  ' + def.nLabel + ' Analysis  |  By ' + def.levelName;

  sheet.setFrozenRows(3);

  var r1 = sheet.getRange(1, 1, 1, numCols);
  r1.merge();
  r1.setValue(titleLabel);
  r1.setBackground(WD_CRIMSON);
  r1.setFontColor(WD_WHITE);
  r1.setFontFamily(WD_FONT);
  r1.setFontSize(12);
  r1.setFontWeight('bold');
  r1.setVerticalAlignment('middle');
  r1.setHorizontalAlignment('left');
  sheet.setRowHeight(1, 40);

  var r2 = sheet.getRange(2, 1, 1, numCols);
  r2.merge();
  r2.setBackground(WD_TEAL);
  r2.setFontColor(WD_NAVY);
  r2.setFontFamily(WD_FONT);
  r2.setFontSize(9);
  r2.setFontWeight('bold');
  r2.setVerticalAlignment('middle');
  r2.setHorizontalAlignment('left');
  sheet.setRowHeight(2, 24);

  var r3 = sheet.getRange(3, 1, 1, numCols);
  r3.setBackground(WD_NAVY);
  r3.setFontColor(WD_WHITE);
  r3.setFontFamily(WD_FONT);
  r3.setFontSize(9);
  r3.setFontWeight('bold');
  r3.setHorizontalAlignment('center');
  r3.setVerticalAlignment('middle');
  sheet.setRowHeight(3, 28);

  sheet.setTabColor(WD_CRIMSON);

  // 5 raw stats + 5 calculated stats = 10 stat columns at the right.
  // Everything to the left of those is a label column and gets more space.
  var labelCols = numCols - 10;
  for (var c = 1; c <= numCols; c++) {
    sheet.setColumnWidth(c, c <= labelCols ? 160 : 120);
  }
  sheet.setColumnWidth(1, 240);
}


// -----------------------------------------------------------------------------
// applyFilter(sheet, def)
// Removes any existing filter on the sheet and applies a fresh one covering
// the header row and all data rows. Removing first ensures a clean state
// when the script runs more than once on the same sheet.
// -----------------------------------------------------------------------------
function applyFilter(sheet, def) {
  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;
  sheet.getRange(3, 1, lastRow - 2, def.header.length).createFilter();
}


// -----------------------------------------------------------------------------
// initialiseSheets(ss, filterText)
// Creates all output sheets (or clears them if they already exist), applies
// brand styling, writes the header row, and returns a registry Map.
//
// The registry stores { sheet, def, nextRow } for each sheet name.
// nextRow tracks where the next batch of data should be written, avoiding
// the need to call getLastRow() on every flush and saving Sheets API calls.
// -----------------------------------------------------------------------------
function initialiseSheets(ss, filterText) {
  var defs     = buildSheetDefs();
  var registry = new Map();

  for (var i = 0; i < defs.length; i++) {
    var def = defs[i];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) sheet = ss.insertSheet(def.name);
    if (CONFIG.clearSpreadsheet) sheet.clear();

    styleSheet(sheet, def);

    // Row 2 value must be set after styleSheet() calls merge() on that row,
    // otherwise the merge wipes the value.
    sheet.getRange('A2').setValue(filterText);

    // Write header values into row 3, then re-apply styling because setValues()
    // strips any formatting that was applied beforehand.
    var hRange = sheet.getRange(3, 1, 1, def.header.length);
    hRange.setValues([def.header]);
    hRange.setBackground(WD_NAVY);
    hRange.setFontColor(WD_WHITE);
    hRange.setFontWeight('bold');

    registry.set(def.name, { sheet: sheet, def: def, nextRow: 4 });
  }

  Logger.log('         ' + defs.length + ' sheets created and headers written.');
  return registry;
}


// -----------------------------------------------------------------------------
// flushToSheets(sheets, ngramMaps, wordCountMap)
// Writes the current batch of accumulated n-gram data to the spreadsheet.
// Called every FLUSH_EVERY rows during processing, and once at the end.
//
// Only rows that meet all threshold settings are written.
//
// entry.nextRow is used instead of sheet.getLastRow() to avoid an extra
// API call per sheet per flush. SpreadsheetApp.flush() at the end commits
// all pending writes in one round trip.
//
// Rows written here may be duplicates of rows from a previous batch if the
// same phrase appeared in both. This is intentional: keeping raw sums per
// batch is cheaper than maintaining a single growing Map across all batches.
// The deduplication pass in finaliseSheets() merges them correctly.
// -----------------------------------------------------------------------------
function flushToSheets(sheets, ngramMaps, wordCountMap) {
  const t = CONFIG.thresholds;

  // Build a lookup from mapKey to the relevant accumulator Map.
  const mapLookup = new Map();
  mapLookup.set('wordcount', wordCountMap);
  for (let n = CONFIG.minNGramLength; n <= CONFIG.maxNGramLength; n++) {
    const level = ngramMaps.get(n);
    mapLookup.set('total_' + n,    level.total);
    mapLookup.set('campaign_' + n, level.campaign);
    mapLookup.set('adgroup_' + n,  level.adgroup);
  }

  for (const [name, entry] of sheets) {
    const sourceMap = mapLookup.get(entry.def.mapKey);
    if (!sourceMap || sourceMap.size === 0) continue;

    const rows = [];
    for (const [key, s] of sourceMap) {
      if (s.queryCount  < t.queryCount)  continue;
      if (s.impressions < t.impressions) continue;
      if (s.clicks      < t.clicks)      continue;
      if (s.cost        < t.cost)        continue;
      if (s.conversions < t.conversions) continue;

      rows.push(buildPrintline(entry.def.keyParser(key), s));
    }

    if (rows.length === 0) continue;

    entry.sheet.getRange(entry.nextRow, 1, rows.length, entry.def.header.length).setValues(rows);
    entry.nextRow += rows.length;
  }

  SpreadsheetApp.flush();
}


// -----------------------------------------------------------------------------
// buildPrintline(labelParts, s)
// Builds a single output row as a flat array.
// labelParts contains the text columns (phrase, campaign name etc.).
// s is the stats object for this phrase.
// Calculated stats (CTR, CPC etc.) are appended at the end.
// If a denominator is zero, a hyphen is written rather than dividing by zero.
// -----------------------------------------------------------------------------
function buildPrintline(labelParts, s) {
  const line = labelParts.concat([s.queryCount,
    s.clicks, s.impressions, s.cost, s.conversions, s.conversionsValue]);

  for (var i = 0; i < CALC_STATS.length; i++) {
    var num = CALC_STATS[i][1];
    var den = CALC_STATS[i][2];
    line.push(s[den] > 0 ? s[num] / s[den] : '-');
  }
  return line;
}


// =============================================================================
// DEDUPLICATION
// The flush-and-reset architecture means the same phrase can appear as multiple
// rows across different batches. The functions below fix this by reading each
// sheet back after all batches are written, merging rows with matching keys,
// and rewriting a clean, deduplicated set of rows.
//
// Why do this here rather than keeping a persistent Map?
// A persistent Map would hold every unique phrase seen so far. On a large
// account with 2M+ search term rows and 1-to-4-gram analysis, that Map grows
// to hundreds of thousands of entries and eventually causes memory errors.
// The flush-and-reset approach keeps memory flat. The deduplication pass reads
// one sheet at a time, which Google Sheets handles fine, and the resulting
// merged Map is much smaller than the full raw-data Map would have been.
// =============================================================================


// -----------------------------------------------------------------------------
// deduplicateSheet(entry)
// Reads all data rows from a sheet, merges rows that share the same key
// (label columns), recalculates derived stats from the merged raw totals,
// and writes the clean rows back to the sheet starting at row 4.
//
// The key for each row is formed by joining its label columns with a null byte.
// The number of label columns is def.labelCount (e.g. 1 for account-level,
// 2 for campaign-level, 3 for ad group-level sheets).
//
// Column layout of each data row (0-indexed):
//   [0 .. labelCount-1]              label columns (phrase, campaign etc.)
//   [labelCount]                     Query Count
//   [labelCount+1 .. labelCount+5]   raw stats (clicks, impressions, cost,
//                                    conversions, conversionsValue)
//   [labelCount+6 .. labelCount+10]  calculated stats (ignored on read,
//                                    recalculated on write)
//
// Returns the number of rows written after deduplication.
// -----------------------------------------------------------------------------
function deduplicateSheet(entry) {
  const sheet      = entry.sheet;
  const def        = entry.def;
  const lastRow    = entry.nextRow - 1;  // nextRow points to the next empty row
  const firstData  = 4;

  if (lastRow < firstData) return 0;

  const numCols    = def.header.length;
  const dataRows   = sheet.getRange(firstData, 1, lastRow - firstData + 1, numCols).getValues();

  // labelCount tells us how many leading columns are labels (not numbers).
  // After those come: queryCount, then RAW_STAT_COUNT raw stat columns.
  const lc         = def.labelCount;
  const qcIdx      = lc;       // column index of Query Count
  const statStart  = lc + 1;   // column index of first raw stat (Clicks)

  // Merge rows by key, summing all numeric columns.
  // We preserve insertion order so the first occurrence of each key determines
  // the key order going into the sort step.
  const merged  = new Map();
  const keyOrder = [];

  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];

    // Build the compound key from the label columns.
    var keyParts = [];
    for (var k = 0; k < lc; k++) {
      keyParts.push(row[k]);
    }
    var key = keyParts.join('\x00');

    if (!merged.has(key)) {
      // First time we have seen this key: store labels and initialise counters.
      var labels = [];
      for (var j = 0; j < lc; j++) {
        labels.push(row[j]);
      }
      merged.set(key, {
        labels:           labels,
        queryCount:       0,
        clicks:           0,
        impressions:      0,
        cost:             0,
        conversions:      0,
        conversionsValue: 0,
      });
      keyOrder.push(key);
    }

    var acc = merged.get(key);

    // Query Count may have been written as a string by a previous run, so
    // we coerce it to a number before adding. Same risk as the Ads Scripts
    // runtime returning metrics as strings.
    acc.queryCount       += parseFloat(row[qcIdx])          || 0;
    acc.clicks           += parseFloat(row[statStart])      || 0;
    acc.impressions      += parseFloat(row[statStart + 1])  || 0;
    acc.cost             += parseFloat(row[statStart + 2])  || 0;
    acc.conversions      += parseFloat(row[statStart + 3])  || 0;
    acc.conversionsValue += parseFloat(row[statStart + 4])  || 0;
  }

  // Rebuild the data rows from the merged Map.
  // buildPrintline expects { clicks, impressions, cost, conversions,
  // conversionsValue, queryCount } as its second argument, which matches
  // exactly what we stored in each merged entry.
  const cleanRows = [];
  for (var ki = 0; ki < keyOrder.length; ki++) {
    var acc2 = merged.get(keyOrder[ki]);
    cleanRows.push(buildPrintline(acc2.labels, acc2));
  }

  // Write clean rows back, starting at row 4. We know exactly how many rows
  // we are writing, so we can clear only what we need to clear and avoid
  // leaving stale data below the new rows.
  //
  // Clear the entire data area first (old row count may be larger or smaller
  // than the new row count, so we cannot rely on overwriting in place).
  sheet.getRange(firstData, 1, lastRow - firstData + 1, numCols).clearContent();
  if (cleanRows.length > 0) {
    sheet.getRange(firstData, 1, cleanRows.length, numCols).setValues(cleanRows);
  }

  SpreadsheetApp.flush();
  return cleanRows.length;
}


// -----------------------------------------------------------------------------
// HEADER_NOTES
// Hover tooltips for each column header. These appear when a user hovers over
// a cell in row 3. They explain what each column means without cluttering
// the view for users who already know.
// -----------------------------------------------------------------------------
var HEADER_NOTES = {
  'Word Count':      'Number of words in the search query',
  'Phrase':          'The n-gram phrase. The =" prefix stops Google Sheets treating it as a formula.',
  'Query Count':     'Number of distinct search queries that contain this phrase',
  'Campaign':        'Google Ads campaign name',
  'Ad Group':        'Google Ads ad group name',
  'Clicks':          'Total clicks from queries containing this phrase',
  'Impressions':     'Total impressions from queries containing this phrase',
  'Cost':            'Total spend from queries containing this phrase (in account currency)',
  'Conversions':     'Total conversions from queries containing this phrase',
  'Conv. Value':     'Total conversion value from queries containing this phrase',
  'CTR':             'Click-through rate: Clicks divided by Impressions',
  'CPC':             'Cost per click: Cost divided by Clicks',
  'Conv. Rate':      'Conversion rate: Conversions divided by Clicks',
  'Cost / Conv.':    'Cost per conversion: Cost divided by Conversions',
  'Conv. Val/Cost':  'Return on ad spend: Conversion Value divided by Cost',
};


// -----------------------------------------------------------------------------
// addHeaderNotes(sheet, header)
// Adds the hover notes from HEADER_NOTES to the column header cells in row 3.
// -----------------------------------------------------------------------------
function addHeaderNotes(sheet, header) {
  for (var c = 0; c < header.length; c++) {
    var note = HEADER_NOTES[header[c]];
    if (note) {
      sheet.getRange(3, c + 1).setNote(note);
    }
  }
}


// -----------------------------------------------------------------------------
// applyAlternatingRows(sheet, firstDataRow, dataRowCount, numCols)
// Colours data rows alternately white and warm cream.
// This makes it easier to read across wide rows with many stat columns.
// -----------------------------------------------------------------------------
function applyAlternatingRows(sheet, firstDataRow, dataRowCount, numCols) {
  for (var r = 0; r < dataRowCount; r++) {
    var bg = (r % 2 === 0) ? WD_WHITE : WD_CREAM;
    sheet.getRange(firstDataRow + r, 1, 1, numCols).setBackground(bg);
  }
}


// -----------------------------------------------------------------------------
// setDataFont(sheet, firstDataRow, dataRowCount, numCols, labelCount)
// Sets data rows to Arial. Montserrat looks good on headers but renders slowly
// at thousands of rows. Arial is fast, clean, and universally available.
//
// Alignment is set in two passes:
//   1. Right-align the entire data range -- numeric and calculated stat columns
//      should all be right-aligned, and the hyphen written for zero-denominator
//      stats (e.g. Cost / Conv. with no conversions) is a string so would
//      left-align by default without this explicit override.
//   2. Left-align just the label columns (phrase, campaign, ad group) so
//      text values read naturally from the left.
// -----------------------------------------------------------------------------
function setDataFont(sheet, firstDataRow, dataRowCount, numCols, labelCount) {
  sheet.getRange(firstDataRow, 1, dataRowCount, numCols)
    .setFontFamily('Arial')
    .setFontSize(9)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('right');

  // Left-align the label columns over the top of the right-align above.
  sheet.getRange(firstDataRow, 1, dataRowCount, labelCount)
    .setHorizontalAlignment('left');
}


// -----------------------------------------------------------------------------
// trimSheet(sheet, usedCols, usedRows)
// Deletes columns and rows beyond the data range. Google Sheets creates 26
// columns and 1,000 rows by default. Trimming keeps the sheet clean and
// reduces the file size when opening large spreadsheets.
//
// usedRows is the index of the last row that contains data (i.e. the row
// number of the last data row, not a count). We delete everything after it so
// the sheet ends exactly at the last data row with no trailing empty rows.
// -----------------------------------------------------------------------------
function trimSheet(sheet, usedCols, usedRows) {
  var maxCols = sheet.getMaxColumns();
  if (maxCols > usedCols) {
    sheet.deleteColumns(usedCols + 1, maxCols - usedCols);
  }
  var maxRows = sheet.getMaxRows();
  if (maxRows > usedRows) {
    sheet.deleteRows(usedRows + 1, maxRows - usedRows);
  }
}


// -----------------------------------------------------------------------------
// finaliseSheets(sheets, finalFilterText)
// Runs once after all data has been flushed. For each sheet it:
//   1. Deduplicates rows written across multiple flush batches.
//   2. Updates the teal status bar with the final filter text.
//   3. Applies number formatting to all data rows in one API call.
//   4. Sorts by label column(s) ascending, then clicks descending.
//   5. Applies alternating row colours.
//   6. Sets the data font to Arial.
//   7. Adds hover notes to the column headers.
//   8. Trims unused rows and columns.
//   9. Adds a column filter to the header row.
//
// Deduplication must run before sorting, formatting, and filtering, because
// those steps all depend on the final row count. Running dedup first also
// means sort and format only touch the deduplicated rows, which is faster.
//
// Doing all formatting here (rather than on each flush) means we only pay
// the Sheets API cost once per sheet, regardless of how many flushes ran.
// -----------------------------------------------------------------------------
function finaliseSheets(sheets, finalFilterText) {
  for (const [, entry] of sheets) {
    const sheet = entry.sheet;
    const def   = entry.def;

    // Merge duplicate rows before doing anything else. entry.nextRow may
    // overcount if the same phrase appeared in multiple flush batches, so
    // we update it from the return value of deduplicateSheet().
    const dedupedRowCount = deduplicateSheet(entry);
    entry.nextRow = 4 + dedupedRowCount;

    Logger.log('         Deduplicated: ' + sheet.getName() + ' (' + (entry.nextRow - 4).toLocaleString() + ' unique rows)');

    // Update the teal status bar and re-apply styling because setValue() can
    // strip formatting on merged cells.
    sheet.getRange('A2').setValue(finalFilterText);
    var r2 = sheet.getRange(2, 1, 1, def.header.length);
    r2.setBackground(WD_TEAL);
    r2.setFontColor(WD_NAVY);
    r2.setFontSize(9);
    r2.setFontWeight('bold');

    const dataRowCount = entry.nextRow - 4;
    if (dataRowCount <= 0) {
      sheet.getRange('A4').setValue('No ' + def.nLabel.toLowerCase() + ' found within thresholds.');
      continue;
    }

    const numCols = def.header.length;

    // Number formatting.
    // Label columns use '@' (plain text format). Query Count uses '#,##0'.
    // Stat and calculated stat columns use the formats defined in COL_FORMATS.
    const labelCount = numCols - COL_FORMATS.length - 1;
    const fullFmtRow = Array(labelCount).fill('@')
      .concat(['#,##0'])
      .concat(COL_FORMATS);
    sheet.getRange(4, 1, dataRowCount, numCols)
      .setNumberFormats(Array(dataRowCount).fill(fullFmtRow));

    // Sort by the first label column ascending, then clicks descending.
    // Ad Group sheets also sort by the second column (ad group name) ascending.
    const clicksCol   = def.header.indexOf('Clicks') + 1;
    const sortColumns = [{ column: 1, ascending: true }];
    if (def.header[1] === 'Ad Group') {
      sortColumns.push({ column: 2, ascending: true });
    }
    sortColumns.push({ column: clicksCol, ascending: false });
    sheet.getRange(4, 1, dataRowCount, numCols).sort(sortColumns);

    applyAlternatingRows(sheet, 4, dataRowCount, numCols);
    setDataFont(sheet, 4, dataRowCount, numCols, def.labelCount);
    addHeaderNotes(sheet, def.header);
    trimSheet(sheet, numCols, entry.nextRow - 1);
    applyFilter(sheet, def);

    Logger.log('         Sorted & formatted: ' + sheet.getName() + ' (' + dataRowCount.toLocaleString() + ' rows)');
  }
}


// =============================================================================
// AI SUMMARY
// =============================================================================
//
// HOW THIS WORKS -- A PLAIN-ENGLISH EXPLANATION
// -----------------------------------------------
// After the script finishes writing all the data sheets, this section takes
// over if you have provided an API key. Here is what happens, step by step:
//
// 1. READING THE DATA
//    The script reads the phrase data back from the Google Sheets it just wrote.
//    It does not re-process the original Google Ads data -- it simply reads the
//    finished output sheets, which is fast.
//
// 2. FORMATTING THE DATA FOR THE AI
//    The raw sheet data is converted into a compact text format (similar to a
//    CSV) that the AI can read efficiently. Only the columns that matter for
//    analysis are included: phrase, clicks, cost, conversions, conversion value,
//    and ROAS. The rest (impressions, CTR, etc.) are left out to keep the prompt
//    small and focused.
//
// 3. SENDING THE DATA TO GEMINI
//    The formatted data and the prompt instructions are sent to the Gemini Flash
//    API over HTTPS using Google Apps Script's built-in UrlFetchApp. This is
//    the same mechanism used by other Google Ads scripts that call external
//    services. The call takes a few seconds per request.
//
// 4. IN 'ACCOUNT' MODE (the default)
//    One API call is made. It contains the top N phrases by clicks from each
//    of the four account-level sheets (words, 2-grams, 3-grams, 4-grams).
//    The AI writes a single account-level summary.
//
// 5. IN 'FULL' MODE
//    One API call is made per campaign. Each call contains all the phrase data
//    for that campaign at both campaign level and ad group level.
//
//    For large campaigns (for example, a campaign with 100+ ad groups or very
//    high search volume), a single call might contain more data than is ideal
//    for the AI to work with. In that case, the script automatically splits the
//    campaign data into chunks that each stay within the token budget, sends
//    them as separate calls, and then stitches the responses together into one
//    campaign section in the output sheet.
//
//    You do not need to configure or think about this. The script measures each
//    campaign's data size automatically and handles chunking without any input
//    from you. It works the same way on every account regardless of structure.
//
// 6. WRITING THE OUTPUT
//    The AI's responses are written to a branded "AI Summary" sheet at the
//    front of the workbook. In 'full' mode, each campaign gets its own clearly
//    labelled section with a teal header row, so you can scroll through and
//    find specific campaigns quickly.
//
// 7. IF SOMETHING GOES WRONG
//    Any failure in the AI section is caught and logged without affecting the
//    data sheets. If the API call fails (wrong key, quota exceeded, network
//    issue), you will see an error message in the script logs and the data
//    sheets will be completely unaffected.
//
// WHAT A 'TOKEN' IS
// -----------------
// AI models measure text length in 'tokens' rather than characters. A token is
// roughly 4 characters or about 0.75 words. The token budget (AI_CONFIG.
// tokenBudget) controls how much data is sent per API call in 'full' mode.
// At 60,000 tokens per call, each call contains roughly 45,000 words of data --
// more than enough for a detailed ad group level analysis of any single campaign.
//
// =============================================================================


// -----------------------------------------------------------------------------
// countTokens(rows)
// Estimates the token count for a 2D array of row data.
// We use a 4-characters-per-token approximation, which is standard for English
// text mixed with numbers. This is used to decide whether a campaign's data
// fits in one API call or needs to be split into chunks.
// -----------------------------------------------------------------------------
function countTokens(rows) {
  var chars = 0;
  for (var i = 0; i < rows.length; i++) {
    for (var j = 0; j < rows[i].length; j++) {
      if (rows[i][j] !== null && rows[i][j] !== undefined) {
        chars += String(rows[i][j]).length;
      }
    }
  }
  return Math.ceil(chars / 4);
}


// -----------------------------------------------------------------------------
// formatCampaignRow(row)
// Converts a campaign-level sheet row into a compact CSV line.
// Campaign sheet column layout (0-indexed):
//   0  Campaign name  1  Phrase  2  Query Count  3  Clicks  4  Impressions
//   5  Cost  6  Conversions  7  Conv. Value  8..12  Calculated stats
// We send: phrase, clicks, cost, conversions, conv_value, roas.
// ROAS is recalculated from cost and conv_value rather than read from the sheet
// so it is always accurate regardless of how the sheet formatted it.
// -----------------------------------------------------------------------------
function formatCampaignRow(row) {
  var phrase = String(row[1]);
  var clicks = Math.round(parseFloat(row[3])  || 0);
  var cost   = Math.round((parseFloat(row[5]) || 0) * 100) / 100;
  var convs  = Math.round((parseFloat(row[6]) || 0) * 10)  / 10;
  var cv     = Math.round((parseFloat(row[7]) || 0) * 100) / 100;
  var roas   = cost > 0 ? Math.round((cv / cost) * 100) / 100 : 0;
  return phrase + ',' + clicks + ',' + cost + ',' + convs + ',' + cv + ',' + roas;
}


// -----------------------------------------------------------------------------
// formatAdGroupRow(row)
// Converts an ad group-level sheet row into a compact CSV line.
// Ad group sheet column layout (0-indexed):
//   0  Campaign  1  Ad Group  2  Phrase  3  Query Count  4  Clicks
//   5  Impressions  6  Cost  7  Conversions  8  Conv. Value  9..13  Calculated
// We send: ad_group, phrase, clicks, cost, conversions, conv_value, roas.
// -----------------------------------------------------------------------------
function formatAdGroupRow(row) {
  var adGroup = String(row[1]);
  var phrase  = String(row[2]);
  var clicks  = Math.round(parseFloat(row[4])  || 0);
  var cost    = Math.round((parseFloat(row[6]) || 0) * 100) / 100;
  var convs   = Math.round((parseFloat(row[7]) || 0) * 10)  / 10;
  var cv      = Math.round((parseFloat(row[8]) || 0) * 100) / 100;
  var roas    = cost > 0 ? Math.round((cv / cost) * 100) / 100 : 0;
  return adGroup + ',' + phrase + ',' + clicks + ',' + cost + ',' + convs + ',' + cv + ',' + roas;
}


// -----------------------------------------------------------------------------
// readSheetRowsByCampaign(ss, sheetNames, campaignColIndex)
// Reads the data rows from a list of sheets and groups them by campaign name.
// campaignColIndex is the 0-based column index of the campaign name field,
// which differs between campaign sheets (0) and ad group sheets (0).
// Returns a Map: campaignName -> array of raw row arrays.
// -----------------------------------------------------------------------------
function readSheetRowsByCampaign(ss, sheetNames, campaignColIndex) {
  var result = new Map();
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;
    var numCols = sheet.getLastColumn();
    var data    = sheet.getRange(4, 1, lastRow - 3, numCols).getValues();
    for (var r = 0; r < data.length; r++) {
      var campName = String(data[r][campaignColIndex]);
      if (!campName) continue;
      if (!result.has(campName)) result.set(campName, []);
      result.get(campName).push(data[r]);
    }
  }
  return result;
}


// -----------------------------------------------------------------------------
// readSheetRowsByCampaignAndAdGroup(ss, sheetNames)
// Like readSheetRowsByCampaign but groups into a nested Map:
//   campaignName -> adGroupName -> array of rows.
// Used for building ad group chunks in full mode.
// -----------------------------------------------------------------------------
function readSheetRowsByCampaignAndAdGroup(ss, sheetNames) {
  var result = new Map();
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;
    var numCols = sheet.getLastColumn();
    var data    = sheet.getRange(4, 1, lastRow - 3, numCols).getValues();
    for (var r = 0; r < data.length; r++) {
      var campName = String(data[r][0]);
      var agName   = String(data[r][1]);
      if (!campName || !agName) continue;
      if (!result.has(campName)) result.set(campName, new Map());
      if (!result.get(campName).has(agName)) result.get(campName).set(agName, []);
      result.get(campName).get(agName).push(data[r]);
    }
  }
  return result;
}


// -----------------------------------------------------------------------------
// calcCampaignAvgRoas(campRows)
// Calculates the average ROAS for a campaign from its raw campaign-level rows.
// Used to inject a benchmark into the campaign prompt so the AI can describe
// individual ad groups as above or below average.
// Returns a formatted string like "4.32" or "0" if cost is zero.
// Campaign cost is at index 5, conv value at index 7.
// -----------------------------------------------------------------------------
function calcCampaignAvgRoas(campRows) {
  var totalCost = 0;
  var totalCv   = 0;
  for (var i = 0; i < campRows.length; i++) {
    totalCost += parseFloat(campRows[i][5]) || 0;
    totalCv   += parseFloat(campRows[i][7]) || 0;
  }
  if (totalCost === 0) return '0';
  return String(Math.round((totalCv / totalCost) * 100) / 100);
}


// -----------------------------------------------------------------------------
// buildCampaignChunks(campaignName, campRows, agRowsByAdGroup)
// Determines how to split a campaign's data into API calls.
//
// The token budget (AI_CONFIG.tokenBudget) is the maximum number of tokens
// allowed per call. Campaign-level rows are always included in every chunk as
// context -- they tell the AI which campaign it is looking at and what the
// overall phrase landscape looks like. Ad group rows are then packed into
// chunks one ad group at a time until the budget is reached.
//
// If a single ad group's rows exceed the budget on their own (this happens with
// very large shopping or broad-match campaigns), that ad group is split at the
// phrase level -- rows are packed until the budget is hit, then a new chunk
// starts. The AI is told in the prompt that it is seeing a partial view.
//
// Returns an array of chunk objects, each with:
//   agRows   -- array of ad group rows to include in this chunk
//   chunkIdx -- 1-based index of this chunk within the campaign
//   total    -- total number of chunks for this campaign
// -----------------------------------------------------------------------------
function buildCampaignChunks(campaignName, campRows, agRowsByAdGroup) {
  var campTok   = countTokens(campRows);
  // Use a slightly tighter budget for the phrase-level fallback to ensure
  // that adding the campaign rows as context never pushes a chunk over the limit.
  var agBudget  = AI_CONFIG.tokenBudget - campTok - 200;

  // First pass: collect all chunks
  var chunks    = [];
  var curChunk  = [];
  var curTok    = 0;

  var agNames = Array.from(agRowsByAdGroup.keys());

  for (var i = 0; i < agNames.length; i++) {
    var agName  = agNames[i];
    var agRows  = agRowsByAdGroup.get(agName);
    var agTok   = countTokens(agRows);

    if (agTok > agBudget) {
      // This single ad group is too large for one chunk -- split at phrase level.
      // Flush whatever is in the current chunk first.
      if (curChunk.length > 0) {
        chunks.push(curChunk);
        curChunk = [];
        curTok   = 0;
      }
      // Now phrase-split this ad group.
      var phraseChunk = [];
      var phraseTok   = 0;
      for (var r = 0; r < agRows.length; r++) {
        var rowTok = countTokens([agRows[r]]);
        if (phraseChunk.length > 0 && phraseTok + rowTok > agBudget) {
          chunks.push(phraseChunk);
          phraseChunk = [];
          phraseTok   = 0;
        }
        phraseChunk.push(agRows[r]);
        phraseTok += rowTok;
      }
      if (phraseChunk.length > 0) chunks.push(phraseChunk);

    } else if (curTok + agTok > agBudget) {
      // Adding this ad group would exceed the budget -- flush current chunk.
      if (curChunk.length > 0) chunks.push(curChunk);
      curChunk = agRows.slice();
      curTok   = agTok;

    } else {
      // This ad group fits in the current chunk.
      curChunk = curChunk.concat(agRows);
      curTok  += agTok;
    }
  }
  if (curChunk.length > 0) chunks.push(curChunk);

  // Annotate each chunk with its index and total so the prompt can tell the
  // AI whether it is seeing the whole campaign or just a part of it.
  var result = [];
  for (var c = 0; c < chunks.length; c++) {
    result.push({ agRows: chunks[c], chunkIdx: c + 1, total: chunks.length });
  }
  return result;
}


// -----------------------------------------------------------------------------
// buildAccountPromptText(ss)
// Reads the four account-level sheets and builds the prompt for the account
// summary call. In 'account' mode only the top AI_CONFIG.topN rows are used.
// In 'full' mode all rows are sent (topN is ignored for this call).
// -----------------------------------------------------------------------------
function buildAccountPromptText(ss) {
  var sheetNames = [
    'Account Word Analysis',
    'Account 2-Gram Analysis',
    'Account 3-Gram Analysis',
    'Account 4-Gram Analysis',
  ];

  var isFullMode   = (AI_CONFIG.mode === 'full');
  var dataSections = [];

  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    var numCols = sheet.getLastColumn();
    var data    = sheet.getRange(4, 1, lastRow - 3, numCols).getValues();

    // Sort by clicks descending (column index 2 on account sheets).
    var sorted = data.slice().sort(function(a, b) {
      return (parseFloat(b[2]) || 0) - (parseFloat(a[2]) || 0);
    });

    var limit = isFullMode ? sorted.length : Math.min(AI_CONFIG.topN, sorted.length);
    var lines = ['phrase,clicks,cost,conversions,conv_value,roas'];

    for (var r = 0; r < limit; r++) {
      var row    = sorted[r];
      var phrase = String(row[0]);
      var clicks = Math.round(parseFloat(row[2])  || 0);
      var cost   = Math.round((parseFloat(row[4]) || 0) * 100) / 100;
      var convs  = Math.round((parseFloat(row[5]) || 0) * 10)  / 10;
      var cv     = Math.round((parseFloat(row[6]) || 0) * 100) / 100;
      var roas   = cost > 0 ? Math.round((cv / cost) * 100) / 100 : 0;
      lines.push(phrase + ',' + clicks + ',' + cost + ',' + convs + ',' + cv + ',' + roas);
    }

    var label = isFullMode
      ? '--- ' + sheetNames[i] + ' (all ' + limit + ' rows) ---'
      : '--- ' + sheetNames[i] + ' (top ' + limit + ' by clicks) ---';
    dataSections.push(label);
    dataSections.push(lines.join('\n'));
  }

  return AI_CONFIG.accountPrompt + '\n\n' + dataSections.join('\n');
}


// -----------------------------------------------------------------------------
// buildCampaignPromptText(campaignName, campRows, agChunkRows, chunkIdx, total, avgRoas)
// Builds the prompt text for a single campaign API call.
// campRows is the full set of campaign-level rows for this campaign (always
// included for context). agChunkRows is the ad group rows for this specific
// chunk. chunkIdx and total are used to tell the AI if it is seeing a partial
// view of a large campaign.
// -----------------------------------------------------------------------------
function buildCampaignPromptText(campaignName, campRows, agChunkRows, chunkIdx, total, avgRoas) {
  var chunkNote = total > 1
    ? '(part ' + chunkIdx + ' of ' + total + ' -- this is a large campaign split across multiple analyses)'
    : '';

  var instruction = AI_CONFIG.campaignPrompt
    .replace('{CAMPAIGN}',  campaignName)
    .replace('{CHUNK_NOTE}', chunkNote)
    .replace('{AVG_ROAS}',  avgRoas);

  // Campaign-level section
  var campLines = ['--- Campaign-level phrase data ---',
                   'phrase,clicks,cost,conversions,conv_value,roas'];
  for (var i = 0; i < campRows.length; i++) {
    campLines.push(formatCampaignRow(campRows[i]));
  }

  // Ad group section
  var agLines = ['--- Ad group phrase data ---',
                 'ad_group,phrase,clicks,cost,conversions,conv_value,roas'];
  for (var j = 0; j < agChunkRows.length; j++) {
    agLines.push(formatAdGroupRow(agChunkRows[j]));
  }

  return instruction + '\n\n' + campLines.join('\n') + '\n\n' + agLines.join('\n');
}


// -----------------------------------------------------------------------------
// callGemini(promptText)
// Sends promptText to the Gemini Flash API and returns the response text.
//
// We use UrlFetchApp rather than any SDK because the Ads Scripts runtime does
// not support npm packages. The Gemini generateContent endpoint accepts a plain
// JSON body and returns a straightforward JSON response.
//
// Model: gemini-flash-latest -- the 'latest' alias always resolves to the
// current production Flash model across generations. Google provides a
// two-week email notice before any swap, so this string will not break
// silently due to deprecation the way a version-pinned string can. It is
// deliberately not pinned to a specific generation (1.5, 2.5, 3 etc.) so
// the script keeps working as new Flash generations are released.
//
// temperature is set to 0.3 (low) for consistent, factual output rather than
// creative variation. maxOutputTokens is 8,192 -- generous enough for a
// detailed campaign analysis and well within the output limits of current
// Flash models.
//
// muteHttpExceptions is true so we can handle error responses ourselves and
// surface a useful message rather than a raw HTTP error.
// -----------------------------------------------------------------------------
function callGemini(promptText) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + AI_CONFIG.apiKey;

  var payload = JSON.stringify({
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: {
      temperature:     0.3,
      maxOutputTokens: 8192,
    },
  });

  var options = {
    method:             'post',
    contentType:        'application/json',
    payload:            payload,
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();
  var body     = JSON.parse(response.getContentText());

  if (code !== 200) {
    var errMsg = (body.error && body.error.message) ? body.error.message : 'HTTP ' + code;
    throw new Error('Gemini API error: ' + errMsg);
  }

  if (!body.candidates || !body.candidates[0] ||
      !body.candidates[0].content || !body.candidates[0].content.parts ||
      !body.candidates[0].content.parts[0]) {
    throw new Error('Gemini API returned an unexpected response structure: ' + JSON.stringify(body));
  }

  return body.candidates[0].content.parts[0].text;
}


// -----------------------------------------------------------------------------
// initialiseSummarySheet(ss, filterText)
// Creates (or clears) the AI Summary sheet, applies brand styling to the three
// header rows, moves the sheet to the front of the workbook, and returns it.
// The body content (account summary and campaign sections) is written
// separately by the calling function as it arrives from the API.
// -----------------------------------------------------------------------------
function initialiseSummarySheet(ss, filterText) {
  var sheetName = AI_CONFIG.sheetName;
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  // Move to position 0 (front of the workbook).
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(0);

  // A single wide column is cleaner than the previous two-column approach.
  // Column B served no purpose other than contributing to the merged cell
  // width, which we achieve more simply by just widening column A.
  var numCols = 1;
  sheet.setFrozenRows(3);

  var r1 = sheet.getRange(1, 1, 1, numCols);
  r1.setValue('WeDiscover  |  AI Summary  |  Powered by Google Gemini');
  r1.setBackground(WD_CRIMSON);
  r1.setFontColor(WD_WHITE);
  r1.setFontFamily(WD_FONT);
  r1.setFontSize(12);
  r1.setFontWeight('bold');
  r1.setVerticalAlignment('middle');
  r1.setHorizontalAlignment('left');
  sheet.setRowHeight(1, 40);

  var modeLabel = AI_CONFIG.mode === 'full'
    ? 'Full mode -- all campaigns analysed at ad group level'
    : 'Account mode -- top ' + AI_CONFIG.topN + ' phrases per account-level sheet';

  var r2 = sheet.getRange(2, 1, 1, numCols);
  r2.setValue(filterText + '  |  ' + modeLabel);
  r2.setBackground(WD_TEAL);
  r2.setFontColor(WD_NAVY);
  r2.setFontFamily(WD_FONT);
  r2.setFontSize(9);
  r2.setFontWeight('bold');
  r2.setVerticalAlignment('middle');
  r2.setHorizontalAlignment('left');
  sheet.setRowHeight(2, 24);

  var r3 = sheet.getRange(3, 1, 1, numCols);
  r3.setValue('Generated by Gemini Flash. Always review AI recommendations before acting on them.');
  r3.setBackground(WD_NAVY);
  r3.setFontColor(WD_WHITE);
  r3.setFontFamily(WD_FONT);
  r3.setFontSize(9);
  r3.setFontWeight('bold');
  r3.setVerticalAlignment('middle');
  r3.setHorizontalAlignment('left');
  sheet.setRowHeight(3, 28);

  sheet.setTabColor(WD_CRIMSON);
  sheet.setColumnWidth(1, 660);

  return sheet;
}


// -----------------------------------------------------------------------------
// normaliseAiResponse(text)
// Post-processes the raw text returned by Gemini before writing to the sheet.
//
// Two normalisation rules:
//   1. 'roas' and 'Roas' -> 'ROAS' everywhere. The model frequently ignores
//      the capitalisation instruction so we enforce it programmatically.
//   2. Bullet hyphens ('- ') at the start of lines are replaced with the
//      Unicode bullet character (u2022) and a small indent, so bullets look
//      clean in the cell rather than starting with a hyphen character.
// -----------------------------------------------------------------------------
function normaliseAiResponse(text) {
  // Capitalise all case variants of roas that are not already fully uppercased.
  // The regex matches 'roas' case-insensitively but excludes the all-caps form
  // so we don't double-process.
  text = text.replace(/\broas\b/gi, function(match) {
    return match === 'ROAS' ? match : 'ROAS';
  });

  // Replace leading '- ' bullet hyphens with a bullet character + indent.
  // We normalise line endings first for safety.
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  text = text.replace(/^- /gm, '  \u2022 ');

  return text;
}


// -----------------------------------------------------------------------------
// appendSectionBlock(sheet, nextRow, headerText, bodyText)
// Writes one section of the AI summary to the sheet:
//   - A teal header row (campaign name, "ACCOUNT SUMMARY" etc.)
//   - A single body cell containing the full section text with embedded
//     line breaks, so it reads as a coherent block rather than fragments.
//
// Why a single cell rather than one row per line?
// One row per line produces 100+ rows per run, makes the sheet hard to
// navigate, and means the row height estimation logic has to fight Google
// Sheets' tendency to clip wrapped text. A single wrapped cell with an
// estimated height is simpler, more readable, and more reliable.
//
// Row height is estimated from the total character count divided by an
// assumed characters-per-line value. The estimate is intentionally generous
// (we would rather have a slightly tall cell than a clipped one).
//
// Returns the new nextRow value after writing both rows.
// -----------------------------------------------------------------------------
function appendSectionBlock(sheet, nextRow, headerText, bodyText) {
  var numCols = 1;

  // -- Teal section header row ------------------------------------------------
  var hRange = sheet.getRange(nextRow, 1, 1, numCols);
  hRange.setValue(headerText);
  hRange.setWrap(false);
  hRange.setVerticalAlignment('middle');
  hRange.setHorizontalAlignment('left');
  hRange.setBackground(WD_TEAL);
  hRange.setFontColor(WD_NAVY);
  hRange.setFontFamily(WD_FONT);
  hRange.setFontSize(9);
  hRange.setFontWeight('bold');
  sheet.setRowHeight(nextRow, 28);
  nextRow++;

  // -- Body cell with full section text ---------------------------------------
  // Normalise line endings and trim trailing whitespace so the cell does not
  // start with a blank line.
  var cleanBody = normaliseAiResponse(bodyText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim());

  var bRange = sheet.getRange(nextRow, 1, 1, numCols);
  bRange.setValue(cleanBody);
  bRange.setWrap(true);
  bRange.setVerticalAlignment('top');
  bRange.setHorizontalAlignment('left');
  bRange.setBackground(WD_WHITE);
  bRange.setFontColor(WD_NAVY);
  bRange.setFontFamily('Arial');
  bRange.setFontSize(10);
  bRange.setFontWeight('normal');

  // Estimate row height. At Arial 10pt in a 660px column, roughly 95
  // characters fit on one line at 18px per line. We count non-empty lines
  // (including wrapped ones) and add generous padding. The aim is never to
  // clip -- a slightly tall cell is fine, a clipped one is not.
  var CHARS_PER_LINE = 95;
  var LINE_HEIGHT_PX = 18;
  var PADDING_PX     = 20;

  var bodyLines = cleanBody.split('\n');
  var wrappedTotal = 0;
  for (var i = 0; i < bodyLines.length; i++) {
    wrappedTotal += Math.max(1, Math.ceil(bodyLines[i].length / CHARS_PER_LINE));
  }
  var rowPx = Math.max(60, wrappedTotal * LINE_HEIGHT_PX + PADDING_PX);
  sheet.setRowHeight(nextRow, rowPx);
  nextRow++;

  return nextRow;
}


// -----------------------------------------------------------------------------
// finaliseSummarySheet(sheet, usedRows)
// Trims unused rows and columns from the summary sheet after all content has
// been written. Called once at the very end of generateAiSummary().
// -----------------------------------------------------------------------------
function finaliseSummarySheet(sheet, usedRows) {
  var numCols = 1;
  var maxCols = sheet.getMaxColumns();
  if (maxCols > numCols) sheet.deleteColumns(numCols + 1, maxCols - numCols);
  var maxRows = sheet.getMaxRows();
  if (maxRows > usedRows) sheet.deleteRows(usedRows + 1, maxRows - usedRows);
  SpreadsheetApp.flush();
}


// -----------------------------------------------------------------------------
// generateAiSummary(ss, mode)
// Orchestrates the full AI summary pipeline for both 'account' and 'full' modes.
//
// Account mode:
//   1. Read account-level sheets and build a single prompt.
//   2. Call Gemini once.
//   3. Write the response to the summary sheet.
//
// Full mode:
//   1. Write the account summary (same as account mode).
//   2. For each campaign, load its campaign-level and ad group rows.
//   3. Calculate token count. If the campaign fits in one call, send it.
//      If not, split into chunks automatically (see buildCampaignChunks).
//   4. For multi-chunk campaigns, collect all chunk responses and combine them
//      under one campaign header in the sheet.
//   5. Log progress after each campaign so you can see how far along it is.
// -----------------------------------------------------------------------------
function generateAiSummary(ss, mode) {
  var filterText  = buildFilterText();
  var sheet       = initialiseSummarySheet(ss, filterText);
  var nextRow     = 4;

  // ---- Account-level summary (always written regardless of mode) -------------
  Logger.log('[ AI ] Writing account summary...');
  var accountPromptText = buildAccountPromptText(ss);
  var accountSummary    = callGemini(accountPromptText);

  nextRow = appendSectionBlock(sheet, nextRow, 'ACCOUNT SUMMARY', accountSummary);

  if (mode !== 'full') {
    finaliseSummarySheet(sheet, nextRow - 1);
    return;
  }

  // ---- Full mode: per-campaign analysis --------------------------------------
  var campSheetNames = [
    'Campaign Word Analysis',
    'Campaign 2-Gram Analysis',
    'Campaign 3-Gram Analysis',
    'Campaign 4-Gram Analysis',
  ];
  var agSheetNames = [
    'Ad Group Word Analysis',
    'Ad Group 2-Gram Analysis',
    'Ad Group 3-Gram Analysis',
    'Ad Group 4-Gram Analysis',
  ];

  // Read all campaign and ad group data grouped by campaign name.
  // readSheetRowsByCampaign returns Map: campaignName -> [rows]
  // readSheetRowsByCampaignAndAdGroup returns Map: campaignName -> Map: adGroupName -> [rows]
  var campRowsByCamp = readSheetRowsByCampaign(ss, campSheetNames, 0);
  var agRowsByCampAg = readSheetRowsByCampaignAndAdGroup(ss, agSheetNames);

  var campNames  = Array.from(campRowsByCamp.keys()).sort();
  var totalCamps = campNames.length;

  Logger.log('[ AI ] ' + totalCamps + ' campaigns to analyse in full mode...');

  for (var ci = 0; ci < campNames.length; ci++) {
    var campName    = campNames[ci];
    var campRows    = campRowsByCamp.get(campName);
    var agByAg      = agRowsByCampAg.get(campName) || new Map();
    var avgRoas     = calcCampaignAvgRoas(campRows);

    Logger.log('[ AI ] Campaign ' + (ci + 1) + '/' + totalCamps + ': ' + campName);

    var chunks     = buildCampaignChunks(campName, campRows, agByAg);
    var campParts  = [];

    for (var ki = 0; ki < chunks.length; ki++) {
      var chunk       = chunks[ki];
      var promptText  = buildCampaignPromptText(
        campName, campRows, chunk.agRows, chunk.chunkIdx, chunk.total, avgRoas
      );
      var chunkResult = callGemini(promptText);
      campParts.push(chunkResult);

      if (chunks.length > 1) {
        Logger.log('[ AI ]   Chunk ' + chunk.chunkIdx + '/' + chunk.total + ' done.');
      }
    }

    // Write a teal campaign header row then the combined response.
    var headerText   = campName + '  |  Avg ROAS: ' + avgRoas;
    var responseText = campParts.join('\n\n--- (continued) ---\n\n');

    nextRow = appendSectionBlock(sheet, nextRow, headerText, responseText);
  }

  finaliseSummarySheet(sheet, nextRow - 1);
}
