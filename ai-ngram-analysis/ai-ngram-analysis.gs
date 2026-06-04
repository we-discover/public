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
 *  Version:     3.13.0
 *  Released:    2026-05-11
 *  Contact:     scripts@we-discover.com
 *
 *  Credits:     Original n-gram concept and structure by Brainlabs Digital
 *               (https://github.com/Brainlabs-Digital/Google-Ads-Scripts).
 *               GAQL migration by Nils Rooijmans (2022, 2025).
 *               Shared set access fix by Arjan Schoorl / Flowboost (2025).
 *               Rewritten for ES6, GAQL and scale by WeDiscover (2026).
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
  startDate: '01/05/2026',   // dd/mm/yyyy -- change this to your start date
  endDate:   '31/05/2026',   // dd/mm/yyyy -- change this to your end date

  // -- Currency ----------------------------------------------------------------
  // Type your currency symbol directly. £ for GBP, $ for USD, € for EUR.
  currencySymbol: '£',

  // -- Campaign filter ---------------------------------------------------------
  // If you only want to look at certain campaigns, type part of their name here.
  // Leave both fields empty to include all campaigns.
  campaignNameContains:       '',  // only include campaigns whose name contains this
  campaignNameDoesNotContain: '',  // exclude campaigns whose name contains this

  // -- Paused campaigns and ad groups ------------------------------------------
  // Set to true to skip paused campaigns or ad groups.
  ignorePausedCampaigns: true,
  ignorePausedAdGroups:  true,

  // -- Negative keywords -------------------------------------------------------
  // Set to true to filter out queries already blocked by your negative keywords.
  // Gives a cleaner picture of what is actually reaching your ads.
  checkNegatives: true,

  // -- Spreadsheet -------------------------------------------------------------
  // Paste the full URL of the Google Sheet you want results written to.
  // The sheet must already exist and this script must have edit access.
  spreadsheetUrl: 'INSERT-SPREADSHEET-URL-HERE',

  // -- N-gram length -----------------------------------------------------------
  // An n-gram is a sequence of words. 1 = single words, 2 = two-word phrases, etc.
  // On high-traffic accounts, keep maxNGramLength at 4 or below.
  minNGramLength: 1,
  maxNGramLength: 4,

  // -- Clear the spreadsheet on each run ---------------------------------------
  // Set to true to wipe the sheet clean before writing new results.
  clearSpreadsheet: true,

  // -- Thresholds --------------------------------------------------------------
  // Phrases that do not meet all of these minimums are excluded from results.
  // On big accounts, a clicks threshold of 50 or higher keeps the output focused.
  thresholds: {
    queryCount:   0,   // minimum number of search queries containing the phrase
    impressions:  0,   // minimum impressions
    clicks:       50,  // minimum clicks
    cost:         0,   // minimum spend
    conversions:  0,   // minimum conversions
  },

  // -- Time limit --------------------------------------------------------------
  // Controls when the script stops mining and moves on to sheet finalisation
  // and the AI summary. The full pipeline is:
  //
  //   mining  →  finaliseSheets (dedup + sort + format)  →  AI summary
  //
  // The script targets a total runtime of 25 minutes to stay well clear of
  // Google's 30-minute hard kill. maxExecutionMinutes controls when mining
  // stops; the AI section then uses whatever time remains up to that 25-minute
  // ceiling. The AI time budget is calculated automatically -- you do not need
  // to set it anywhere.
  //
  // Rule of thumb for this setting:
  //   Without AI:             set to 25.
  //   With AI, account mode:  set to 20.
  //   With AI, full mode:     set to 15. Lower (e.g. 12) on very large accounts.
  //
  // If mining hits this limit before finishing, the sheet will be marked
  // PARTIAL but all data processed up to that point is written correctly.
  maxExecutionMinutes: 25,
};
// =============================================================================
// END OF CONFIGURATION -- do not edit anything below this line
// =============================================================================


// =============================================================================
// AI CONFIGURATION (OPTIONAL)
//
// ---- HOW TO GET A FREE API KEY (takes about two minutes) -------------------
//   1. Go to https://aistudio.google.com in any browser.
//   2. Sign in with any Google account.
//   3. Click "Get API key" in the left-hand navigation panel.
//   4. Click "Create API key", select any project, then click "Create".
//   5. Copy the key and paste it into the apiKey field below.
//
//   The free tier gives you 1,500 requests per day at no cost. No credit card
//   required. Even a heavy 'full' mode run uses around 40-50 requests.
//
// ---- TWO MODES: 'account' vs 'full' ----------------------------------------
//   mode: 'account'
//     One API call. Sends the top N phrases by clicks from each of the four
//     account-level sheets. Fast, good for a quick read of the whole account.
//
//   mode: 'full'
//     One API call per campaign using all phrase data at campaign and ad group
//     level. Produces specific, actionable recommendations per campaign.
//     For large campaigns the script splits data into chunks automatically and
//     combines the responses -- no configuration needed.
// =============================================================================
const AI_CONFIG = {

  // -- API key -----------------------------------------------------------------
  // Paste your Google AI Studio key here. Leave blank ('') to skip AI entirely.
  apiKey: '',

  // -- Mode --------------------------------------------------------------------
  // 'account' -- one call, top N phrases per account-level sheet. Fast.
  // 'full'    -- one call per campaign, all data at campaign and ad group level.
  mode: 'full',

  // -- Token budget (full mode only) -------------------------------------------
  // Maximum tokens of input data per API call. 60,000 is a good balance between
  // giving the AI enough context and keeping each call fast. No need to change
  // this unless you are hitting timeout issues, in which case try 40,000.
  tokenBudget: 60000,

  // -- Number of phrases per sheet ---------------------------------------------
  // Maximum rows sent to the AI per campaign-level and per ad-group-level sheet.
  // Rows are selected by clicks descending, so the AI always sees the most
  // impactful phrases first.
  //
  // In account mode: limits rows from each account-level sheet.
  // In full mode: limits rows per campaign and per individual ad group.
  //
  // 100-150 is a good default for full mode. The AI does not improve with more
  // rows -- beyond ~200 the prompt becomes too diffuse and quality drops.
  topN: 100,

  // -- Maximum chunks per campaign (full mode only) ----------------------------
  // Hard cap on the number of API calls made per campaign. When the natural
  // chunk count exceeds this limit, the script merges ad group rows into fewer,
  // larger chunks rather than making more calls.
  //
  // 3 gives the AI enough context for a thorough analysis without spending
  // excessive time on any single campaign. Set to 1 for fastest results.
  maxChunksPerCampaign: 3,

  // -- AI time budget ----------------------------------------------------------
  // Calculated automatically from CONFIG.maxExecutionMinutes. The script targets
  // a total runtime of 25 minutes; the AI section gets whatever is left after
  // mining and sheet finalisation (typically 5-10 minutes on a 15-minute mine).
  // You do not need to change this.
  maxAiMinutes: (30 - CONFIG.maxExecutionMinutes),

  // -- Model ------------------------------------------------------------------
  // The Gemini model used for all AI calls.
  //
  // Free-tier limits observed (May 2026) -- these vary by account and change
  // without notice, so check yours at https://aistudio.google.com/rate-limit:
  //
  //   gemini-3.1-flash-lite   15 RPM   500 RPD   <-- recommended default
  //   gemini-2.5-flash-lite   10 RPM    20 RPD
  //   gemini-3-flash           5 RPM    20 RPD
  //   gemini-2.5-flash         5 RPM    20 RPD
  //
  // RPD (requests per day) is the binding constraint on the free tier.
  // At 20 RPD you can exhaust your daily quota in a single full-mode run.
  // gemini-3.1-flash-lite's 500 RPD means you can run across multiple accounts
  // or re-run the same account without burning through the day's allowance.
  //
  // To switch models, change this string. Nothing else needs updating.
  // Full model strings: https://ai.google.dev/gemini-api/docs/models
  model: 'gemini-3.1-flash-lite',

  // -- Delay between API calls -------------------------------------------------
  // Milliseconds to pause after each Gemini call. This keeps requests well
  // within the free-tier rate limit window regardless of which model is in use.
  // Google does not publish free-tier RPM numbers in their docs -- they vary by
  // model and can change without notice. The observed limit for gemini-3.1-flash-lite
  // is 15 RPM, which works out to one request every 4 seconds minimum.
  //
  // 4,500ms gives a comfortable margin below that window. On a 20-campaign run
  // this adds roughly 90 seconds to the AI section, which is well within the
  // time budget. Set to 0 if you are on a paid tier with higher RPM limits.
  callDelayMs: 4500,

  // -- Output sheet name -------------------------------------------------------
  sheetName: 'AI Summary',

  // -- Account summary prompt --------------------------------------------------
  // The instruction sent to the AI for the account-level summary. The phrase
  // data is appended automatically. Keep this under a few hundred words.
accountPrompt: [
    'You are a Lead Performance Consultant at WeDiscover, a London-based performance marketing agency. You are auditing paid search n-gram data on behalf of a client. The client\'s brand name is: {CLIENT_BRAND}.',
    '',
    'WeDiscover manages the account -- WeDiscover is NOT the brand being advertised. Always refer to the client\'s customers as "{CLIENT_BRAND}\'s customers", never as "WeDiscover\'s customers".',
    '',
    '1. UNDERSTANDING THE DATA:',
    'CRITICAL -- read this carefully before writing anything.',
    'The data contains N-GRAMS: words and phrases extracted from search queries, NOT keywords.',
    'An n-gram like \'birthday gifts for her\' is a fragment that appeared inside one or more search queries.',
    'It does NOT represent a keyword that exists in the account.',
    'It tells you what THEMES and MODIFIERS appear in the searches that triggered the ads.',
    '',
    'THEREFORE:',
    '- NEVER recommend adding an n-gram directly as a keyword. Instead, describe the THEME it represents.',
    '- NEVER recommend match types for n-grams. Match types apply to keywords, not to search query fragments.',
    '- When identifying keyword opportunities, describe the intent or theme the n-gram reveals, and suggest what TYPE of keyword could be built to capture or exclude that intent.',
    '- When suggesting negatives, recommend the concept or theme, not the exact n-gram string.',
    '- EXCEPTION: Section 5 is a factual spend callout and DOES name exact terms verbatim. This is an observation, not a keyword recommendation, so the "describe the theme not the string" rule does not apply there.',
    '',
    '2. TONE & STYLE:',
    '- Be professional, insightful, and collaborative. Use \'We suggest\' or \'Consider\' rather than \'Do this\'.',
    '- Maintain dry British wit, but keep it constructive.',
    '- Use British English (e.g. \'optimise\', \'personalised\').',
    `- ALWAYS use ${CONFIG.currencySymbol} signs for spend and value.`,
    '- Format ROAS in ALL CAPS.',
    '- Wrap every n-gram or phrase in \'single quotes\'.',
    '',
    '3. JUDGING PERFORMANCE:',
    '- Judge efficiency relative to what the account is actually achieving overall, not against any fixed threshold.',
    '- A term with a low absolute ROAS is only wasteful if it is also low relative to its campaign context.',
    '- A term with a ROAS that is below average but still profitable (positive return) should be noted as underperforming, not written off as waste.',
    '',
    '4. STRUCTURE:',
    'Use the exact numbered section headers below. Put a blank line between each section.',
    '',
    '1. PROFITABLE THEMES',
    'Identify the 3-4 strongest themes across the account. For each, name the theme, give spend and ROAS, and say which type of campaign it is most relevant to (brand, shopping, or generic search). Be specific.',
    '',
    '2. BIGGEST OPPORTUNITIES',
    'Based on the highest-performing n-gram themes, identify where the account could do more. Rules:',
    '- Name the theme, its ROAS, and the specific reason it represents an opportunity.',
    '- Do NOT suggest creating ad groups, campaigns, landing pages, keyword sets, or product groups -- you have no visibility of what already exists.',
    '- Do NOT suggest building out themes, expanding coverage, or increasing investment -- these are generic, not insights.',
    '- Do NOT recommend any action at all. This section identifies WHAT is working and WHY, based purely on the data. The reader decides what to do with it.',
    '- Every sentence must reference a specific n-gram and its actual ROAS figure from the data.',
    '',
    '3. BIGGEST PROBLEMS',
    'Identify the 3-4 themes with the largest combination of high spend AND low ROAS. Rules:',
    '- Order by actual £ spend descending.',
    '- For each, state spend, revenue, and ROAS.',
    '- Do not flag terms spending under £20.',
    '- Only flag a theme if its ROAS is more than 15% below the account average across all non-brand campaigns. Do not flag terms that are marginally below average -- only flag genuine, material underperformance.',
    '- Do not calculate hypothetical benchmark-adjusted figures. State actual numbers only.',
    '- Do NOT list, mention, or annotate themes that were considered but excluded.',
    '',
    '4. SUMMARY',
    'One senior strategic observation that ties the above together. Be direct. If the account has a clear structural problem, name it.',
    '',
    '5. HIGHEST-SPEND ZERO-CONVERTING TERMS',
    'A short, factual callout of where spend is going nowhere. Identify the single highest-spending term with ZERO conversions at two lengths: the top one-word (unigram) term and the top two-word (bigram) term. Rules:',
    '- Name the exact term verbatim in \'single quotes\' and state its actual ${} spend for each length.',
    '- These have a ROAS of 0.00 by definition -- state plainly that the spend returned nothing. Do not dress it up.',
    '- This is an OBSERVATION, not a recommendation. Do not tell the reader what to do about it.',
    '- When identifying the top unigram, IGNORE single-word stop words (\'for\', \'and\', \'the\', \'of\', \'in\', \'a\', \'to\', \'with\', \'uk\') -- they carry no intent and are extraction artefacts.',
    '- Do not flag terms spending under £20. If no zero-converting term clears £20 at a given length, say so plainly for that length.',
    '- Keep this to one or two sentences. Example shape: \'The highest-spending non-converting terms are \'recycling\' (one-word, £X) and \'recycling bin\' (two-word, £Y) -- neither has driven a single conversion.\'',
    '',
    'Keep the response under 500 words. Do not use markdown, asterisks, or hash symbols.',
  ].join('\n'),

  // -- Campaign prompt (full mode only) ----------------------------------------
  // Placeholders filled in automatically before each call:
  //   {CLIENT_BRAND} -- the client's brand name, detected from ad final URLs
  //   {CAMPAIGN}     -- the campaign name
  //   {CHUNK_NOTE}   -- blank for single-chunk campaigns; "(part 2 of 4)" for splits
  //   {AVG_ROAS}     -- the average ROAS for this campaign
  campaignPrompt: [
    'You are a Senior PPC Analyst at WeDiscover, a London-based performance marketing agency. You are auditing \'{CAMPAIGN}\' (Avg ROAS: {AVG_ROAS}) {CHUNK_NOTE} on behalf of the client {CLIENT_BRAND}.',
    '',
    'WeDiscover manages the account -- WeDiscover is NOT the brand being advertised. Always refer to the client\'s customers as "{CLIENT_BRAND}\'s customers", never as "WeDiscover\'s customers".',
    '',
    '1. UNDERSTANDING THE DATA:',
    'CRITICAL -- read this carefully before writing anything.',
    'The data contains N-GRAMS: words and phrases extracted from search queries, NOT keywords.',
    'An n-gram like \'birthday gifts\' appeared inside searches that triggered this campaign\'s ads.',
    'It does NOT represent a keyword that exists in the account.',
    '',
    'THEREFORE:',
    '- NEVER recommend adding an n-gram directly as a keyword or as a negative keyword.',
    '- NEVER recommend match types for n-grams.',
    '- NEVER recommend creating, reviewing, or restructuring ad groups -- you have no visibility of the existing account structure.',
    '- NEVER make assumptions about current ad copy, landing pages, or bidding strategy -- you can only see what the n-gram data shows.',
    '- IGNORE single-word stop words such as \'for\', \'and\', \'the\', \'of\', \'in\', \'a\', \'to\', \'with\', \'uk\' when they appear alone as single-word n-grams. These carry no intent signal and are artefacts of the n-gram extraction process. Only use them when they appear as part of a longer phrase (e.g. \'gifts for him\').',
    '- When suggesting negatives, describe the THEME or INTENT to block, not the exact n-gram string.',
    '- When suggesting keyword expansion, describe what the n-gram reveals about user intent.',
    '',
    '2. CAMPAIGN CONTEXT:',
    'The average ROAS for this campaign is {AVG_ROAS}.',
    'Use this as the benchmark for all efficiency judgements -- not any fixed threshold.',
    'A term is only wasteful if its ROAS is significantly below {AVG_ROAS} AND it is spending material budget. Do not flag terms that are mildly below average if the absolute spend is small or the ROAS is still positive.',
    '',
    'Campaign type rules -- check the campaign name:',
    '- Contains \'Brand\' or \'EX\': This is a brand campaign. You MUST rename Section 3 to \'LOWER-EFFICIENCY SEGMENTS\' (not DRAINAGE). Only include themes in this section where ROAS is below 3.0 AND spend is material. If a theme has a ROAS above 3.0 -- even if it is below the campaign average -- it must not appear in this section at all, not even with a note. A term returning 5x, 8x, or 12x ROAS is profitable and is not a problem. Suggested actions for brand campaigns should focus only on query-level refinement for genuine inefficiencies (e.g. misspellings, abbreviations with notably lower ROAS) -- not on excluding profitable brand-adjacent traffic.',
    '- Contains \'DSA\': This is a Dynamic Search Ads campaign. There are no keywords to add or adjust. Suggestions must focus on search query THEMES to exclude via negatives. Do NOT suggest reviewing or improving the product feed or landing pages -- you cannot see these. Base all suggestions solely on which search query themes are converting and which are not.',
    '- Contains \'PLA\' or \'Shopping\': Focus on search query themes and product group exclusions only.',
    '',
    '3. WASTE CALCULATION:',
    'When calculating waste in the DRAINAGE section, report actual spend and actual revenue only.',
    'Do NOT calculate a hypothetical \'revenue at benchmark ROAS\' figure -- this implies a certainty of return that does not exist.',
    'Instead, state: spend, actual ROAS, and actual revenue. The reader can judge the gap themselves.',
    'Example: \'birthday\' spent £31.54 and returned £9.78 (ROAS 0.31) against a campaign average of {AVG_ROAS}.',
    '',
    '4. TONE & STYLE:',
    '- Use British English throughout.',
    `- ALWAYS use ${CONFIG.currencySymbol} signs for currency values.`,
    '- Every n-gram or phrase mentioned MUST be wrapped in \'single quotes\'.',
    '- Write ROAS in ALL CAPS.',
    '- Use \'We suggest\' or \'Consider\' rather than \'Do this\'.',
    '',
    '5. STRUCTURE:',
    'Use the exact numbered section headers below. Put a blank line between each section.',
    '',
    '1. PERFORMANCE OVERVIEW',
    'How is the campaign performing against the {AVG_ROAS} ROAS benchmark overall? One short paragraph.',
    '',
    '2. WINNING SEGMENTS',
    'Which n-gram themes are driving the strongest performance? What do they reveal about customer intent?',
    '',
    '3. DRAINAGE',
    'Which themes are spending material budget significantly below the {AVG_ROAS} benchmark? Rules:',
    '- ONLY include themes where ROAS is strictly below {AVG_ROAS}. If a theme has ROAS at or above {AVG_ROAS}, it must not appear here, even with a qualifying note.',
    '- ONLY include themes where the ROAS gap from benchmark is at least 15% (e.g. on a 1.00 benchmark, only flag themes at 0.85 or below; on a 0.50 benchmark, only flag themes at 0.43 or below). Marginal underperformance is noise, not insight.',
    '- Order by actual £ spend descending so the biggest problems appear first.',
    '- For each, state: theme, spend, revenue, and ROAS.',
    '- Do not flag themes where spend is under £10.',
    '- Do not calculate hypothetical benchmark-adjusted waste figures.',
    '- If fewer than 3 distinct themes qualify after applying all rules above, write: \'No material drainage themes identified in this data set.\' Do not pad with marginal or repeated examples.',
    '- Do NOT list, mention, or annotate themes that were considered but excluded from this section. Only show themes that qualify. The reader does not need to see your exclusion logic.',
    '',
    '4. SUGGESTED ACTIONS',
    '3 specific, data-led suggestions grounded solely in what the n-grams reveal. Each must: (a) name the specific n-gram theme, (b) state the actual performance figures, (c) give a clear, specific recommendation. Rules:',
    '- Do NOT use vague phrases like \'consider reviewing\', \'ensure alignment\', or \'investigate performance\' -- these are not actionable.',
    '- Do NOT tell the analyst to look at their search terms report, their product feed, or any other data source -- they already have access to that data. The suggestion must come from what the n-grams themselves show.',
    '- Do NOT suggest ad groups, audience targeting, or creative assets -- you have no visibility of these.',
    '- Use \'Consider\' or \'Test\' only when followed by a concrete, specific action.',
    '- Say exactly what to do and why the data supports it.',
    '- If the campaign data contains fewer than 3 distinct actionable n-gram themes (after excluding stop words and marginal performers), write 3 suggestions based on what IS in the data -- do not repeat the same theme multiple times to fill the word count.',
    '',
    'Keep the whole analysis under 350 words. Do not use markdown, asterisks, or hash symbols.',
  ].join('\n'),

};
// =============================================================================
// END OF AI CONFIGURATION
// =============================================================================


// -----------------------------------------------------------------------------
// Constants
// Derived from CONFIG. Changing these will break things.
// -----------------------------------------------------------------------------
const STAT_COLS   = ['clicks', 'impressions', 'cost', 'conversions', 'conversionsValue'];
const STAT_LABELS = ['Clicks', 'Impressions', 'Cost', 'Conversions', 'Conv. Value'];

// Each entry is [column label, numerator field, denominator field].
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

// Recorded at startup to track elapsed time throughout the run.
const START_TS = Date.now();

const RAW_STAT_COUNT  = STAT_COLS.length;   // 5
const CALC_STAT_COUNT = CALC_STATS.length;  // 5


// -----------------------------------------------------------------------------
// main()
// Entry point. Orchestrates the full pipeline:
//   1. Open the spreadsheet and fetch campaign IDs.
//   2. Load negative keywords for fast lookup.
//   3. Initialise output sheets.
//   4. Stream search query rows, accumulate n-gram stats, flush to sheets in
//      batches so memory stays flat regardless of data volume.
//   5. Deduplicate, sort, format, and add filters to all sheets.
//   6. Generate the AI summary if an API key is configured.
// -----------------------------------------------------------------------------
function main() {
  validateConfig();

  const ss          = openSpreadsheet();
  const gaqlRange   = buildDateRange();
  const campaignIds = fetchActiveCampaignIds(gaqlRange);

  if (campaignIds.length === 0) {
    Logger.log('⚠️  No active campaigns found with impressions in this date range. Nothing to mine.');
    return;
  }
  Logger.log('');
  Logger.log('🔴 WeDiscover | Search Query N-Gram Analysis');
  Logger.log('📊 Mining ' + campaignIds.length + ' campaign(s). Hang tight.');
  Logger.log('');

  const { negsByAdGroup, negsByCampaign } = CONFIG.checkNegatives
    ? fetchAllNegatives(campaignIds, gaqlRange)
    : { negsByAdGroup: new Map(), negsByCampaign: new Map() };

  Logger.log('✅ Negative keywords mapped.');

  const filterText = buildFilterText() + ' [processing...]';
  const sheets     = initialiseSheets(ss, filterText);
  Logger.log('✅ Sheets ready.');

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

    // Every FLUSH_EVERY rows, write current Maps to the sheet and reset them.
    // This keeps memory flat; the deduplication pass in finaliseSheets()
    // merges any rows that appear across multiple batches.
    if (rowsProcessed % FLUSH_EVERY === 0) {
      const elapsed = (Date.now() - START_TS) / 60000;
      Logger.log('⛏️  ' + rowsProcessed.toLocaleString() + ' rows | ' + elapsed.toFixed(1) + ' min | Writing batch...');

      flushToSheets(sheets, ngramMaps, wordCountMap);
      flushCount++;
      ngramMaps    = buildEmptyNgramMaps();
      wordCountMap = new Map();

      Logger.log('   ✓ Batch ' + flushCount + ' written. Back to mining...');

      if ((Date.now() - START_TS) / 60000 > CONFIG.maxExecutionMinutes) {
        Logger.log('⏱️  ' + CONFIG.maxExecutionMinutes + '-min limit reached. Wrapping up what we have...');
        stoppedEarly = true;
        break;
      }
    }
  }

  Logger.log('⛏️  Writing final batch...');
  flushToSheets(sheets, ngramMaps, wordCountMap);

  Logger.log('✅ Mining complete: ' + rowsProcessed.toLocaleString() + ' rows processed, ' + rowsSkipped.toLocaleString() + ' skipped by negatives.');

  const finalFilterText = buildFilterText() + (stoppedEarly ? ' [PARTIAL -- stopped early]' : '');
  finaliseSheets(sheets, finalFilterText);
  deleteEmptySheets(ss);

  if (AI_CONFIG.apiKey) {
    var aiMode = (AI_CONFIG.mode === 'full') ? 'full' : 'account';
    Logger.log('');
    Logger.log('🤖 AI summary starting (mode: ' + aiMode + ')...');
    try {
      generateAiSummary(ss, aiMode);
      Logger.log('✅ AI summary written.');
    } catch (e) {
      Logger.log('⚠️  AI summary failed (data sheets are fine): ' + e);
    }
  }

  Logger.log('');
  Logger.log('🎉 Done! Completed in ' + ((Date.now() - START_TS) / 60000).toFixed(2) + ' minutes.');
  Logger.log('📎 Your results: ' + CONFIG.spreadsheetUrl);
  Logger.log('');
}


// -----------------------------------------------------------------------------
// validateConfig()
// Basic sanity checks on CONFIG before anything else runs.
// Logs warnings rather than throwing so the script still produces data output
// even when the AI section is misconfigured.
// -----------------------------------------------------------------------------
function validateConfig() {
  if (CONFIG.maxNGramLength > 6) {
    Logger.log('⚠️  maxNGramLength > 6 risks timeouts on large accounts. Consider setting it to 4.');
  }
  if (AI_CONFIG.apiKey && AI_CONFIG.apiKey.indexOf('YOUR_') > -1) {
    Logger.log('⚠️  AI_CONFIG.apiKey looks like a placeholder. Replace it with your real key or leave it blank.');
  }
  if (AI_CONFIG.apiKey && AI_CONFIG.mode !== 'account' && AI_CONFIG.mode !== 'full') {
    Logger.log('⚠️  AI_CONFIG.mode must be "account" or "full". Defaulting to "account".');
  }
  if (CONFIG.maxExecutionMinutes > 25) {
    Logger.log('⚠️  maxExecutionMinutes > 25 leaves no room for the AI summary before Google\'s 30-min limit. Try 15-20.');
  }
}


// -----------------------------------------------------------------------------
// openSpreadsheet()
// Opens the spreadsheet at CONFIG.spreadsheetUrl. Throws a clear error if the
// URL is missing or the sheet cannot be accessed.
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
// Deletes any sheet with no content. Called at the end of main() to remove
// blank sheets left over from a previous run or created automatically by Google.
// Checks sheet count before each deletion to avoid leaving the workbook empty.
// -----------------------------------------------------------------------------
function deleteEmptySheets(ss) {
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getLastRow() <= 1 && ss.getSheets().length > 1) {
      Logger.log('   🗑️  Removed empty sheet: ' + allSheets[i].getName());
      ss.deleteSheet(allSheets[i]);
    }
  }
}


// -----------------------------------------------------------------------------
// buildDateRange()
// Converts DD/MM/YYYY dates from CONFIG into the GAQL date filter string
// that AdsApp.search() expects (YYYY-MM-DD format).
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
    return parts[2] + '-' + parts[1] + '-' + parts[0];
  }
  return "segments.date BETWEEN '" + toGaqlDate(CONFIG.startDate) + "' AND '" + toGaqlDate(CONFIG.endDate) + "'";
}


// -----------------------------------------------------------------------------
// fetchActiveCampaignIds(dateRange)
// Returns an array of campaign IDs that had impressions in the date range
// and match any name filters set in CONFIG. Scoping all subsequent queries
// to these IDs means we only pull data we actually need.
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
// Builds two Maps of negative keywords: one by ad group ID and one by
// campaign ID. Includes negatives from shared negative keyword lists.
// Using Maps gives O(1) lookup per query row rather than scanning arrays.
// Each Map value is an array of [keywordText, matchType] pairs.
// -----------------------------------------------------------------------------
function fetchAllNegatives(campaignIds, _dateRange) {
  const negsByAdGroup  = new Map();
  const negsByCampaign = new Map();

  const adGroupStatusClause = CONFIG.ignorePausedAdGroups
    ? 'AND ad_group.status = \'ENABLED\''
    : 'AND ad_group.status IN (\'ENABLED\', \'PAUSED\')';

  const campaignIdList = campaignIds.join(',');

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
// Returns true if the search query would be blocked by a negative keyword.
// Exact match negatives must match the whole query. Phrase and broad match
// negatives need only appear as whole words anywhere in the query.
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
// Returns an AdsApp iterator over the search term view. Returning the iterator
// directly rather than collecting into an array means the full report is never
// held in memory at once.
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
// parseInt and parseFloat are used because the Ads Scripts runtime sometimes
// returns metric fields as strings. Cost arrives in micros (millionths of the
// currency unit) and is divided by 1,000,000 to get the actual amount.
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
// Creates a fresh set of accumulator Maps, one per n-gram length.
// Each length gets three Maps: total (account level), campaign, and ad group.
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
// Extracts every n-gram of every configured length from the search query and
// adds this row's stats to the running total for each phrase.
//
// Keys use a null byte separator (e.g. "Campaign Name\x00=\"running shoes\"")
// so campaign and ad group data can be stored in flat Maps without nesting.
// A Set prevents the same phrase being counted twice within one query.
//
// The =" prefix on phrases forces Google Sheets to treat the cell as plain
// text, stopping numbers like "1000" being interpreted as numeric values.
// -----------------------------------------------------------------------------
function accumulateNgrams(row, ngramMaps, wordCountMap) {
  const words = row.query.split(' ');
  const stats = statsFromRow(row);

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

      const displayPhrase = '="' + phrase + '"';

      addStats(level.total,    displayPhrase, stats);
      addStats(level.campaign, row.campaignName + '\x00' + displayPhrase, stats);
      addStats(level.adgroup,  row.campaignName + '\x00' + row.adGroupName + '\x00' + displayPhrase, stats);
    }
  }
}


// -----------------------------------------------------------------------------
// statsFromRow(row)
// Pulls metric fields out of a parsed query row. queryCount starts at 1
// because this represents one search query row.
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
// Adds stats from one query row to the running total in the Map.
// Object.assign is used on first insertion so the original object is not mutated.
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
// buildDateRangeText()
// Returns the analysis date range in British (DD/MM/YYYY) format for display
// in the teal status bar (row 2) of every output sheet. Dates are normalised
// to zero-padded DD/MM/YYYY so the range reads consistently regardless of how
// they were entered in CONFIG.
// -----------------------------------------------------------------------------
function buildDateRangeText() {
  function fmt(ddmmyyyy) {
    var parts = String(ddmmyyyy).split('/');
    if (parts.length !== 3) return ddmmyyyy;
    var dd = parts[0].length === 1 ? '0' + parts[0] : parts[0];
    var mm = parts[1].length === 1 ? '0' + parts[1] : parts[1];
    return dd + '/' + mm + '/' + parts[2];
  }
  return 'Date range: ' + fmt(CONFIG.startDate) + ' to ' + fmt(CONFIG.endDate);
}


// -----------------------------------------------------------------------------
// buildFilterText()
// Human-readable summary of the analysis date range plus which campaigns and
// ad groups are included. Appears in the teal status bar (row 2) at the top of
// each sheet. Every row-2 status string flows through this function -- the data
// sheets via initialiseSheets() and finaliseSheets(), and the AI Summary sheet
// via initialiseSummarySheet() -- so adding the date range here puts it on
// every output sheet.
// -----------------------------------------------------------------------------
function buildFilterText() {
  let text = CONFIG.ignorePausedAdGroups ? 'Active ad groups' : 'All ad groups';
  text += CONFIG.ignorePausedCampaigns ? ' in active campaigns' : ' in all campaigns';
  if (CONFIG.campaignNameContains)       text += " containing '" + CONFIG.campaignNameContains + "'";
  if (CONFIG.campaignNameDoesNotContain) text += " not containing '" + CONFIG.campaignNameDoesNotContain + "'";
  return buildDateRangeText() + '  |  ' + text;
}


// =============================================================================
// SPREADSHEET OUTPUT
// =============================================================================

// WeDiscover brand colours matched from we-discover.com.
// Montserrat is the closest Google Sheets equivalent to WeDiscover's brand fonts.
var WD_CRIMSON = '#C0392B';
var WD_TEAL    = '#3ECFB2';
var WD_NAVY    = '#1A1F36';
var WD_CREAM   = '#FAF8F5';
var WD_WHITE   = '#FFFFFF';
var WD_FONT    = 'Montserrat';


// -----------------------------------------------------------------------------
// buildSheetDefs()
// Returns an array describing every output sheet. Each entry defines the tab
// name, column headers, how to parse a compound Map key back into label columns,
// and how many label columns precede the numeric stats. Adding a new sheet
// only requires one change here -- initialiseSheets, flushToSheets, and
// finaliseSheets all read from this definition.
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
//   Row 2: teal status bar (value set separately in initialiseSheets).
//   Row 3: navy column header row.
// Also sets the tab colour, freezes headers, and sets column widths.
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

  // 10 stat columns on the right; everything to the left is a label column.
  var labelCols = numCols - 10;
  for (var c = 1; c <= numCols; c++) {
    sheet.setColumnWidth(c, c <= labelCols ? 160 : 120);
  }
  sheet.setColumnWidth(1, 240);
}


// -----------------------------------------------------------------------------
// applyFilter(sheet, def)
// Removes any existing filter and applies a fresh one covering the header and
// all data rows. Removing first ensures a clean state on repeat runs.
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
// Creates all output sheets (or clears existing ones), applies brand styling,
// writes column headers, and returns a registry Map.
// The registry stores { sheet, def, nextRow } for each sheet name.
// nextRow tracks where the next batch should be written, avoiding the need to
// call getLastRow() on every flush.
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

    // Row 2 value must be set after styleSheet() calls merge() on it.
    sheet.getRange('A2').setValue(filterText);

    var hRange = sheet.getRange(3, 1, 1, def.header.length);
    hRange.setValues([def.header]);
    hRange.setBackground(WD_NAVY);
    hRange.setFontColor(WD_WHITE);
    hRange.setFontWeight('bold');

    registry.set(def.name, { sheet: sheet, def: def, nextRow: 4 });
  }

  Logger.log('   📋 ' + defs.length + ' sheets created.');
  return registry;
}


// -----------------------------------------------------------------------------
// flushToSheets(sheets, ngramMaps, wordCountMap)
// Writes the current batch of n-gram data to the spreadsheet.
// Only rows that meet all threshold settings are written.
// SpreadsheetApp.flush() at the end commits all pending writes in one round trip.
//
// The same phrase can appear in multiple flush batches when data is large.
// This is intentional -- keeping raw sums per batch is cheaper than maintaining
// one growing Map. deduplicateSheet() in finaliseSheets() merges them.
// -----------------------------------------------------------------------------
function flushToSheets(sheets, ngramMaps, wordCountMap) {
  const t = CONFIG.thresholds;

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
// labelParts contains the text columns (phrase, campaign name, etc.).
// Calculated stats (CTR, CPC, etc.) are appended at the end.
// A hyphen is written instead of dividing by zero.
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
//
// The flush-and-reset architecture means the same phrase can appear as multiple
// rows across different batches. These functions fix that by reading each sheet
// back after all batches are written, merging rows with matching keys, and
// rewriting a clean deduplicated set.
//
// Why not keep a persistent Map instead?
// On large accounts a persistent Map holding every unique phrase seen so far
// grows to hundreds of thousands of entries and causes memory errors. The
// flush-and-reset approach keeps memory flat. The deduplication pass reads one
// sheet at a time, which Sheets handles easily, and the merged Map is much
// smaller than the full raw-data Map would have been.
// =============================================================================


// -----------------------------------------------------------------------------
// deduplicateSheet(entry)
// Reads all data rows from a sheet, merges rows that share the same label key,
// recalculates derived stats from the merged raw totals, and rewrites the sheet.
//
// Column layout (0-indexed):
//   [0 .. labelCount-1]            label columns (phrase, campaign, etc.)
//   [labelCount]                   Query Count
//   [labelCount+1 .. labelCount+5] raw stats (clicks, impressions, cost,
//                                  conversions, conversionsValue)
//   [labelCount+6 .. labelCount+10] calculated stats (recalculated on write)
//
// Returns the number of rows written after deduplication.
// -----------------------------------------------------------------------------
function deduplicateSheet(entry) {
  const sheet      = entry.sheet;
  const def        = entry.def;
  const lastRow    = entry.nextRow - 1;
  const firstData  = 4;

  if (lastRow < firstData) return 0;

  const numCols    = def.header.length;
  const dataRows   = sheet.getRange(firstData, 1, lastRow - firstData + 1, numCols).getValues();

  const lc         = def.labelCount;
  const qcIdx      = lc;
  const statStart  = lc + 1;

  const merged   = new Map();
  const keyOrder = [];

  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];

    var keyParts = [];
    for (var k = 0; k < lc; k++) {
      keyParts.push(row[k]);
    }
    var key = keyParts.join('\x00');

    if (!merged.has(key)) {
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

    // Coerce to number -- a previous run may have written strings.
    acc.queryCount       += parseFloat(row[qcIdx])          || 0;
    acc.clicks           += parseFloat(row[statStart])      || 0;
    acc.impressions      += parseFloat(row[statStart + 1])  || 0;
    acc.cost             += parseFloat(row[statStart + 2])  || 0;
    acc.conversions      += parseFloat(row[statStart + 3])  || 0;
    acc.conversionsValue += parseFloat(row[statStart + 4])  || 0;
  }

  const cleanRows = [];
  for (var ki = 0; ki < keyOrder.length; ki++) {
    var acc2 = merged.get(keyOrder[ki]);
    cleanRows.push(buildPrintline(acc2.labels, acc2));
  }

  sheet.getRange(firstData, 1, lastRow - firstData + 1, numCols).clearContent();
  if (cleanRows.length > 0) {
    sheet.getRange(firstData, 1, cleanRows.length, numCols).setValues(cleanRows);
  }

  SpreadsheetApp.flush();
  return cleanRows.length;
}


// -----------------------------------------------------------------------------
// HEADER_NOTES
// Hover tooltips for each column header, shown when a user hovers over row 3.
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
// Adds hover notes from HEADER_NOTES to column header cells in row 3.
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
// Colours data rows alternately white and warm cream for readability.
// -----------------------------------------------------------------------------
function applyAlternatingRows(sheet, firstDataRow, dataRowCount, numCols) {
  for (var r = 0; r < dataRowCount; r++) {
    var bg = (r % 2 === 0) ? WD_WHITE : WD_CREAM;
    sheet.getRange(firstDataRow + r, 1, 1, numCols).setBackground(bg);
  }
}


// -----------------------------------------------------------------------------
// setDataFont(sheet, firstDataRow, dataRowCount, numCols, labelCount)
// Sets data rows to Arial. Montserrat is great on headers but slow to render
// at thousands of rows. Arial is fast, clean, and universally available.
// Numeric columns are right-aligned; label columns are left-aligned.
// The hyphen written for zero-denominator stats (e.g. Cost/Conv. with no
// conversions) is a string and would left-align without the explicit override.
// -----------------------------------------------------------------------------
function setDataFont(sheet, firstDataRow, dataRowCount, numCols, labelCount) {
  sheet.getRange(firstDataRow, 1, dataRowCount, numCols)
    .setFontFamily('Arial')
    .setFontSize(9)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('right');

  sheet.getRange(firstDataRow, 1, dataRowCount, labelCount)
    .setHorizontalAlignment('left');
}


// -----------------------------------------------------------------------------
// trimSheet(sheet, usedCols, usedRows)
// Deletes columns and rows beyond the data range. Google Sheets creates 26
// columns and 1,000 rows by default; trimming keeps the file lean.
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
// Runs once after all data has been flushed. For each sheet:
//   1. Deduplicates rows written across multiple flush batches.
//   2. Updates the teal status bar with the final filter text.
//   3. Applies number formatting to all data rows in one API call.
//   4. Sorts: label columns ascending, then clicks descending.
//   5. Applies alternating row colours.
//   6. Sets data font to Arial.
//   7. Adds hover notes to column headers.
//   8. Trims unused rows and columns.
//   9. Adds a column filter to the header row.
//
// Deduplication runs first because sort and format depend on the final row count.
// All formatting is done here rather than on each flush to pay the Sheets API
// cost once per sheet regardless of how many flushes ran.
// -----------------------------------------------------------------------------
function finaliseSheets(sheets, finalFilterText) {
  for (const [, entry] of sheets) {
    const sheet = entry.sheet;
    const def   = entry.def;

    const dedupedRowCount = deduplicateSheet(entry);
    entry.nextRow = 4 + dedupedRowCount;

    Logger.log('   🔄 ' + sheet.getName() + ': ' + (entry.nextRow - 4).toLocaleString() + ' unique rows');

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

    // Label columns use '@' (plain text). Query Count uses '#,##0'.
    // Stat and calculated columns use the formats in COL_FORMATS.
    const labelCount = numCols - COL_FORMATS.length - 1;
    const fullFmtRow = Array(labelCount).fill('@')
      .concat(['#,##0'])
      .concat(COL_FORMATS);
    sheet.getRange(4, 1, dataRowCount, numCols)
      .setNumberFormats(Array(dataRowCount).fill(fullFmtRow));

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

    Logger.log('   ✓ Sorted & formatted: ' + sheet.getName());
  }
}


// =============================================================================
// AI SUMMARY
//
// HOW THIS WORKS
// --------------
// After all data sheets are written, this section generates an AI summary
// using Google's Gemini Flash model via the Google AI Studio API.
//
// 1. The script reads phrase data back from the finished Google Sheets output
//    (not from the original Ads data -- this is fast).
//
// 2. The client's brand name is detected automatically by querying the most
//    common domain across all ad final URLs. This is more reliable than using
//    the account name, which may reflect the agency name or other naming
//    conventions rather than the actual client brand.
//
// 3. Data is formatted as compact CSV (phrase, clicks, cost, conversions,
//    conv_value, roas) and sent to Gemini along with prompt instructions.
//
// 4. In 'account' mode: one API call with the top N phrases per account-level
//    sheet. In 'full' mode: one call per campaign using campaign and ad group
//    level data.
//
// 5. For large campaigns the script automatically splits data into chunks that
//    fit within the token budget, sends them separately, and combines responses.
//
// 6. Results are written to a branded "AI Summary" sheet at the front of the
//    workbook. Any campaigns not reached before the time limit are listed in a
//    NOTICE section at the bottom of the sheet.
//
// 7. Any failure in the AI section is caught and logged without affecting the
//    data sheets.
// =============================================================================


// -----------------------------------------------------------------------------
// fetchClientBrand()
// Identifies the client's proper brand name in two steps.
//
// Step 1 -- domain detection:
//   Queries a sample of enabled ads and extracts the registered domain from
//   each final URL. The most frequently occurring domain wins.
//   e.g. "notonthehighstreet.com" -> raw domain "notonthehighstreet"
//        "shop.example.co.uk"     -> raw domain "example"
//
//   Falls back to the Google Ads account descriptive name if no URLs are found,
//   and to the string "the client" if that also fails.
//
// Step 2 -- brand name resolution via Gemini:
//   The raw domain string is passed to Gemini with a tight prompt asking it to
//   return the correctly formatted brand name and nothing else. This handles
//   cases that simple string manipulation cannot:
//     "notonthehighstreet" -> "Not On The High Street"
//     "marksandspencer"    -> "Marks & Spencer"
//     "johnlewis"          -> "John Lewis"
//   If the Gemini call fails for any reason, the raw domain (title-cased as a
//   fallback) is used so the rest of the summary is not blocked.
// -----------------------------------------------------------------------------
function fetchClientBrand() {
  var rawBrand = 'the client';

  // -- Step 1: detect the most common domain from ad final URLs ----------------
  try {
    var adQuery = [
      'SELECT ad_group_ad.ad.final_urls',
      'FROM   ad_group_ad',
      'WHERE  ad_group_ad.status = \'ENABLED\'',
      '       AND campaign.status = \'ENABLED\'',
      'LIMIT  500',
    ].join(' ');

    var domainCount = {};
    for (var row of AdsApp.search(adQuery)) {
      var urls = row.adGroupAd.ad.finalUrls;
      if (!urls || urls.length === 0) continue;

      var url      = urls[0];
      var hostname = url.replace(/^https?:\/\//i, '').split('/')[0].split('?')[0].toLowerCase();
      hostname     = hostname.replace(/^www\./, '');

      // Extract the registered domain (the part before the public suffix).
      // For two-part TLDs like .co.uk or .com.au, take the third-from-last segment.
      var parts = hostname.split('.');
      var registeredName;
      if (parts.length >= 3 && parts[parts.length - 2].length <= 3) {
        registeredName = parts[parts.length - 3];
      } else {
        registeredName = parts[parts.length - 2];
      }

      if (registeredName) {
        domainCount[registeredName] = (domainCount[registeredName] || 0) + 1;
      }
    }

    var topDomain = null;
    var topCount  = 0;
    for (var domain in domainCount) {
      if (domainCount[domain] > topCount) {
        topCount  = domainCount[domain];
        topDomain = domain;
      }
    }

    if (topDomain) {
      rawBrand = topDomain;
      Logger.log('🔍 Domain detected: "' + rawBrand + '" (' + topCount + ' ads)');
    } else {
      throw new Error('no URLs found');
    }

  } catch (e) {
    Logger.log('⚠️  Could not detect domain from URLs. Trying account name...');

    // Fallback: use the Google Ads account descriptive name.
    try {
      var customerQuery = 'SELECT customer.descriptive_name FROM customer LIMIT 1';
      for (var custRow of AdsApp.search(customerQuery)) {
        var name = custRow.customer.descriptiveName;
        if (name) {
          rawBrand = name;
          Logger.log('🔍 Using account name as brand: "' + rawBrand + '"');
          break;
        }
      }
    } catch (e2) {
      Logger.log('⚠️  Could not read account name either. Using generic fallback.');
    }
  }

  // If detection produced nothing useful, return early without burning an API call.
  if (rawBrand === 'the client') return rawBrand;

  // -- Step 2: ask Gemini to resolve the proper brand name ---------------------
  // The domain string alone is often not the correct brand name. "notonthehighstreet"
  // should become "Not On The High Street"; "marksandspencer" should become
  // "Marks & Spencer". Simple title-casing cannot handle these cases reliably.
  // One small Gemini call resolves this correctly for virtually any brand.
  try {
    var brandPrompt = [
      'A Google Ads account has the domain name "' + rawBrand + '".',
      'What is the correct, properly formatted brand name for this company?',
      'Reply with ONLY the brand name -- no explanation, no punctuation, no extra words.',
      'Examples:',
      '  notonthehighstreet -> Not On The High Street',
      '  marksandspencer    -> Marks & Spencer',
      '  johnlewis          -> John Lewis',
      '  ikea               -> IKEA',
      '  asos               -> ASOS',
    ].join('\n');

    var resolvedBrand = callGemini(brandPrompt).trim();

    // Sanity check: reject the response if it is suspiciously long (more than
    // 60 characters suggests Gemini returned an explanation rather than a name)
    // or empty, and fall back to a simple title-cased version of the raw domain.
    if (!resolvedBrand || resolvedBrand.length > 60) {
      throw new Error('unexpected response: "' + resolvedBrand + '"');
    }

    Logger.log('✅ Brand: "' + resolvedBrand + '"');
    return resolvedBrand;

  } catch (e3) {
    // Non-fatal: fall back to title-casing the raw domain string.
    var fallback = rawBrand.split('-').map(function(w) {
      return w.charAt(0).toUpperCase() + w.slice(1);
    }).join(' ');
    Logger.log('⚠️  Brand resolution failed. Using: "' + fallback + '"');
    return fallback;
  }
}


// -----------------------------------------------------------------------------
// countTokens(rows)
// Estimates the token count for a 2D array of row data.
// Uses a 4-characters-per-token approximation, standard for mixed
// English text and numbers.
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
// Converts a campaign-level sheet row to a compact CSV line.
// Campaign sheet columns (0-indexed): 0 campaign, 1 phrase, 2 query count,
// 3 clicks, 4 impressions, 5 cost, 6 conversions, 7 conv value, 8-12 calc stats.
// Sends: phrase, clicks, cost, conversions, conv_value, roas.
// ROAS is recalculated from cost and conv_value for accuracy.
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
// Converts an ad group-level sheet row to a compact CSV line.
// Ad group sheet columns (0-indexed): 0 campaign, 1 ad group, 2 phrase,
// 3 query count, 4 clicks, 5 impressions, 6 cost, 7 conversions, 8 conv value,
// 9-13 calculated stats.
// Sends: ad_group, phrase, clicks, cost, conversions, conv_value, roas.
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
// Reads data rows from a list of sheets and groups them by campaign name.
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
// Calculates average ROAS for a campaign from its campaign-level rows.
// Campaign cost is at index 5, conv value at index 7.
// Returns a formatted string like "4.32", or "0" if cost is zero.
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
// topNRowsByClicks(rows, clicksColIndex, n)
// Returns the top n rows sorted by clicks descending.
// Used to limit rows sent to the AI per campaign and per ad group, which is
// the primary control on chunk count for large accounts.
// If n is 0 or negative, all rows are returned unchanged.
// -----------------------------------------------------------------------------
function topNRowsByClicks(rows, clicksColIndex, n) {
  if (!n || n <= 0) return rows;
  var sorted = rows.slice().sort(function(a, b) {
    return (parseFloat(b[clicksColIndex]) || 0) - (parseFloat(a[clicksColIndex]) || 0);
  });
  return sorted.slice(0, n);
}


// -----------------------------------------------------------------------------
// buildCampaignChunks(campaignName, campRows, agRowsByAdGroup)
// Determines how to split a campaign's data into API calls.
//
// topN filtering is applied first (clicks descending) to both campaign-level
// rows and each individual ad group's rows. This is the primary control on
// chunk count -- on most accounts it alone is sufficient to keep calls to 1-2.
//
// After filtering, ad groups are packed into chunks up to the token budget.
// If a single ad group still exceeds the budget it is split at phrase level.
//
// If the natural chunk count still exceeds maxChunksPerCampaign, all filtered
// ad group rows are re-partitioned evenly across that many chunks as a
// last-resort safety net.
//
// Each returned chunk object contains:
//   agRows    -- ad group rows for this chunk
//   campRows  -- filtered campaign-level rows (same for every chunk)
//   chunkIdx  -- 1-based index of this chunk
//   total     -- total number of chunks for this campaign
// -----------------------------------------------------------------------------
function buildCampaignChunks(campaignName, campRows, agRowsByAdGroup) {

  var filteredCampRows = topNRowsByClicks(campRows, 3, AI_CONFIG.topN);

  var campTok  = countTokens(filteredCampRows);
  var agBudget = AI_CONFIG.tokenBudget - campTok - 200;

  var filteredAgByAg    = new Map();
  var allFilteredAgRows = [];
  var agNames           = Array.from(agRowsByAdGroup.keys());

  for (var i = 0; i < agNames.length; i++) {
    var agName   = agNames[i];
    var agRows   = agRowsByAdGroup.get(agName);
    var filtered = topNRowsByClicks(agRows, 4, AI_CONFIG.topN);
    filteredAgByAg.set(agName, filtered);
    for (var r = 0; r < filtered.length; r++) {
      allFilteredAgRows.push(filtered[r]);
    }
  }

  // Normal chunking: pack ad groups into chunks up to the token budget.
  var chunks   = [];
  var curChunk = [];
  var curTok   = 0;

  for (var i2 = 0; i2 < agNames.length; i2++) {
    var agName2 = agNames[i2];
    var agRows2 = filteredAgByAg.get(agName2);
    var agTok   = countTokens(agRows2);

    if (agTok > agBudget) {
      // Ad group still too large after topN filtering -- split at phrase level.
      if (curChunk.length > 0) {
        chunks.push(curChunk);
        curChunk = [];
        curTok   = 0;
      }
      var phraseChunk = [];
      var phraseTok   = 0;
      for (var r2 = 0; r2 < agRows2.length; r2++) {
        var rowTok = countTokens([agRows2[r2]]);
        if (phraseChunk.length > 0 && phraseTok + rowTok > agBudget) {
          chunks.push(phraseChunk);
          phraseChunk = [];
          phraseTok   = 0;
        }
        phraseChunk.push(agRows2[r2]);
        phraseTok += rowTok;
      }
      if (phraseChunk.length > 0) chunks.push(phraseChunk);

    } else if (curTok + agTok > agBudget) {
      if (curChunk.length > 0) chunks.push(curChunk);
      curChunk = agRows2.slice();
      curTok   = agTok;

    } else {
      curChunk = curChunk.concat(agRows2);
      curTok  += agTok;
    }
  }
  if (curChunk.length > 0) chunks.push(curChunk);

  // Enforce maxChunksPerCampaign: if still over the cap, re-partition evenly.
  var cap = AI_CONFIG.maxChunksPerCampaign;
  if (cap > 0 && chunks.length > cap) {
    Logger.log('   ✂️  ' + campaignName + ': ' + chunks.length + ' chunks → capped at ' + cap + '.');
    chunks = [];
    var rowsPerChunk = Math.ceil(allFilteredAgRows.length / cap);
    for (var c = 0; c < cap; c++) {
      var slice = allFilteredAgRows.slice(c * rowsPerChunk, (c + 1) * rowsPerChunk);
      if (slice.length > 0) chunks.push(slice);
    }
  }

  var result = [];
  for (var ci = 0; ci < chunks.length; ci++) {
    result.push({ agRows: chunks[ci], campRows: filteredCampRows, chunkIdx: ci + 1, total: chunks.length });
  }
  return result;
}


// -----------------------------------------------------------------------------
// buildAccountPromptText(ss, clientBrand)
// Reads the four account-level sheets and builds the prompt for the account
// summary call. Sends the top AI_CONFIG.topN rows by clicks from each sheet.
// -----------------------------------------------------------------------------
function buildAccountPromptText(ss, clientBrand) {
  var sheetNames = [
    'Account Word Analysis',
    'Account 2-Gram Analysis',
    'Account 3-Gram Analysis',
    'Account 4-Gram Analysis',
  ];

  var dataSections = [];

  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    var numCols = sheet.getLastColumn();
    var data    = sheet.getRange(4, 1, lastRow - 3, numCols).getValues();

    var sorted = data.slice().sort(function(a, b) {
      return (parseFloat(b[2]) || 0) - (parseFloat(a[2]) || 0);
    });

    var limit = Math.min(AI_CONFIG.topN, sorted.length);
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

    dataSections.push('--- ' + sheetNames[i] + ' (top ' + limit + ' by clicks) ---');
    dataSections.push(lines.join('\n'));
  }

  var prompt = AI_CONFIG.accountPrompt.replace(/\{CLIENT_BRAND\}/g, clientBrand);
  return prompt + '\n\n' + dataSections.join('\n');
}


// -----------------------------------------------------------------------------
// buildCampaignPromptText(campaignName, campRows, agChunkRows, chunkIdx, total, avgRoas, clientBrand)
// Builds the prompt for a single campaign API call.
// campRows is the already-filtered campaign-level data (same for all chunks).
// agChunkRows is the ad group data for this specific chunk.
// -----------------------------------------------------------------------------
function buildCampaignPromptText(campaignName, campRows, agChunkRows, chunkIdx, total, avgRoas, clientBrand) {
  var chunkNote = total > 1
    ? '(part ' + chunkIdx + ' of ' + total + ' -- this is a large campaign split across multiple analyses)'
    : '';

  var instruction = AI_CONFIG.campaignPrompt
    .replace(/\{CLIENT_BRAND\}/g, clientBrand)
    .replace('{CAMPAIGN}',        campaignName)
    .replace('{CHUNK_NOTE}',      chunkNote)
    .replace(/\{AVG_ROAS\}/g,     avgRoas);

  var campLines = ['--- Campaign-level phrase data ---',
                   'phrase,clicks,cost,conversions,conv_value,roas'];
  for (var i = 0; i < campRows.length; i++) {
    campLines.push(formatCampaignRow(campRows[i]));
  }

  var agLines = ['--- Ad group phrase data ---',
                 'ad_group,phrase,clicks,cost,conversions,conv_value,roas'];
  for (var j = 0; j < agChunkRows.length; j++) {
    agLines.push(formatAdGroupRow(agChunkRows[j]));
  }

  return instruction + '\n\n' + campLines.join('\n') + '\n\n' + agLines.join('\n');
}


// -----------------------------------------------------------------------------
// callGemini(promptText)
// Sends promptText to Gemini Flash and returns the response text.
//
// Uses UrlFetchApp (the standard Apps Script HTTP client) rather than any SDK,
// since the Ads Scripts runtime does not support npm packages.
//
// Model: controlled by AI_CONFIG.model. Defaults to gemini-2.5-flash-lite,
// which offers the best free-tier RPM (30/min) of any stable Flash model.
// To switch models, update AI_CONFIG.model -- nothing else needs to change.
//
// temperature: 0.3 for consistent, factual output.
// maxOutputTokens: 8,192 -- enough for a detailed per-campaign analysis.
// muteHttpExceptions: true so 429 and other error responses can be handled
// cleanly rather than throwing an unformatted HTTP error.
//
// Rate limit handling:
// On a 429 (quota exceeded) the function waits 60 seconds and retries once.
// The free tier limit resets on a per-minute window, so a 60-second pause is
// always sufficient regardless of the retry delay Gemini suggests in the error
// message. If the retry also returns 429, the error is thrown normally.
// -----------------------------------------------------------------------------
function callGemini(promptText) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + AI_CONFIG.model + ':generateContent?key=' + AI_CONFIG.apiKey;
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

  // On 429 (rate limit), wait 60 seconds and try once more.
  if (code === 429) {
    Logger.log('⏳ Rate limit hit. Waiting 60 seconds before retrying...');
    Utilities.sleep(60000);
    response = UrlFetchApp.fetch(url, options);
    code     = response.getResponseCode();
    body     = JSON.parse(response.getContentText());
  }

  if (code !== 200) {
    var errMsg = (body.error && body.error.message) ? body.error.message : 'HTTP ' + code;
    throw new Error('Gemini API error: ' + errMsg);
  }

  if (!body.candidates || !body.candidates[0] ||
      !body.candidates[0].content || !body.candidates[0].content.parts ||
      !body.candidates[0].content.parts[0]) {
    throw new Error('Gemini API returned an unexpected response structure: ' + JSON.stringify(body));
  }

  // Pace requests to stay within the free-tier rate limit window.
  if (AI_CONFIG.callDelayMs > 0) Utilities.sleep(AI_CONFIG.callDelayMs);

  return body.candidates[0].content.parts[0].text;
}


// -----------------------------------------------------------------------------
// initialiseSummarySheet(ss, filterText)
// Creates or clears the AI Summary sheet, applies brand styling, moves it to
// the front of the workbook, and returns it.
// -----------------------------------------------------------------------------
function initialiseSummarySheet(ss, filterText) {
  var sheetName = AI_CONFIG.sheetName;
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(0);

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
    ? 'Full mode  |  ' + AI_CONFIG.topN + ' phrases per campaign analysed  |  ' + AI_CONFIG.maxChunksPerCampaign + ' API calls per campaign max'
    : 'Account mode  |  Top ' + AI_CONFIG.topN + ' phrases per account-level sheet';

  var r2 = sheet.getRange(2, 1, 1, numCols);
  r2.setValue(filterText + '  |  ' + modeLabel);
  r2.setBackground(WD_TEAL);
  r2.setFontColor(WD_NAVY);
  r2.setFontFamily(WD_FONT);
  r2.setFontSize(9);
  r2.setFontWeight('bold');
  r2.setVerticalAlignment('middle');
  r2.setHorizontalAlignment('left');
  r2.setWrap(true);
  sheet.setRowHeight(2, 40);

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
// Post-processes raw Gemini output before writing to the sheet.
// 1. Forces ROAS to all-caps (the model frequently ignores this instruction).
// 2. Replaces leading hyphen bullets with the Unicode bullet character.
// -----------------------------------------------------------------------------
function normaliseAiResponse(text) {
  text = text.replace(/\broas\b/gi, function(match) {
    return match === 'ROAS' ? match : 'ROAS';
  });

  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  text = text.replace(/^- /gm, '  \u2022 ');

  return text;
}


// -----------------------------------------------------------------------------
// appendSectionBlock(sheet, nextRow, headerText, bodyText)
// Writes one section of the AI summary: a teal header row followed by a single
// body cell containing the full section text with embedded line breaks.
//
// A single wrapped cell is used rather than one row per line because it is
// simpler to navigate, avoids row height estimation fighting with Sheets'
// text clipping, and scales cleanly to any response length.
//
// Row height is estimated generously from character count to avoid clipping.
// Returns the updated nextRow value.
// -----------------------------------------------------------------------------
function appendSectionBlock(sheet, nextRow, headerText, bodyText) {
  var numCols = 1;

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

  var CHARS_PER_LINE = 95;
  var LINE_HEIGHT_PX = 18;
  var PADDING_PX     = 20;

  var bodyLines    = cleanBody.split('\n');
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
// Trims unused rows and columns from the summary sheet after all content
// has been written.
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
// Orchestrates the full AI summary pipeline.
//
// Time management:
//   The AI section targets completion within 25 minutes total script runtime.
//   maxAiMinutes = 25 - CONFIG.maxExecutionMinutes gives the budget remaining
//   after mining. Before each campaign the elapsed AI time is checked; if
//   less than 30 seconds remain the loop stops, skipped campaigns are listed
//   in the sheet, and the script exits cleanly rather than being hard-killed.
//
// Account mode:
//   One call with the top N phrases from each account-level sheet.
//
// Full mode:
//   One call per campaign with campaign and ad group level data.
//   Large campaigns are split into chunks automatically.
//   All chunk responses are combined under one campaign header in the sheet.
// -----------------------------------------------------------------------------
function generateAiSummary(ss, mode) {
  var filterText = buildFilterText();
  var sheet      = initialiseSummarySheet(ss, filterText);
  var nextRow    = 4;

  Logger.log('🧠 Model: ' + AI_CONFIG.model);

  // Detect the client's brand from ad final URLs before making any AI calls.
  var clientBrand = fetchClientBrand();

  var AI_START_TS    = Date.now();
  var AI_DEADLINE_MS = AI_CONFIG.maxAiMinutes * 60 * 1000;

  // Minimum headroom per campaign before starting it. 30 seconds is
  // conservative -- single-chunk campaigns typically take 5-15 seconds.
  var MIN_MS_PER_CAMPAIGN = 30 * 1000;

  // ---- Account-level summary -------------------------------------------------
  Logger.log('📝 Writing account summary...');
  var accountPromptText = buildAccountPromptText(ss, clientBrand);
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

  var campRowsByCamp = readSheetRowsByCampaign(ss, campSheetNames, 0);
  var agRowsByCampAg = readSheetRowsByCampaignAndAdGroup(ss, agSheetNames);

  var campNames    = Array.from(campRowsByCamp.keys()).sort();
  var totalCamps   = campNames.length;
  var skippedCamps = [];

  Logger.log('📋 ' + totalCamps + ' campaigns to analyse...');

  for (var ci = 0; ci < campNames.length; ci++) {
    var campName = campNames[ci];

    // Time check: stop before starting a campaign we cannot finish.
    var aiElapsed = Date.now() - AI_START_TS;
    if (aiElapsed + MIN_MS_PER_CAMPAIGN > AI_DEADLINE_MS) {
      Logger.log('⏱️  Time limit reached after ' + (aiElapsed / 60000).toFixed(1) + ' min. Stopping before: ' + campName);
      for (var sk = ci; sk < campNames.length; sk++) {
        skippedCamps.push(campNames[sk]);
      }
      break;
    }

    var campRows = campRowsByCamp.get(campName);
    var agByAg   = agRowsByCampAg.get(campName) || new Map();
    var avgRoas  = calcCampaignAvgRoas(campRows);

    Logger.log('   📊 Campaign ' + (ci + 1) + '/' + totalCamps + ': ' + campName);

    var chunks    = buildCampaignChunks(campName, campRows, agByAg);
    var campParts = [];

    for (var ki = 0; ki < chunks.length; ki++) {
      var chunk      = chunks[ki];
      var promptText = buildCampaignPromptText(
        campName, chunk.campRows, chunk.agRows, chunk.chunkIdx, chunk.total, avgRoas, clientBrand
      );
      var chunkResult = callGemini(promptText);
      campParts.push(chunkResult);

      if (chunks.length > 1) {
        Logger.log('      ✓ Chunk ' + chunk.chunkIdx + '/' + chunk.total);
      }
    }

    var headerText   = campName + '  |  Avg ROAS: ' + avgRoas;
    var responseText = campParts.join('\n\n--- (continued) ---\n\n');
    nextRow = appendSectionBlock(sheet, nextRow, headerText, responseText);
  }

  // If any campaigns were skipped, append a notice so the output is
  // self-documenting and the user knows exactly what is missing and why.
  if (skippedCamps.length > 0) {
    var skipNote = [
      'The following ' + skippedCamps.length + ' campaign(s) were not analysed because the AI time limit (' + AI_CONFIG.maxAiMinutes + ' min) was reached.',
      '',
      'To include them, either:',
      '  1. Reduce CONFIG.maxExecutionMinutes to give the AI more time (e.g. 12 instead of 15).',
      '  2. Use campaignNameContains in CONFIG to run the script for a subset of campaigns.',
      '',
      'Skipped campaigns:',
    ].concat(skippedCamps.map(function(n) { return '  \u2022 ' + n; })).join('\n');

    nextRow = appendSectionBlock(sheet, nextRow, 'NOTICE: CAMPAIGNS NOT ANALYSED (time limit reached)', skipNote);
    Logger.log('⚠️  ' + skippedCamps.length + ' campaign(s) skipped. See NOTICE in the AI Summary sheet.');
  }

  finaliseSummarySheet(sheet, nextRow - 1);
}
