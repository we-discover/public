/*
    Name:           WeDiscover - CPC Distribution Tool, Google Ads Script
    Description:    This script analyses the CPCs of search terms across the account in order
                    to work out how Max CPC bid caps could impact performance.

                    It gives you a spreadsheet output where you can select the campaign of
                    interest from a dropdown menu to see:
                    1. the distribution of CPCs
                    2. the percentage of clicks lost given a particular Max CPC cap
                    3. the percentage of conversions lost given a particular Max CPC cap
                    4. the percentage of conversion value lost given a particular Max CPC cap

                    There's also a table with the percentage of clicks, conversions and conversion
                    value lost given particular Max CPC caps.
                 
    License:        https://github.com/we-discover/public/blob/master/LICENSE
    Version:        1.0.0
    Released:       2024-04-09
    Author:         Nathan Ifill (@nathanifill)
    Contact:        scripts@we-discover.com
*/

/************************************************* OPTIONAL SETTINGS *********************************************************/

// Set your lookback period here. Allowed options include "TODAY", "YESTERDAY", "LAST_7_DAYS", "THIS_MONTH", "LAST_MONTH", "LAST_14_DAYS" or "LAST_30_DAYS". If you leave this blank, it will default to use "LAST_30_DAYS"
//
// Example: let lookbackPeriod = "LAST_30_DAYS";

let lookbackPeriod = "";

// Campaign name contains

// If you only want to look at some specific campaign(s), enter their name (or part
// of their name) in speech marks here. For example, ["RLSA"] would only look at
// campaigns with 'RLSA' in the name, while ["RLSA", "Competitors"] would only look
// at campaigns with either 'RLSA' OR 'Competitors' in their name. To include all
// campaigns, simply leave it as [""].
//
// Please note that campaign name filters cannot contain punctuation.
//
// Example: const campaignNameContains = ["Brand"];

const campaignNameContains = [""];

// Campaign name does not contain

// If you want to exclude any campaigns, enter their name (or part of their name)
// in speech marks here. For example, entering ["PMax"] would ignore any campaigns
// with 'PMax' in the name, while ["PMax", "Brand"] would ignore any campaigns with
// either 'PMax' or 'Brand' in the name. If you don't want to exclude any campaigns,
// just leave it as [""].
//
// Please note that campaign name filters cannot contain punctuation.
//
// Example: const campaignNameDoesNotContain = ["Generic"];

const campaignNameDoesNotContain = [""];

// If you'd like to send an email with a link to the spreadsheet, add your email address below.
// Otherwise, leave it blank.
//
// If you'd like to add more than one email address, just separate them with commas.
//
// Example:
// const emailAddresses = "aldo@example.com",
// const emailAddresses = "pia@example.com, kirandeep@example.com"

const emailAddresses = "";

/*****************************************************************************************************************************/
/*                      PLEASE DON'T TOUCH ANYTHING UNDER HERE OR UNSPEAKABLY TERRIBLE THINGS MAY HAPPEN                     */
/*****************************************************************************************************************************/

const accountName = AdsApp.currentAccount().getName();
const accountId = AdsApp.currentAccount().getCustomerId();
const title = accountName + " (" + accountId + ") - WeDiscover CPC Distribution Tool";
let emailLog = "";

function main() {
  // RE2 syntax that can't appear in campaign filters
  const regex = /[[:punct:]]|\.|\[|\]|\^|\:|\{|\}|\?|\,|\*|\-|\=|\+|\(|\)|\'|\"|\#|\@|\%|\$|\<|\!|\&/g;
  
  // Concatenate name filter arrays
  const allFilterArr = campaignNameContains.concat(campaignNameDoesNotContain);
  
  allFilterArr.forEach(element => { 
    if (element.search(regex) >= 0) {
      throw new Error("Campaign name filters cannot contain punctuation. Please remove this punctuation and try again.");
    }
  })
  
  const allowedDateRanges = ["TODAY", "YESTERDAY", "LAST_7_DAYS", "THIS_MONTH", "LAST_MONTH", "LAST_14_DAYS", "LAST_30_DAYS"];

  // Set lookback period to LAST_30_DAYS if it's not already one of the supported options
  if (!allowedDateRanges.includes(lookbackPeriod)) {
    lookbackPeriod = "LAST_30_DAYS";
  }

  let whereStatement = " WHERE metrics.clicks > 0 ";

  const campaignNameContainsLength = campaignNameContains.map((el) => el.trim().replace(/"/g, '\\"')).join("").length;
  const campaignNameDoesNotContainLength = campaignNameDoesNotContain
    .map((el) => el.trim().replace(/"/g, '\\"'))
    .join("").length;

  if (campaignNameContainsLength > 0) {
    const regexString = createRegexString(campaignNameContains);
    whereStatement += 'AND campaign.name REGEXP_MATCH "' + regexString + '" ';
  }

  if (campaignNameDoesNotContainLength > 0) {
    const regexString = createRegexString(campaignNameDoesNotContain);
    whereStatement += 'AND campaign.name NOT REGEXP_MATCH "' + regexString + '" ';
  }

  const queries = {};

  // Get all search terms over time period for the selected campaigns,
  // alongside their campaign name, clicks, conversions, conversion value and avg. cpc
  const report = AdsApp.report(
    "SELECT search_term_view.search_term, campaign.name, metrics.clicks, metrics.conversions, metrics.conversions_value, metrics.average_cpc, metrics.cost_micros" +
      " FROM search_term_view " +
      whereStatement +
      " AND segments.date DURING " +
      lookbackPeriod +
      " ORDER BY metrics.cost_micros DESC LIMIT 10000"
  );

  const rows = report.rows();

  while (rows.hasNext()) {
    const row = rows.next();

    const metrics = [
      row["search_term_view.search_term"],
      row["campaign.name"],
      row["metrics.clicks"] || 0,
      row["metrics.conversions"] || 0.0,
      row["metrics.conversions_value"] || 0.0,
      row["metrics.average_cpc"] / 1000000 || 0,
    ];

    queries[row["search_term_view.search_term"]] = metrics;
  }

  // Copy the CPC distribution tool template spreadsheet to Drive of user
  // Rename the copied template spreadsheet to include the account name
  const templateFileId = "1sOei4D5IDkhLAsFAzkhOhFIFZGZQo4LqnAzn-jyidtU";
  const ssFile = DriveApp.getFileById(templateFileId).makeCopy(title);
  const spreadsheetUrl = ssFile.getUrl();

  // Open spreadsheet file with SpreadsheetApp
  const ss = SpreadsheetApp.open(ssFile);
  const dataSheet = ss.getSheetByName("Data");

  if (!dataSheet) throw error;

  const numberOfQueries = Object.values(queries).length;

  // Spit all of the metrics into the "Data" sheet of the CPC distribution tool template spreadsheet
  const dataRange = dataSheet.getRange(2, 1, numberOfQueries, 6);
  dataRange.clear();
  dataRange.setValues(Object.values(queries));

  // Log the spreadsheet URL
  Logger.log("WeDiscover CPC Distribution Tool");
  Logger.log("--------------------------------");
  Logger.log(" ");
  logTextAndAppendToEmailLog("Lookback period: " + lookbackPeriod);
  logTextAndAppendToEmailLog("");
  logTextAndAppendToEmailLog("Your spreadsheet is here:");
  logTextAndAppendToEmailLog(spreadsheetUrl);

  // Email the recipients above (if applicable)
  // If email addresses are not falsy, send an email
  if (emailAddresses) {
    sendSummaryEmail(emailAddresses);
    Logger.log(" ");
    Logger.log("Email sent to: " + emailAddresses);
  }
}

function sendSummaryEmail(recipientEmails) {
  const emailIntroduction =
    "Hi there,<br><br>" +
    "This is your automated email from the WeDiscover CPC Distribution Tool.<br><br>" +
    "Below is the log for the recent script execution on " +
    accountName +
    " (" +
    accountId +
    ").<br>";

  const emailBody = emailLog;

  const emailFooter =
    "<br><br>" +
    "All the best,<br>" +
    "WeDiscover<br>" +
    "<br>" +
    "<strong>If you have any questions about this script, please email " +
    '<a href="mailto:scripts@we-discover.com">scripts@we-discover.com</a></strong>';

  const body = emailIntroduction + emailBody + emailFooter;

  if (recipientEmails) {
    // Try GmailApp first and use MailApp as a fallback if it doesn't work.
    try {
      GmailApp.sendEmail(recipientEmails, title, "", { htmlBody: body });
    } catch (e) {
      try {
        MailApp.sendEmail(recipientEmails, title, "", { htmlBody: body });
      } catch (e) {
        throw "Error sending the email: '" + e + "'. Please check the email addresses are valid.";
      }
    }
  }
}

function logTextAndAppendToEmailLog(text) {
  // If email addresses are set, add text to email log.
  if (emailAddresses) {
    emailLog += "<br>" + text;
  }

  Logger.log(text);
}

// Create an RE2 string for matching campaign names
function createRegexString(arr) {
  return arr.map((el) => `.*` + el.trim() + `.*`).join("|");
}
