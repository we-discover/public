/*
    Name:           WeDiscover - Cross Negative Script
    Description:    This script allows you to add cross negative keywords to campaigns or ad groups
                    with the option to use the original match type or convert all to exact match.
    License:        https://github.com/we-discover/public/blob/master/LICENSE
    Version:        1.2.0
    Released:       2025-01-01
    Author:         Nathan Ifill (@nathanifill)
    Contact:        scripts@we-discover.com
*/

/************************************************* SETTINGS **************************************************************/

//  Options to choose between campaign-level and ad group-level negative keywords
const ADD_CAMPAIGN_LEVEL_NEGATIVES = true;
const ADD_AD_GROUP_LEVEL_NEGATIVES = true;
const MAX_NEGATIVE_KEYWORDS = 5000; // This determines the maximum number of negative keywords to add to a campaign or ad group

//  Option to choose negative keyword match type
const USE_ORIGINAL_MATCH_TYPE = false; // Set to false to use negative exact match for all or true to keep original match type

/* 
    Option to filter which campaigns to include or exclude by name.
    Note: Matching is case-insensitive, so 'Campaign', 'cAmpAigN' and 'campaign' are treated the same.
    
    Please note - if you leave this blank, the script will work across all campaigns and/or ad groups. Leaving both of these
    blank is NOT recommended as (depending on the size of your account) the script may not complete before Google's
    30-minute execution limit.
 */
const INCLUDE_CAMPAIGN_NAMES = []; // Example: ['Brand', 'Generic'] or []
const EXCLUDE_CAMPAIGN_NAMES = []; // Example: ['U.S.', 'Christmas'] or []

//  Option to enable or disable logging for faster execution
const ENABLE_LOGGING = true; // Set to true to log what changes the script is making or set to false to speed up the script

//  Option to generate spreadsheet log and send it to you via email
const ENABLE_SPREADSHEET_LOG = true; // Set to true to generate a spreadsheet detailing the changes and email you the log
const EMAILS = ""; // Enter emails to send spreadsheet log to. If you don't want the email, leave this blank, e.g. ""
// To enter multiple email addresses, separate them with commas
// Example: "pia@example.com" or "indi@example.com,nadia@example.com".

/*************************************************************************************************************************/

let ss;
let ssId;
let body = "";
const ssData = {};
const currentAccount = AdsApp.currentAccount();
const accountName = currentAccount.getName();
const accountId = currentAccount.getCustomerId();

function main() {
  log(`Starting process...`);
  log("");
  log(`Add campaign-level negatives? ${ADD_CAMPAIGN_LEVEL_NEGATIVES}`);
  log(`Add ad group-level negatives? ${ADD_AD_GROUP_LEVEL_NEGATIVES}`);
  log(`Use original match type for negatives? ${USE_ORIGINAL_MATCH_TYPE}`);
  log("");

  if (ENABLE_SPREADSHEET_LOG) {
    const ssTemplateUrl = "https://docs.google.com/spreadsheets/d/18zlEN-7ddxBndhfNnGG1WgWqBfugTcPSikyFeUiD-Xc/edit?gid=0#gid=0";
    const ssTemplate = SpreadsheetApp.openByUrl(ssTemplateUrl);
    ss = ssTemplate.copy(`${accountName} (${accountId}) Cross Negative Script | WeDiscover`);
    ssId = ss.getId();
  }

  if (INCLUDE_CAMPAIGN_NAMES.length > 0) {
    log(`Campaign name must contain: ${INCLUDE_CAMPAIGN_NAMES}`);
  }
  if (EXCLUDE_CAMPAIGN_NAMES.length > 0) {
    log(`Campaign name must not contain: ${EXCLUDE_CAMPAIGN_NAMES}`);
  }
  log("");

  if (ADD_CAMPAIGN_LEVEL_NEGATIVES) {
    log(`ADDING CAMPAIGN-LEVEL NEGATIVES:`);
    log("");
    addCampaignLevelNegatives();
  }
  if (ADD_AD_GROUP_LEVEL_NEGATIVES) {
    log(`ADDING AD GROUP-LEVEL NEGATIVES:`);
    log("");
    addAdGroupLevelNegatives();
  }

  if (ENABLE_SPREADSHEET_LOG) {
    let rowCount = 2;
    const sheet = ss.getSheetByName("DATA DUMP");

    for (let [keyword, entities] of Object.entries(ssData)) {
      sheet.getRange(1,1,1,2).setValues([["Negative Keyword", "Entity"]]).setFontWeights([['bold', 'bold']]);
      sheet.getRange(rowCount, 1, 1, 2).setValues([[keyword, entities]]);
      rowCount++;
    }

    sheet.setFrozenRows(1);
    //sheet.deleteColumns(3, sheet.getMaxColumns());
    // Sort the data by keyword ascending
    sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn()).sort({column: 1, ascending: true});

    log("");
    log("Get your spreadsheet log here: " + ss.getUrl());
    log("");

    try {
      const subject = `[CNS] Cross Negative Script Log for ${accountName} (${accountId}) | WeDiscover`;
      // If no email address specified and log requested, use the owner of the spreadsheet
      const userEmail = EMAILS ? EMAILS : ss.getOwner().getEmail();
      // Turn the user entered string into a string array String[] with trimmed email addresses
      const userEmailArr = userEmail.split(",").map(el => el.trim());
      // Set access of spreadsheet log to the email addresses
      DriveApp.getFileById(ssId).addEditors(userEmailArr);
      MailApp.sendEmail(userEmail, subject, body);
    } catch (e) {
      throw("Unable to send Cross Negative Script Log email to the following emails: " + userEmailArr);
    }
  }
  
  log("Script processing complete. Have a nice day!");
}

/**
 * Adds negative keywords at the campaign level.
 */
function addCampaignLevelNegatives() {
  const campaigns = [];
  const campaignIterator = AdsApp.campaigns()
    .withCondition("campaign.status = ENABLED")
    .withCondition("campaign.experiment_type = BASE")
    .get();
  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    if (shouldProcessCampaign(campaign.getName())) {
      const keywords = getKeywords(campaign);
      campaigns.push({ campaign, keywords });
    }
  }
  campaigns.forEach(({ campaign, keywords }) => {
    log(`Adding ${keywords.length} keywords from ${campaign.getName()} to all of the other campaigns.`);
    log("");
    campaigns.forEach(({ campaign: otherCampaign }) => {
      if (campaign.getId() !== otherCampaign.getId()) {
        log(`  - ${otherCampaign.getName()}`);
        addNegativesToEntity(otherCampaign, keywords);
      }
    });
    log("");
    log(`Adding complete.`);
    log("");
  });
}

/**
 * Adds negative keywords at the ad group level.
 */
function addAdGroupLevelNegatives() {
  const campaignIterator = AdsApp.campaigns()
    .withCondition("campaign.status = ENABLED")
    .withCondition("campaign.experiment_type = BASE")
    .get();
  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    if (shouldProcessCampaign(campaign.getName())) {
      log(`Now processing the ${campaign.getName()} campaign.`);
      log("");
      const adGroups = [];
      const adGroupIterator = campaign
        .adGroups()
        .withCondition("ad_group.status = ENABLED")
        .withCondition("campaign.status = ENABLED")
        .get();
      while (adGroupIterator.hasNext()) {
        const adGroup = adGroupIterator.next();
        const keywords = getKeywords(adGroup);
        adGroups.push({ adGroup, keywords });
      }
      adGroups.forEach(({ adGroup, keywords }) => {
        if (keywords.length > 0) {
          log(
            `Adding ${
              keywords.length
            } keywords from the ${adGroup.getName()} ad group to the other ad groups in the ${campaign.getName()} campaign.`
          );
          log("");
          adGroups.forEach(({ adGroup: otherAdGroup }) => {
            if (adGroup.getId() !== otherAdGroup.getId()) {
              addNegativesToEntity(otherAdGroup, keywords);
            }
          });
          log(`Adding complete.`);
          log("");
        }
      });
    }
  }
}

/**
 * Determines if a campaign should be processed based on its name.
 */
function shouldProcessCampaign(campaignName) {
  const includeRegex = INCLUDE_CAMPAIGN_NAMES.length > 0 ? new RegExp(INCLUDE_CAMPAIGN_NAMES.join("|"), "i") : null;
  const excludeRegex = EXCLUDE_CAMPAIGN_NAMES.length > 0 ? new RegExp(EXCLUDE_CAMPAIGN_NAMES.join("|"), "i") : null;

  const isIncluded = includeRegex ? includeRegex.test(campaignName) : true;
  const isExcluded = excludeRegex ? excludeRegex.test(campaignName) : false;

  return isIncluded && !isExcluded;
}

/**
 * Retrieves keywords from an entity.
 */
function getKeywords(entity) {
  const keywords = [];
  const keywordIterator = entity.keywords().withCondition("Status = ENABLED").orderBy("Cost DESC").get();
  while (keywordIterator.hasNext() && keywords.length < MAX_NEGATIVE_KEYWORDS) {
    const keyword = keywordIterator.next();
    keywords.push({
      text: keyword.getText(),
      matchType: keyword.getMatchType(),
    });
  }
  return keywords;
}

/**
 * Adds negative keywords to an entity.
 */

function addNegativesToEntity(entity, keywords) {
  keywords.forEach(({ text, matchType }) => {
    const negativeMatchType = USE_ORIGINAL_MATCH_TYPE ? matchType : "EXACT";
    let formattedText;
    switch (negativeMatchType) {
      case "BROAD":
        formattedText = text;
        createNegativeKeyword(formattedText, entity);
        break;
      case "PHRASE":
        formattedText = `"${text}"`;
        createNegativeKeyword(formattedText, entity);
        break;
      case "EXACT":
        formattedText = `[${text}]`;
        createNegativeKeyword(formattedText, entity);
        break;
    }
  });
}

function createNegativeKeyword(formattedText, entity) {
  entity.createNegativeKeyword(formattedText);
  ssData[formattedText] = ssData[formattedText] ? ssData[formattedText] : [];
  const entityType = entity.getEntityType();
  let entityObj;

  if (entityType === "AdGroup") {
    //entityObj = { campaign: entity.getBaseCampaign().getName(), adGroup: entity.getBaseAdGroup().getName() };
    entityObj = `${entity.getBaseCampaign().getName()} [${entity.getBaseAdGroup().getName()}]`;
  }

  if (entityType === "Campaign") {
    //entityObj = { campaign: entity.getBaseCampaign().getName() };
    entityObj = `${entity.getBaseCampaign().getName()}`;
  }

  // If spreadsheet data doesn't already have this entity in it, add it to the list
  if (!ssData[formattedText].includes(entityObj)) {
    ssData[formattedText].push(entityObj);
  }
}

/**
 * Logs messages if logging is enabled.
 */
function log(message) {
  body += message + "\n";

  if (ENABLE_LOGGING) {
    Logger.log(message);
  }
}
