/*
    Name:           WeDiscover - Cross Negative Script
    Description:    This script allows you to add cross negative keywords to campaigns or ad groups
                    with the option to use the original match type or convert all to exact match.
    License:        https://github.com/we-discover/public/blob/master/LICENSE
    Version:        1.0.0
    Released:       2024-10-16
    Author:         Nathan Ifill (@nathanifill)
    Contact:        scripts@we-discover.com
*/

/************************************************* SETTINGS **************************************************************/

// Options to choose between campaign-level and ad group-level negative keywords
const ADD_CAMPAIGN_LEVEL_NEGATIVES = true;
const ADD_AD_GROUP_LEVEL_NEGATIVES = true;
const MAX_NEGATIVE_KEYWORDS = 5000; // This determines the maximum number of negative keywords to add to a campaign or ad group

// Option to choose negative keyword match type
const USE_ORIGINAL_MATCH_TYPE = false; // Set to false to use negative exact match for all or true to keep original match type

// Option to filter which campaigns to include or exclude by name. If you leave this blank, the script will work across all
// campaigns and/or ad groups. Leaving both of these blank is not recommended as (depending on the size of your account)
// the script may not complete before Google's 30-minute execution limit.
const INCLUDE_CAMPAIGN_NAMES = ['ZA > Search']; // Example: ['IncludePart1', 'IncludePart2']
const EXCLUDE_CAMPAIGN_NAMES = []; // Example: ['ExcludePart1', 'ExcludePart2']

// Option to enable or disable logging for faster execution
const ENABLE_LOGGING = false; // Set to true to log what changes the script is making or set to false to speed up the script

/*************************************************************************************************************************/

function main() {
  log("Starting process...");
  log(" ");
  log("Add campaign-level negatives? " + ADD_CAMPAIGN_LEVEL_NEGATIVES);
  log("Add ad group-level negatives? " + ADD_AD_GROUP_LEVEL_NEGATIVES);
  log("Use original match type for negatives? " + USE_ORIGINAL_MATCH_TYPE);
  log("");
  if (INCLUDE_CAMPAIGN_NAMES.length > 0) {
    log("Campaign name must contain: " + INCLUDE_CAMPAIGN_NAMES);
  }
  if (EXCLUDE_CAMPAIGN_NAMES.length > 0) {
    log("Campaign name must not contain: " + EXCLUDE_CAMPAIGN_NAMES);
  }

  if (ADD_CAMPAIGN_LEVEL_NEGATIVES) {
    log("Adding campaign-level negatives:".toUpperCase());
    log(" ");
    addCampaignLevelNegatives();
  }
  if (ADD_AD_GROUP_LEVEL_NEGATIVES) {
    log("Adding ad group-level negatives:".toUpperCase());
    log(" ");
    addAdGroupLevelNegatives();
  }

  log("Script processing complete. Have a nice day!");
}

/**
 * Adds negative keywords at the campaign level.
 */
function addCampaignLevelNegatives() {
  const campaigns = [];
  const campaignIterator = AdsApp.campaigns().withCondition("Status = ENABLED").get();
  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    if (shouldProcessCampaign(campaign.getName())) {
      const keywords = getKeywords(campaign);
      campaigns.push({ campaign, keywords });
    }
  }
  campaigns.forEach(({ campaign, keywords }) => {
    log("Adding " + keywords.length + " keywords from " + campaign.getName() + " to all of the other campaigns.");
    log("");
    campaigns.forEach(({ campaign: otherCampaign }) => {
      if (campaign.getId() !== otherCampaign.getId()) {
        log("  - " + otherCampaign.getName());
        addNegativesToCampaign(otherCampaign, keywords);
      }
    });
    log(" ");
    log("Adding complete.");
    log(" ");
  });
}

/**
 * Adds negative keywords at the ad group level.
 */
function addAdGroupLevelNegatives() {
  const campaignIterator = AdsApp.campaigns().withCondition("Status = ENABLED").get();
  while (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    if (shouldProcessCampaign(campaign.getName())) {
      log("Now processing the " + campaign.getName() + " campaign.");
      log(" ");
      const adGroups = [];
      const adGroupIterator = campaign.adGroups().withCondition("Status = ENABLED").get();
      while (adGroupIterator.hasNext()) {
        const adGroup = adGroupIterator.next();
        const keywords = getKeywords(adGroup);
        adGroups.push({ adGroup, keywords });
      }
      adGroups.forEach(({ adGroup, keywords }) => {
        if (keywords.length > 0) {
          log(
            "Adding " +
              keywords.length +
              " keywords from the " +
              adGroup.getName() +
              " ad group to the other ad groups in the " +
              campaign.getName() +
              " campaign."
          );
          log("");
          adGroups.forEach(({ adGroup: otherAdGroup }) => {
            if (adGroup.getId() !== otherAdGroup.getId()) {
              addNegativesToAdGroup(otherAdGroup, keywords);
            }
          });
          log("Adding complete.");
          log(" ");
        }
      });
    }
  }
}

/**
 * Determines if a campaign should be processed based on its name.
 * @param {string} campaignName - The name of the campaign.
 * @returns {boolean} - True if the campaign should be processed, false otherwise.
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
 * @param {Object} entity - The entity from which to retrieve keywords.
 * @returns {Array} - An array of keywords with their text and match type.
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
 * Adds negative keywords to a campaign.
 * @param {Object} campaign - The campaign to which negative keywords will be added.
 * @param {Array} keywords - The keywords to add as negatives.
 */
function addNegativesToCampaign(campaign, keywords) {
  keywords.forEach(({ text, matchType }) => {
    const negativeMatchType = USE_ORIGINAL_MATCH_TYPE ? matchType : "EXACT";
    switch (negativeMatchType) {
      case "BROAD":
        campaign.createNegativeKeyword(text);
        break;
      case "PHRASE":
        campaign.createNegativeKeyword('"' + text + '"');
        break;
      case "EXACT":
        campaign.createNegativeKeyword('[' + text + ']');
        break;
    }
  });
}

/**
 * Adds negative keywords to an ad group.
 * @param {Object} adGroup - The ad group to which negative keywords will be added.
 * @param {Array} keywords - The keywords to add as negatives.
 */
function addNegativesToAdGroup(adGroup, keywords) {
  keywords.forEach(({ text, matchType }) => {
    const negativeMatchType = USE_ORIGINAL_MATCH_TYPE ? matchType : "EXACT";
    switch (negativeMatchType) {
      case "BROAD":
        adGroup.createNegativeKeyword(text);
        break;
      case "PHRASE":
        adGroup.createNegativeKeyword('"' + text + '"');
        break;
      case "EXACT":
        adGroup.createNegativeKeyword('[' + text + ']');
        break;
    }
  });
}

/**
 * Logs messages if logging is enabled.
 * @param {string} message - The message to log.
 */
function log(message) {
  if (ENABLE_LOGGING) {
    Logger.log(message);
  }
}
