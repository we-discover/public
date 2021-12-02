/*
    Name:        WeDiscover - ETA To RSA Builder, Google Ads Script
    Description: A script to build RSA ad copy from existing ETA ad
                 copy.
    License:     https://github.com/we-discover/public/blob/master/LICENSE
    Version:     1.0.0
    Released:    2021-12-13
    Contact:     scripts@we-discover.com
*/

// EDIT ME -- Check for RSAs in Paused Ads, Ad Groups or Campaigns
var checkPausedCampaigns = false;
var checkPausedAdGroups = false;
var includePausedRsas = false;
var pullFromPausedEtas = false;

// Script entrypoint
function main(){
  var executionContext = getExecutionContext();
  var topLevelAccountName = AdsApp.currentAccount().getName();
  var topLevelAccountId = AdsApp.currentAccount().getCustomerId();
  var ssName = topLevelAccountName + " (" + topLevelAccountId + ") | ETA to RSA Builder";
  
  Logger.log("Making Google Sheet...");
  var ss = SpreadsheetApp.create(ssName);
  
  // If MCC, run data process on a loop through all accounts
  if (executionContext === 'manager_account') {
    var accountIterator = AdsManagerApp.accounts().get();
    while(accountIterator.hasNext()) {
      AdsManagerApp.select(accountIterator.next());
      
      var accountName = AdsApp.currentAccount().getName();
      var accountId = AdsApp.currentAccount().getCustomerId();
      var accountNameAndId = accountName + " (" + accountId + ")";  
      var successCount = 0;
      
      Logger.log("Looking in account " + accountNameAndId + " ...");
      var assets = getAssets(accountName, accountId);
      if (assets !== 'no_rsas') {
        var sheet = ss.insertSheet(accountNameAndId)
        var maxNoOfAssets = getMaxAssets(assets);
        writeToSheet(sheet, assets, maxNoOfAssets, accountName, accountId);
        
        successCount++;
      }
      
      else if (assets === 'no_rsas' && accountIterator.hasNext()) {
        Logger.log("Trying next account...");
      }
      else if (assets === 'no_rsas' && !accountIterator.hasNext()) {
        Logger.log("No more accounts found. Terminating script.");
      }   
       
      Logger.log("--------------------------------------------------");
    }
    
    // Remove first sheet, which is blank
    ss.deleteSheet(ss.getSheets()[0]);
    if (successCount === 0) {
      Logger.log("No ad groups missing RSAs found. Google Sheet can be safely deleted.");
    }
    Logger.log("Script run complete. Google Sheet location: " + ss.getUrl());
  }
  
  // If child account, process on that account only
  else if (executionContext === 'client_account') {
    var accountName = AdsApp.currentAccount().getName();
    var accountId = AdsApp.currentAccount().getCustomerId();
    var accountNameAndId = accountName + " (" + accountId + ")";
    var sheet = ss.getSheets()[0].setName(accountNameAndId);
    
    var assets = getAssets(accountName, accountId);
    if (assets === 'no_rsas') {
      Logger.log("Terminating script. Google Sheet can be safely deleted.");
    }
    
    else if (assets !== 'no_rsas') {
      var maxNoOfAssets = getMaxAssets(assets);
      writeToSheet(sheet, assets, maxNoOfAssets, accountName, accountId);
    }
    
    Logger.log("--------------------------------------------------");
    Logger.log("Script run complete. Google Sheet location: " + ss.getUrl()); 

  }
}

// ========= UTILITY FUNCTIONS ==========================================================================================

// Determine the type of account in which the script is running
function getExecutionContext() {
    if (typeof AdsManagerApp != "undefined") {
        return 'manager_account';
    }
    return 'client_account';
}

// Collect ETA assets and group by ad group ID
function getAssets(accountName, accountId) {
  Logger.log("Getting ad groups without RSAs...");
  var adGroupIdsWithoutRsas = getAdGroupIdsWithoutRsas();
  
  Logger.log(adGroupIdsWithoutRsas.length + " ad groups found without RSAs.");
  if(adGroupIdsWithoutRsas.length === 0) {
    Logger.log("No RSAs found in " + accountName + " (" + accountId + ").");
    return "no_rsas";
  }
  
  var assetsQueryWithIds = assetsQueryTemplate.replace("*INSERT_AD_GROUP_IDS*", adGroupIdsWithoutRsas.join(","));
  Logger.log("Getting ETAs from ad groups without RSAs...");
  var assetsReport = AdsApp.search(assetsQueryWithIds);
  
  var groupedAssets = {};
  
  Logger.log("Extracting assets...");
  while(assetsReport.hasNext()) {
    var reportRow = assetsReport.next();
    
    // Attributes  
    var adGroupId = reportRow.adGroup.id;
    var adGroupName = reportRow.adGroup.name;
    var campaignId = reportRow.campaign.id;
    var campaignName = reportRow.campaign.name;

    // Assets
    var headline1 = reportRow.adGroupAd.ad.expandedTextAd.headlinePart1;
    var headline2 = reportRow.adGroupAd.ad.expandedTextAd.headlinePart2;
    var headline3 = reportRow.adGroupAd.ad.expandedTextAd.headlinePart3;

    var description1 = reportRow.adGroupAd.ad.expandedTextAd.description;
    var description2 = reportRow.adGroupAd.ad.expandedTextAd.description2;

    var path1 = reportRow.adGroupAd.ad.expandedTextAd.path1;
    var path2 = reportRow.adGroupAd.ad.expandedTextAd.path2;
    var finalUrl = reportRow.adGroupAd.ad.finalUrls[0];
    
    // Mobile URL not always present
    var finalMobileUrl = "";
    if (reportRow.adGroupAd.ad.finalMobileUrls) {
      var finalMobileUrl = reportRow.adGroupAd.ad.finalMobileUrls[0];
    }
     
    groupedAssets[adGroupId] = groupedAssets[adGroupId] || {
      'ad_group_id': adGroupId,
      'ad_group_name': adGroupName,
      'campaign_id': campaignId,
      'campaign_name': campaignName,
      'headlines': [],
      'descriptions': [],
      'path1s': [],
      'path2s': [],
      'final_urls': [],
      'final_mobile_urls': []
     };
    
    pushElementsIfNeeded(groupedAssets[adGroupId]['headlines'], [headline1, headline2, headline3]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['descriptions'], [description1, description2]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['path1s'], [path1]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['path2s'], [path2]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['final_urls'], [finalUrl]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['final_mobile_urls'], [finalMobileUrl]);
    
  }
  
  return groupedAssets;
}

// Write headers and body to sheet
function writeToSheet(sheet, groupedAssets, maxNoOfAssets, accountName, accountId) {
  setHeaders(sheet, maxNoOfAssets);
  pushDataToSheet(sheet, groupedAssets, maxNoOfAssets, accountName, accountId);
}

// Set headers in a sheet
function setHeaders(sheet, maxNoAssets) {
  
  // Attribute headers
  var headers = ["Account Name", "Account ID", "Campaign Name", "Campaign ID", "Ad Group Name", "Ad Group ID"];
  var numColumns = headers.length;
  
  // Asset headers (of variable length)
  for (var i = 0; i < maxNoAssets['max_headlines']; i++) {
    headers.push("Headline " + (i+1));
  }
  numColumns += maxNoAssets['max_headlines'];
  
  for (var i = 0; i < maxNoAssets['max_descriptions']; i++) {
    headers.push("Description " + (i+1));
  }
  numColumns += maxNoAssets['max_descriptions'];

  headers.push("Path 1(s)");
  for (var i = 1; i < maxNoAssets['max_path1s']; i++) {
    headers.push("...");
  }
  numColumns += Math.max(1,maxNoAssets['max_path1s'], 1);
  
  headers.push("Path 2(s)");
  for (var i = 1; i < maxNoAssets['max_path2s']; i++) {
    headers.push("...");
  }
  numColumns += Math.max(1,maxNoAssets['max_path2s'], 1);
 
  headers.push("Final URL(s)");
  for (var i = 1; i < maxNoAssets['max_urls']; i++) {
    headers.push("...");
  }
  numColumns += Math.max(1, maxNoAssets['max_urls']);
  
  headers.push("Final Mobile URL(s)");
  for (var i = 1; i < maxNoAssets['max_mobile_urls']; i++) {
    headers.push("...");
  }
  numColumns += Math.max(1, maxNoAssets['max_mobile_urls']);
  
  // Font weight row to allow us to set the headers to bold
  var boldings = [];
  for (var i = 0; i < numColumns; i++) {
    boldings.push("bold");
  }
  
  // Set values and bolding
  sheet.getRange(1, 1, 1, numColumns).setValues([headers]).setFontWeights([boldings]);
  
  return null;
}

// Set outputs for sheet
function pushDataToSheet(sheet, groupedAssets, maxNoAssets, accountName, accountId) {
  var outputData = []
  
  for (var adGroupId in groupedAssets) {
    var data = groupedAssets[adGroupId];
    
    var row = [accountName, accountId, data['campaign_name'], data['campaign_id'], data['ad_group_name'], data['ad_group_id']];
    
    // For each type of asset, must blank fill columns where the ad group has less than the maximum number of that asset type
    // This ensures columns all line up with their headers
    row.push.apply(row, data['headlines'])
    for (var i = data['headlines'].length; i < maxNoAssets['max_headlines']; i++) {
      row.push("");
    }
    
    row.push.apply(row, data['descriptions'])
    for (var i = data['descriptions'].length; i < maxNoAssets['max_descriptions']; i++) {
      row.push("");
    }
    row.push.apply(row, data['path1s'])
    for (var i = data['path1s'].length; i < maxNoAssets['max_path1s']; i++) {
      row.push("");
    }
    row.push.apply(row, data['path2s'])
    for (var i = data['path2s'].length; i < maxNoAssets['max_path2s']; i++) {
      row.push("");
    }
    row.push.apply(row, data['final_urls'])
    for (var i = data['final_urls'].length; i < maxNoAssets['max_urls']; i++) {
      row.push("");
    }
    row.push.apply(row, data['final_mobile_urls'])
    for (var i = data['final_mobile_urls'].length; i < maxNoAssets['max_mobile_urls']; i++) {
      row.push("");
    }
    
    outputData.push(row);
  }
  
  // Push to sheet and autosize columns
  sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  sheet.autoResizeColumns(1, outputData[0].length);
  
  return null;
}

// Push elements of an array into another array, if new elements are not null/undefined and not present in the original array
function pushElementsIfNeeded(arr, elements) {
  for (var i = 0; i < elements.length; i++) {
    var el = elements[i]; 
    
    if (arr.indexOf(el) === -1 && el) {
      arr.push(el);
    }
  }

  return null;
}

// Get the max number of each asset type present in an ad group
function getMaxAssets(obj) {
  var maxs = {
    'max_headlines': 0,
    'max_descriptions': 0,
    'max_path1s': 0,
    'max_path2s': 0,
    'max_urls': 0,
    'max_mobile_urls': 0
  }
  
  for (var adGroupId in obj) {
    maxs['max_headlines'] = Math.max(maxs['max_headlines'], obj[adGroupId]['headlines'].length);
    maxs['max_descriptions'] = Math.max(maxs['max_descriptions'], obj[adGroupId]['descriptions'].length);
    maxs['max_path1s'] = Math.max(maxs['max_path1s'], obj[adGroupId]['path1s'].length);
    maxs['max_path2s'] = Math.max(maxs['max_path2s'], obj[adGroupId]['path2s'].length);
    maxs['max_urls'] = Math.max(maxs['max_urls'], obj[adGroupId]['final_urls'].length);
    maxs['max_mobile_urls'] = Math.max(maxs['max_mobile_urls'], obj[adGroupId]['final_mobile_urls'].length);
  }
  
  return maxs;
}

// Returns a list of Ad Group IDs that correspond to a given condition
function getAdGroupsWithCondition(condition) {
    var query = adGroupQueries[condition];
    var ids = [];

    var iterator = AdsApp.report(query).rows();
    while (iterator.hasNext()) {
        var row = iterator.next();
        ids.push(row['ad_group.id']);
    }

    return ids;
}

// Process to extract entities without RSAs from a given account
function getAdGroupIdsWithoutRsas() {

    var allAdGroupIds = getAdGroupsWithCondition('all_ad_groups');
    var adGroupIdsWithRsa = getAdGroupsWithCondition('ad_groups_with_rsas');
    var adGroupIdsWithoutRsa = allAdGroupIds.filter(function(id) {
        return adGroupIdsWithRsa.indexOf(id) === -1
    });
  
    return adGroupIdsWithoutRsa;
}

// ========= GAQL QUERIES ==============================================================================================

var today = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");

var constraintsForQueries = (" \
    campaign.advertising_channel_type = 'SEARCH' \
    AND ad_group.type IN ('SEARCH_STANDARD') \
    AND campaign.status IN ('ENABLED'" + (checkPausedCampaigns ? ", 'PAUSED'" : "") + ") \
    AND ad_group.status IN ('ENABLED'" + (checkPausedAdGroups ? ", 'PAUSED'" : "") + ") \
    AND campaign.end_date >= '" + today + "'"
);

var queryAllAdGroups = (" \
    SELECT \
        ad_group.id \
    FROM \
        ad_group \
    WHERE \
        " + constraintsForQueries
  ).replace(/ +(?= )/g, '');

var queryAdGroupsWithRsa = (" \
    SELECT \
        ad_group.id \
    FROM \
        ad_group_ad \
    WHERE \
        " + constraintsForQueries + " \
        AND ad_group_ad.status IN ('ENABLED'" + (includePausedRsas ? ", 'PAUSED'" : "") + ") \
        AND ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD'"
  ).replace(/ +(?= )/g, '');

var assetsQueryTemplate = (" \
  SELECT \
    ad_group_ad.ad.expanded_text_ad.headline_part1, ad_group_ad.ad.expanded_text_ad.headline_part2, ad_group_ad.ad.expanded_text_ad.headline_part3, \
    ad_group_ad.ad.expanded_text_ad.description, ad_group_ad.ad.expanded_text_ad.description2, \
    ad_group_ad.ad.expanded_text_ad.path1, ad_group_ad.ad.expanded_text_ad.path2, \
    ad_group_ad.ad.final_urls, ad_group_ad.ad.final_mobile_urls, \
    campaign.name, campaign.id, ad_group.name, ad_group.id \
  FROM \
    ad_group_ad \
  WHERE \
    ad_group.id IN (*INSERT_AD_GROUP_IDS*) \
    AND ad_group_ad.ad.type = 'EXPANDED_TEXT_AD' \
    AND ad_group_ad.status IN ('ENABLED'" + (pullFromPausedEtas ? ", 'PAUSED'" : "") + ")"
).replace(/ +(?= )/g, '');

var adGroupQueries = {
    'all_ad_groups': queryAllAdGroups,
    'ad_groups_with_rsas': queryAdGroupsWithRsa
};
