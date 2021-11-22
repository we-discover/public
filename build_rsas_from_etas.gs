// EDIT ME -- Check for RSAs in Paused Ads, Ad Groups or Campaigns
var checkPausedCampaigns = false;
var checkPausedAdGroups = false;
var includePausedRsas = false;
var pullFromPausedEtas = false;

// Script entrypoint
function main() {
  Logger.log("Getting ad groups without RSAs...");
  var adGroupIdsWithoutRsas = getAdGroupIdsWithoutRsas();
  
  Logger.log(adGroupIdsWithoutRsas.length + " ad groups found without RSAs.");
  if(adGroupIdsWithoutRsas.length === 0) {
    Logger.log("Terminating script.");
    return;
  }
  
  var assetsQueryWithIds = assetsQueryTemplate.replace("*INSERT_AD_GROUP_IDS*", adGroupIdsWithoutRsas.join(","));
  Logger.log("Getting ETAs from ad groups without RSAs...");
  var assetsReport = AdsApp.report(assetsQueryWithIds).rows();
  
  var groupedAssets = {};
  
  Logger.log("Extracting assets...");
  while(assetsReport.hasNext()) {
    var reportRow = assetsReport.next();
    
    // Attributes  
    var adGroupId = reportRow['ad_group.id'];
    var adGroupName = reportRow['ad_group.name'];
    var campaignId = reportRow['campaign.id'];
    var campaignName = reportRow['campaign.name'];

    // Assets
    var headline1 = reportRow['ad_group_ad.ad.expanded_text_ad.headline_part1'];
    var headline2 = reportRow['ad_group_ad.ad.expanded_text_ad.headline_part2'];
    var headline3 = reportRow['ad_group_ad.ad.expanded_text_ad.headline_part3'];

    var description1 = reportRow['ad_group_ad.ad.expanded_text_ad.description'];
    var description2 = reportRow['ad_group_ad.ad.expanded_text_ad.description2'];

    var path1 = reportRow['ad_group_ad.ad.expanded_text_ad.path1'];
    var path2 = reportRow['ad_group_ad.ad.expanded_text_ad.path2'];
    var finalUrl = reportRow['ad_group_ad.ad.final_urls'];

    
    groupedAssets[adGroupId] = groupedAssets[adGroupId] || {
      'ad_group_id': adGroupId,
      'ad_group_name': adGroupName,
      'campaign_id': campaignId,
      'campaign_name': campaignName,
      'headlines': [],
      'descriptions': [],
      'path1s': [],
      'path2s': [],
      'final_urls': []
     };
    
    pushElementsIfNeeded(groupedAssets[adGroupId]['headlines'], [headline1, headline2, headline3]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['descriptions'], [description1, description2]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['path1s'], [path1]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['path2s'], [path2]);
    pushElementsIfNeeded(groupedAssets[adGroupId]['final_urls'], [finalUrl]);
    
  }
  
  Logger.log("Counting assets...");
  var maxNoOfAssets = getMaxAssets(groupedAssets);
  
  Logger.log("Making Google Sheet...");
//  var ss = SpreadsheetApp.create("Test");
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1D1gASkVDWM0C_RHwwk_f8R-Gij0xlJu6OF81jg9xWSI/edit#gid=0");
  var sheet = ss.getSheets()[0];
  
  Logger.log("Setting headers...");
  setHeaders(sheet, maxNoOfAssets);
  
  Logger.log("Pushing asset values...");
  pushDataToSheet(sheet, groupedAssets, maxNoOfAssets)
  
  Logger.log("-----------------------------\nScript run complete. Google Sheet location: " + ss.getUrl());
}


// ========= UTILITY FUNCTIONS ==========================================================================================

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
  
  // Font weight row to allow us to set the headers to bold
  var boldings = [];
  for (var i = 0; i < numColumns; i++) {
    boldings.push("bold");
  }
  
  // Set values and bolding
  sheet.getRange(1, 1, 1, numColumns).setValues([headers]).setFontWeights([boldings]);
  
  // Autosize columns
  sheet.autoResizeColumns(1, numColumns);
  
  return null;
}

// Set outputs for sheet
function pushDataToSheet(sheet, groupedAssets, maxNoAssets) {
  var outputData = []
  
  for (var adGroupId in groupedAssets) {
    var data = groupedAssets[adGroupId];
    
    var row = ["-", "-", data['campaign_name'], data['campaign_id'], data['ad_group_name'], data['ad_group_id']];
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
    
//    outputData.push(row);
    sheet.appendRow(row);
  }
  
//  Logger.log(outputData[0]);
//  sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  
}

// Push elements of an array into another array, if new elements are not null and not present in the original array
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
    'max_urls': 0
  }
  
  for (var adGroupId in obj) {
    maxs['max_headlines'] = Math.max(maxs['max_headlines'], obj[adGroupId]['headlines'].length);
    maxs['max_descriptions'] = Math.max(maxs['max_descriptions'], obj[adGroupId]['descriptions'].length);
    maxs['max_path1s'] = Math.max(maxs['max_path1s'], obj[adGroupId]['path1s'].length);
    maxs['max_path2s'] = Math.max(maxs['max_path2s'], obj[adGroupId]['path2s'].length);
    maxs['max_urls'] = Math.max(maxs['max_urls'], obj[adGroupId]['final_urls'].length);
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
    ad_group_ad.ad.final_urls, \
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
