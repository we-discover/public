// Filter out ended experiments if poss
// Daily changes only or all ad groups


// Inputs:
// Must be set to true or false
var checkPausedCampaigns = false;
var checkPausedAdGroups = false;
var checkPausedAds = false;
var reportDailyChangeOnly = true;

// Must be text enclosed within quote marks ""
var recipientEmail = "example@gmail.com";
var labelName = "Ad groups without RSAs"

/******************************************************
        DO NOT EDIT ANYTHING BELOW THIS LINE
******************************************************/

var allAdGroupsQuery =
    "SELECT " +
    "ad_group.id " +
    "FROM " +
    "ad_group " +
    "WHERE " +
    "campaign.status IN ('ENABLED'" + (checkPausedCampaigns ? ", 'PAUSED'" : "") + ") " +
    "AND ad_group.status IN ('ENABLED'" + (checkPausedAdGroups ? ", PAUSED'" : "") + ")" +
    "AND campaign.advertising_channel_type = 'SEARCH'";

var rsaAdGroupsQuery = 
    "SELECT " +
    "ad_group.id " +
    "FROM " +
    "ad_group_ad " +
    "WHERE " +
    "ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD' " +
    "AND campaign.status IN ('ENABLED'" + (checkPausedCampaigns ? ", 'PAUSED'" : "") + ") " +
    "AND ad_group.status IN ('ENABLED'" + (checkPausedAdGroups ? ", PAUSED'" : "") + ")" +
    "AND ad_group_ad.status IN ('ENABLED'" + (checkPausedAds ? ", PAUSED'" : "") + ")" +
    "AND campaign.advertising_channel_type = 'SEARCH'";

var emailOpener = 
    "Hi there,<br><br>" +
    "This is your automated WeDiscover RSA Ad Group Checker.<br><br>" +
    "Included below are the details of ad groups without RSAs in your account(s). The label '<i>" + labelName + "</i>' has been applied to all ad groups listed for your convenience.<br><br>";

var emailFooter =
    "All the best,<br>" +
    "WeDiscover<br>" +
    "<br>" +
    "*If you have any questions about this script, please email <a href = \"mailto:scripts@we-discover.com\">scripts@we-discover.com</a>";

function main() {
  var mainEmailBody = "";
  var emailSubject = AdsApp.currentAccount().getName() + " | WeDiscover RSA Ad Group Checker";

  var executionContext = 'client_account';
  if (typeof AdsManagerApp != "undefined") {
    executionContext = 'manager_account';
  }

  if (executionContext === 'manager_account') {
    var accountIterator = AdsManagerApp.accounts().get();

    while(accountIterator.hasNext()) {
      AdsManagerApp.select(accountIterator.next());
      mainEmailBody += accountMain();
    }

  }

  else if (executionContext === 'client_account') {
    mainEmailBody += accountMain();
  }

  var fullEmailBody = emailOpener + mainEmailBody + emailFooter;

  MailApp.sendEmail(recipientEmail, emailSubject, "", {htmlBody: fullEmailBody});

}

// Function to be run on individual accounts
function accountMain() {
  var adGroupsAlreadyLabelled = [];
  var accountId = AdsApp.currentAccount().getCustomerId();
  var accountName = AdsApp.currentAccount().getName();

  addLabelIfNeeded(labelName)

  // Find ad groups without RSAs
  var allAdGroupIds = processQuery(allAdGroupsQuery);
  var hasRsaAdGroupIds = processQuery(rsaAdGroupsQuery);

//  var allAdGroupIds = Object.keys(allAdGroups);
//  var hasRsaIds = Object.keys(hasRsaAdGroups);

  var noRsaAdGroupIds = getElementsInFirstArrayOnly(allAdGroupIds, hasRsaAdGroupIds);
  var noRsaAdGroups = AdsApp.adGroups().withIds(noRsaAdGroupIds).get(); 

  // Update already labelled ad groups
  var labelledAdGroups = AdsApp.adGroups().withCondition("LabelNames CONTAINS_ANY ['" + labelName + "']").get()
  while (labelledAdGroups.hasNext()) {
    var adGroup = labelledAdGroups.next();
    var adGroupId = adGroup.getId();

    // If labelled ad group has an RSA, remove label and don't add it to our list for today
    if (hasRsaAdGroupIds.indexOf(adGroupId) !== -1) {
      adGroup.removeLabel(labelName);
      continue;
    }

    adGroupsAlreadyLabelled.push(adGroupId);
  }  

  // Get ad groups that require labels
  var idsToLabel = getElementsInFirstArrayOnly(noRsaAdGroupIds, adGroupsAlreadyLabelled);
  var adGroupsToLabel = AdsApp.adGroups().withIds(idsToLabel).get();

  // Apply labels if required, and group ad groups by campaign
  var labelledToday = iterateThroughAdGroups(adGroupsToLabel, true);
  var allLabelled = iterateThroughAdGroups(noRsaAdGroups, false);

  var emailBody = buildEmailBody(labelledToday, allLabelled, adGroupsToLabel.totalNumEntities(), noRsaAdGroups.totalNumEntities(), accountId, accountName);
  return emailBody;
}

// Returns object of form {ad_group_id:{'camapaignName': <campaign name>, 'adGroupName': <ad group name>}} for a given query
function processQuery(query) {
  var ids = [];
  var iterator = AdsApp.report(query).rows();

  while (iterator.hasNext()) {
    var row = iterator.next();
    ids.push(row['ad_group.id']);
  }

  return ids;
}

// Checks account to see if label exists, creates it if not
function addLabelIfNeeded(labelName) {
  var labelIterator = AdsApp.labels().withCondition("Name = '" + labelName + "'").get();
  var labelExists = Boolean(labelIterator.totalNumEntities());

  if (!labelExists) {
    AdsApp.createLabel(labelName);
  }
}

// Build body of email, listing campaigns & ad groups
function buildEmailBody(labelledToday, allLabelled, numLabelledToday, numAllLabelled, accountId, accountName) {
  var labelledTodayCampaigns = Object.keys(labelledToday).sort();
  var allLabelledCampaigns = Object.keys(allLabelled).sort();

  var accountIntro =
      "<b>Account: "+ accountName + " (CID: " + accountId + ")</b><br>"

  var body =
      "We have found " + numLabelledToday + " ad groups in this account which don't currently have an RSA in them. " + numAllLabelled + " of these are new today.<br>" +
      "<br>" +
      "New ad groups found per campaign:<br>";

  body += makeListForEmailBody(labelledTodayCampaigns, labelledToday);

  body += "Total ad groups per campaign:<br>";
  body += makeListForEmailBody(allLabelledCampaigns, allLabelled);

  return accountIntro + body;
}

function makeListForEmailBody(campaignList, entityInfo) {
  var listBody = "";
  for(var i = 0; i < campaignList.length; i++) {
    var campaignName = campaignList[i];
    var campaignId = entityInfo[campaignName]['campaignId'];
    var adGroups = entityInfo[campaignName]['adGroups'];

    listBody += "<ul>";
    listBody += "<li>" + campaignName + "   <i>(ID: " + campaignId + ")</i></li>";
    listBody += "<ul>"

    for(var j = 0; j < adGroups.length; j++) {
      var currentAdGroup = adGroups[j];
      listBody += "<li>" + currentAdGroup['adGroupName'] + "    <i>(ID: " + currentAdGroup['adGroupId'] + ")</i></li>";
    }

    listBody += "</ul>";
    listBody += "</ul>";
  }

  listBody += "<br>";

  return listBody;
}

// Iterate through ad groups, applying label if desired
// Return dictionary of ad group info grouped by campaign name, of the form:
// {<campaign name>: {'campaignId': <campaign id>, 'adGroups': <array of ['adGroupId': <ad group id>, 'adGroupName': <ad group name>]s}}
function iterateThroughAdGroups(adGroupIterator, applyLabel) {
  var groupedByCampaigns = {};

  while (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();  
    var adGroupId = adGroup.getId();
    var adGroupName = adGroup.getName();
    var campaignName = adGroup.getCampaign().getName();
    var campaignId = adGroup.getCampaign().getId();

    if (applyLabel) {
      adGroup.applyLabel(labelName);
    }

    if (groupedByCampaigns[campaignName]) {
      groupedByCampaigns[campaignName]['adGroups'].push({'adGroupId':adGroupId, 'adGroupName': adGroupName});
    }

    else if (!groupedByCampaigns[campaignName]) {
      groupedByCampaigns[campaignName] = {'campaignId': campaignId, 'adGroups':[{'adGroupId':adGroupId, 'adGroupName': adGroupName}]};
    }
  }
  return groupedByCampaigns;
}

// Return elements which are present in array1 but not in array2
function getElementsInFirstArrayOnly(array1, array2) {
  return array1.filter(function arrayFilter(element) {return array2.indexOf(element) === -1;});
}
