/*
    Name:        WeDiscover - RSA Checker, Google Ads Script
    Description: A script to run an email based alert where ad groups do not have
                 an RSA present.
    License:     https://github.com/we-discover/public/blob/master/LICENSE
    Version:     1.0.0
    Released:    2021-09-13
    Contact:     scripts@we-discover.com
*/


// EDIT ME -- Check for RSAs in Paused Ads, Ad Groups or Campaigns
var checkPausedCampaigns = false;
var checkPausedAdGroups = false;
var checkPausedAds = false;

// EDIT ME -- Whether or not to send an email only when new groups are found without RSAs
var alertOnNewEntitiesOnly = false;

// EDIT ME -- Set email addresses to receive the alert, separated by commas
var recipientEmails = "example@gmail.com,example.2@gmail.com";

// EDIT ME -- Label name for controlling the 'new' detection behaviour
var labelName = "no_rsa_present";


// ========= CORE FUNCTIONS ============================================================================================


// Script entrypoint
function main() {
    var executionContext = getExecutionContext();
    var topLevelAccount = AdsApp.currentAccount();
    var totalGroupsWithoutAnRsa = 0;
    var accountCheckSummaries = [];

    // If MCC, run data collection process on a loop through all accounts
    if (executionContext === 'manager_account') {
        var accountIterator = AdsManagerApp.accounts().get();
        while(accountIterator.hasNext()) {
            AdsManagerApp.select(accountIterator.next());
            accountCheck = checkEntitiesForRSAs();
            totalGroupsWithoutAnRsa += accountCheck['nAdGroupIdsWithoutRsa'];
            accountCheckSummaries.push(accountCheck);
        }
    }

    // If client account, run data collection process on that account only
    if (executionContext === 'client_account') {
        accountCheck = checkEntitiesForRSAs();
        totalGroupsWithoutAnRsa += accountCheck['nAdGroupIdsWithoutRsa'];
        accountCheckSummaries.push(accountCheck);
    }

    if (totalGroupsWithoutAnRsa >= 1 && !alertOnNewEntitiesOnly) {
        sendEmail = true
    };

    sendSummaryEmail(topLevelAccount, accountCheckSummaries);
}


// Process to extract entities without RSAs from a given account
function checkEntitiesForRSAs() {

    var allAdGroupIds = getAdGroupsWithCondition('all_ad_groups');
    var adGroupIdsWithRsa = getAdGroupsWithCondition('ad_groups_with_rsas');
    var adGroupIdsWithoutRsa = allAdGroupIds.filter(function(id) {
        return adGroupIdsWithRsa.indexOf(id) === -1
    });

    addLabelToAccountIfNeeded(labelName);

    // Remove label from any Ad Groups with an RSA
    handleGroupsWithRsa(adGroupIdsWithRsa);

    // Add labels to any Ad Groups without an RSA and collect details
    var alertableEntities = handleGroupsWithoutRsa(adGroupIdsWithoutRsa);

    return {
        'accountId': AdsApp.currentAccount().getCustomerId(),
        'accountName': AdsApp.currentAccount().getName(),
        'nAdGroupsWithRsa': adGroupIdsWithRsa.length,
        'nAdGroupIdsWithoutRsa': adGroupIdsWithoutRsa.length,
        'alertableEntities': alertableEntities
    }
}


// ========= UTILITY FUNCTIONS =========================================================================================


// Determine the type of account in which the script is running
function getExecutionContext() {
    if (typeof AdsManagerApp != "undefined") {
        return 'manager_account';
    }
    return 'client_account';
}

// Returns a list of Ad Group IDs that correspond to a given condition
function getAdGroupsWithCondition(condition) {
    var query = queries[condition];
    var ids = [];

    var iterator = AdsApp.report(query).rows();
    while (iterator.hasNext()) {
        var row = iterator.next();
        ids.push(row['ad_group.id']);
    }

    return ids;
}

// Checks account to see if label exists, creates it if not
function addLabelToAccountIfNeeded(labelName) {
    var labelIterator = AdsApp.labels()
        .withCondition("Name = '" + labelName + "'")
        .get();

    var labelExists = labelIterator.totalNumEntities() >= 1;

    if (!labelExists) {
        AdsApp.createLabel(labelName);
    }
}

// Remove label from any Ad Groups with an RSA
function handleGroupsWithRsa(ids) {
    var adGroupsWithRsa = AdsApp.adGroups()
        .withIds(ids)
        .get();

    while (adGroupsWithRsa.hasNext()) {
        var adGroup = adGroupsWithRsa.next();
        adGroup.removeLabel(labelName);
    }
}

// Add a campaign/group record to output object, handling existence
function updateOutputObjectWithAdGroup(outputObj, type, campaign, adGroup) {
    if (!outputObj[type].hasOwnProperty(campaign.getId())) {
        outputObj[type][campaign.getId()] = {
            'campaignId': campaign.getId(),
            'campaignName': campaign.getName(),
            'adGroups':[]
        };
    }
    outputObj[type][campaign.getId()]['adGroups'].push({
        'adGroupId': adGroup.getId(),
        'adGroupName': adGroup.getName()
    });
    return outputObj;
}

// Add labels to any Ad Groups without an RSA and collect details
function handleGroupsWithoutRsa(ids) {

    var alertableEntities = {'new': {}, 'existing': {}};

    var adGroupsWithoutRsa = AdsApp.adGroups()
        .withIds(ids)
        .get();

    while (adGroupsWithoutRsa.hasNext()) {
        var adGroup = adGroupsWithoutRsa.next();
        var campaign = adGroup.getCampaign();

        var isAlreadyLabelled = adGroup.labels()
            .withCondition("LabelName = '" + labelName + "'")
            .get().totalNumEntities() >= 1;

        var type = 'existing';
        if (!isAlreadyLabelled) {
            type = 'new';
            adGroup.applyLabel(labelName);
            sendEmail = true;
        }

        alertableEntities = updateOutputObjectWithAdGroup(
            alertableEntities, type, campaign, adGroup
        );
    }
    return alertableEntities;
}


// ========= GAQL QUERIES ==============================================================================================

var today = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd");

var queryPart_entityConstraints = (" \
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
        " + queryPart_entityConstraints
  ).replace(/ +(?= )/g, '');

var queryAdGroupsWithRsa = (" \
    SELECT \
        ad_group.id \
    FROM \
        ad_group_ad \
    WHERE \
        " + queryPart_entityConstraints + " \
        AND ad_group_ad.status IN ('ENABLED'" + (checkPausedAds ? ", PAUSED'" : "") + ") \
        AND ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD'"
  ).replace(/ +(?= )/g, '');

var queries = {
    'all_ad_groups': queryAllAdGroups,
    'ad_groups_with_rsas': queryAdGroupsWithRsa
}


// ========= EMAIL COPY & FUNCTIONS ====================================================================================


var sendEmail = false;

var emailIntroduction =
    "Hi there,<br><br>" +
    "This is your automated email from the WeDiscover RSA Ad Group Checker.<br><br>" +
    "Included below are the details of ad groups without RSAs in your account(s). " +
    "The label '<i>" + labelName + "</i>' has been applied to all ad groups listed.<br><br>";

var emailFooter =
    "All the best,<br>" +
    "WeDiscover<br>" +
    "<br>" +
    "*If you have any questions about this script, please email " +
    "<a href = \"mailto:scripts@we-discover.com\">scripts@we-discover.com</a>";


function sendSummaryEmail(topLevelAccount, accountChecks) {
    var subject = topLevelAccount.getName() + " | WeDiscover RSA Ad Group Checker";

    accountSections = "";
    for (var i in accountChecks) {
        accountSections += buildAccountEmailSection(accountChecks[i]);
    }

    var body = emailIntroduction + accountSections + emailFooter;

    if (sendEmail) {
        MailApp.sendEmail(recipientEmails, subject, "", {htmlBody: body});
    }
}

// Build body of an email for an account check listing alertable campaigns & ad groups
function buildAccountEmailSection(check) {
    var body = "<b>Account: " + check['accountName'] + " (CID: " + check['accountId'] + ")</b><br>";
    body += "There are " + check['nAdGroupIdsWithoutRsa'] + " ad groups in this account without an RSA";

    if (Object.keys(check['alertableEntities']['new']).length >= 1) {
        var sectionData = makeListForEmailBody(check['alertableEntities']['new']);
        body += ", and " + sectionData['nGroups'] + " of these have been detected for the first time.<br><br>";
        body += "Newly detected Ad Groups: " + sectionData['nGroups'] + "<br>";
        body += sectionData['body'];
    } else {
        body += ", and none of these have been detected for the first time.<br><br>";
    }
    if (Object.keys(check['alertableEntities']['existing']).length >= 1) {
        var sectionData = makeListForEmailBody(check['alertableEntities']['existing']);
        body += "Previously detected Ad Groups: " + sectionData['nGroups'] + "<br>";
        body += sectionData['body'];
    }

    return body + "<br>";
}

// Create a list with campaign/group hierachy for email body
function makeListForEmailBody(entities) {
    var body = "";
    var nGroups = 0;
    for (var i = 0; i < Object.keys(entities).length; i++) {
        entity = entities[Object.keys(entities)[i]];

        body += "<ul>";
        body += "<li>" + entity['campaignName'] + " <i>(ID: " + entity['campaignId'] + ")</i></li>";
        body += "<ul>"
        for (var j = 0; j < entity['adGroups'].length; j++) {
            var group = entity['adGroups'][j];
            body += "<li>" + group['adGroupName'] + " <i>(ID: " + group['adGroupId'] + ")</i></li>";
            nGroups += 1;
        }
        body += "</ul>";
        body += "</ul>";
    }
    return {
        'body': body + "<br>",
        'nGroups': nGroups
    }
}
