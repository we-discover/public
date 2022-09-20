/*
    Name:        WeDiscover - Experimentation Studio, Google Ads Script
    Description: A script to export experimental observation data from Google Ads to a
                 templated Google Sheet in order to perform statistical evaluations of
                 D&E tests or those configured with labelled Campaigns, AdGroups or Ads.
    License:     https://github.com/we-discover/public/blob/master/LICENSE
    Version:     1.1.1
    Released:    2021-07-31
    Contact:     scripts@we-discover.com
*/


function main() {

  // EDIT ME -- Google Sheet ID for Template
  const gsheetId = 'XXXX';

  // Read all test configurations from GSheet
  var testConfigurations = loadTestConfigsFromSheet(gsheetId);

  // Determine runtime environment
  var executionContext = 'client_account';
  if (typeof AdsManagerApp != "undefined") {
    executionContext = 'manager_account';
  }

  // If MCC, run data collection process on a loop through all accounts
  if (executionContext === 'manager_account') {
    var managerAccount = AdsApp.currentAccount();
    var accountIterator = AdsManagerApp.accounts().get();
    while (accountIterator.hasNext()) {
      var account = accountIterator.next();
      AdsManagerApp.select(account);
      Logger.log('Info: Processing account ' + AdsApp.currentAccount().getName());
      testConfigurations = collectDataForTestConfigs(testConfigurations, gsheetId);
    }
    AdsManagerApp.select(managerAccount);
  }

  // If client account, run data collection process on that account only
  if (executionContext === 'client_account') {
    testConfigurations = collectDataForTestConfigs(testConfigurations, gsheetId);
  }

  // Loop through each test and export extracted data
  for (var i = 0; i < testConfigurations.length; i++) {
    try {
      if (testConfigurations[i].update && testConfigurations[i].processed) {
        exportDataToSheet(gsheetId, testConfigurations[i])
      }
    } catch (anyErrors) {
      Logger.log(anyErrors);
      continue;
    }
  }
}


// ========= UTILITY FUNCTIONS ==============================================================


// Basic assertion function
function assert(check, condition) {
  if (!condition) {
    throw new Error('Validation failed: ' + check);
  }
}

// Validates an imported test configuration
function validateConfiguration(config) {
  assert("supported test type", RegExp('^(label|experiment|prepost)$').test(config.config_type));

  if (RegExp('^(label|prepost)$').test(config.config_type)) {
    assert("supported label type", RegExp('^(ads|adGroups|campaigns)$').test(config.label_type));
  }

  var validDatePattern = RegExp('^[0-9]{8}$');

  if (RegExp('^(label|experiment)$').test(config.config_type)) {
    assert("start date formatted correctly", validDatePattern.test(config.start_date));
    assert("end date formatted correctly",  validDatePattern.test(config.end_date));
    assert("start date before end date", Number(config.start_date) < Number(config.end_date));
  }

  if (config.config_type === 'prepost') {
    var dateKeys = [
      'pre_start_date',
      'pre_end_date',
      'post_start_date',
      'pre_start_date'
    ];
    for (var i = 0; i < dateKeys.length; i++) {
      assert(dateKeys[i] + " formatted correctly", validDatePattern.test(config[dateKeys[i]]));
    }
    assert("pre start date before end date", Number(config.pre_start_date) < Number(config.pre_end_date));
    assert("post start date before end date", Number(config.post_start_date) < Number(config.post_end_date));
    assert("pre period before post period", Number(config.pre_end_date) < Number(config.post_start_date));
  }
}


// Extracts and format test configurations from sheet
function collectFromConfigSheet(type, sheet) {
    var configs = [];

    var [rows, columns] = [sheet.getLastRow(), sheet.getLastColumn()];
    var data = sheet.getRange(1, 1, rows, columns).getValues();
    const header = data[0];

    data.shift();
    data.map(function(row) {
      var empty = row[0] === '';
      if (!empty) {

        // Build config object from row
        var config = header.reduce(function(o, h, i) {
          o[h] = row[i];
          return o;
        }, {});
        config['data'] = {};
        config['processed'] = false;

        // Collect valid configs
        try {
          validateConfiguration(config);
          configs.push(config);
        } catch (validationErrors) {
          Logger.log('Error: Invalid config loaded from GSheet');
          Logger.log(validationErrors);
          Logger.log(config);
        }

      }
    });
    return configs
}


// Function to load test configurations from GSheet
function loadTestConfigsFromSheet(gsheetId) {
    const testTypes = ['label', 'experiment', 'prepost']

    var testConfigurations = [];

    try {
      var spreadsheet = SpreadsheetApp.openById(gsheetId);
      Logger.log('Info: Sucessfully connected to gsheet.');
    } catch (e) {
      throw Error('Failed to connect to gsheet.')
    }

    try {
      for (i=0; i < testTypes.length; i++) {
        var configSheetName = 'EXPORT - ' + testTypes[i] + ' configs';
        var configSheet = spreadsheet.getSheetByName(configSheetName);
        var typeConfigs = collectFromConfigSheet(testTypes[i], configSheet);
        testConfigurations = testConfigurations.concat(typeConfigs);
      }
    } catch (e) {
      throw Error('Failed to load test configurations from gsheet.')
    }

    return testConfigurations;
}


// Extracts variant_id from label (test_id:[...]$var_id:[...])
function extractVariantIdFromLabelName(labelName) {
  var matches = labelName.match('\\$var_id:([0-9]{1,3})');
  if (matches) {
    return matches[1];
  }
}


// Get the ID of entities for given test config
function getEntityIdsForTest(config) {

  var variantEntities = {};
  var labelIds = [];

  if (config.config_type === 'label') {
    var labelReport = AdsApp.report(
      "SELECT label.id " +
      "FROM label " +
      "WHERE label.name REGEXP_MATCH '" + config.mvt_label + "\\\\$.*'"
    ).rows();

    while(labelReport.hasNext()) {
      labelIds.push(Number(labelReport.next()["label.id"]));
    }

    var labelIterator = AdsApp.labels()
     .withIds(labelIds)
     .get();

    while (labelIterator.hasNext()) {
      var label = labelIterator.next();
      var variantId = extractVariantIdFromLabelName(label.getName());

      if (variantId === undefined) {
        continue;
      }

      // Evaluates like label.adGroups().get() (or similar)
      var entityIterator = eval('label.' + config.label_type + '().get()')

      if (entityIterator.totalNumEntities() < 1) {
        continue;
      }

      variantEntities[variantId] = [];
      while (entityIterator.hasNext()) {
        var entity = entityIterator.next();
        variantEntities[variantId].push(entity.getId());
      }
    }
  }

  if (config.config_type === 'prepost') {
    var labelIterator = AdsApp.labels()
      .withCondition("label.name = '" + config.mvt_label + "'")
      .get();

    if (labelIterator.totalNumEntities() === 1) {
      var label = labelIterator.next();

      variantEntities['pre'] = [];

      var entityIterator = eval('label.' + config.label_type + '().get()')

      while (entityIterator.hasNext()) {
        var entity = entityIterator.next();
        variantEntities['pre'].push(entity.getId());
      }
      variantEntities['post'] = variantEntities['pre'];
    }
  }

  if (config.config_type === 'experiment') {
    var experimentIterator = AdsApp.campaigns()
      .withCondition("label.name REGEXP_MATCH '(?i).*" + config.mvt_label + ".*'")
      .withCondition("campaign.experiment_type = EXPERIMENT")
      .get();

    if (experimentIterator.totalNumEntities() > 0) {
      variantEntities['control'] = [];
      variantEntities['variant'] = [];

      while (experimentIterator.hasNext()) {
        var experiment = experimentIterator.next();
        variantEntities['control'].push(experiment.getBaseCampaign().getId());
        variantEntities['variant'].push(experiment.getId());
      }
    }

  }

  return variantEntities;
}


// Get applicable report type based on test type
function getApplicableReportingValues(config) {

  // Defaults on all tests tests
  var reportType = 'campaign';
  var entityIdName = 'campaign.id';

  if (RegExp('^(label|prepost)$').test(config.config_type)) {

    if (config.label_type === 'adGroups') {
      reportType = 'ad_group';
      entityIdName = 'ad_group.id';
    }
    if (config.label_type === 'ads') {
      reportType = 'ad_group_ad';
      entityIdName = 'ad_group_ad.ad.id';
    }
  }
  return [reportType, entityIdName];
}


// Generate GAQL queries to pull data for each variant
function buildQueriesForVariants(config) {

  var [reportType, entityIdName] = getApplicableReportingValues(config);

  var variantIds = Object.keys(config.entities);
  var variantQueries = {};

  for (var i = 0; i < variantIds.length; i++) {
    var variantId = variantIds[i];
    var entityIds = config.entities[variantId];

    var dateCondition = config.start_date + " AND " + config.end_date;
    if (config.config_type === 'prepost' && variantId === 'pre') {
      dateCondition = config.pre_start_date + "," + config.pre_end_date;
    }
    if (config.config_type === 'prepost' && variantId === 'post') {
      dateCondition = config.post_start_date + " AND " + config.post_end_date;
    }

    variantQueries[variantId] = (" \
      SELECT \
          customer.descriptive_name \
        , segments.date \
        , metrics.cost_micros \
        , metrics.impressions \
        , metrics.clicks \
        , metrics.conversions \
        , metrics.conversions_value \
      FROM \
        " + reportType + " \
      WHERE \
        " + entityIdName + " IN (" + entityIds.join(',') + ") \
        AND metrics.impressions > 0 \
        AND segments.date BETWEEN " +
        " " + dateCondition
    ).replace(/ +(?= )/g, '')
  }

  return variantQueries;
}


// Runs GAQL query and aggregates data on a daily variant level
function queryAndAggregateData(gaqlQueries) {
  var dataObj = {};

  var variantIds = Object.keys(gaqlQueries);

  for (var i = 0; i < variantIds.length; i++) {
    var varId = variantIds[i];
    var resultIterator = AdsApp.report(gaqlQueries[varId]).rows();

    while (resultIterator.hasNext()) {
      var result = resultIterator.next();

      var date = result["Date"];

      if (!dataObj.hasOwnProperty(varId)) {
        dataObj[varId] = {};
      }

      if (!dataObj[varId].hasOwnProperty(date)) {
        dataObj[varId][date] = {
          'account_name': AdsApp.currentAccount().getName(),
          'currency': AdsApp.currentAccount().getCurrencyCode(),
          'cost': 0,
          'impressions': 0,
          'clicks': 0,
          'conversions': 0,
          'conversion_value': 0
        };
      }

      dataObj[varId][date]['cost'] += result["metrics.cost_micros"] / 1e6 || 0;
      dataObj[varId][date]['impressions'] += result["metrics.impressions"] || 0;
      dataObj[varId][date]['clicks'] += result["metrics.clicks"] || 0;
      dataObj[varId][date]['conversions'] += result["metrics.conversions"] || 0;
      dataObj[varId][date]['conversion_value'] += result["metrics.conversions_value"] || 0;
    }

  }
  return dataObj;
}


// Run on appropriate accounts to collect data for all relevant tests
function collectDataForTestConfigs(testConfigurations, gsheetId) {

  var processedConfigurations = [];

  for (var i = 0; i < testConfigurations.length; i++) {

    var config = testConfigurations[i];

    var accountTestMessage = (
      ' ' + config.config_type + ' test ' + config.name +
      ' in account ' + AdsApp.currentAccount().getName()
    );
    Logger.log('Info: Starting export for' + accountTestMessage);

    // Skip if not set to update or already processed
    if (!config.update || config.processed) {
      processedConfigurations.push(config);
      continue;
    }

     try {
      config.entities = getEntityIdsForTest(config);
      // Skip if no entities identified for test in current account
      if (Object.keys(config.entities).length === 0) {
        Logger.log('Info: No entities found for:' + accountTestMessage);
        processedConfigurations.push(config);
        continue;
      }
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Failed to identify entities for:' + accountTestMessage);
      continue;
    }

    try {
      var variantQueries = buildQueriesForVariants(config);
      var exportData = queryAndAggregateData(variantQueries);
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Failed to load data for:' + accountTestMessage);
      continue;
    }


  return processedConfigurations;
}


// Converts aggregated test data into array based GSheet rows
function formatTestDataForExport(config) {
  var output = [[
    'account_name',
    'currency',
    'test_name',
    'mvt_label',
    'variant_id',
    'date',
    'cost',
    'impressions',
    'clicks',
    'conversions',
    'conversion_value'
  ]];

  var data = config['data'];
  var accountName = AdsApp.currentAccount().getName();
  var currency = AdsApp.currentAccount().getCurrencyCode();

  for (var variantId in data) {
    for (var date in data[variantId]) {
      output.push([
        data[variantId][date]["account_name"],
        data[variantId][date]["currency"],
        config.name,
        config.mvt_label,
        variantId,
        date,
        data[variantId][date]["cost"],
        data[variantId][date]["impressions"],
        data[variantId][date]["clicks"],
        data[variantId][date]["conversions"],
        data[variantId][date]["conversion_value"]
      ]);
    }
  }

  return output;
}


// Connects to a Google Sheet and writes data for a single test
function exportDataToSheet(gsheetId, config) {

    var data = formatTestDataForExport(config);

    try {
      var spreadsheet = SpreadsheetApp.openById(gsheetId);
      Logger.log('Info: Sucessfully connected to sheet for test: ' + config.name);
    } catch (e) {
      throw Error('Connection to sheet failed for test: ' + config.name)
    }

    var importSheetName = "Data Import: " + config.name;
    var importSheet = spreadsheet.getSheetByName(importSheetName);
    if (importSheet === null) {
      importSheet = spreadsheet.insertSheet(importSheetName, 99);
    }
    importSheet.clear();
    Logger.log('Info: Sucessfully loaded data import sheet for test: ' + config.name);

    var importRange = importSheet.getRange(1, 1, data.length, data[0].length);
    importRange.setValues(data);
    importSheet.hideSheet();
    Logger.log('Info: Sucessfully exported test data for test: ' + config.name);
}
