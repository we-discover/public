/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: A set of common objects and functions for global usage.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

const colours = {
  'white': '#FFFFFF',
  'grey': '#d9d9d9',
  'lightGrey': '#f3f3f3',
  'navy': '#0e2244',
  'lightRed': '#f4c7c3',
  'lightGreen': '#b7e1cd',
  'variants': [
    '#0e2244',
    '#93c47d',
    '#6d9eeb',
    '#38761d',
    '#f6b26b',
    '#674ea7'
  ]
};

const defaultDateFormat = "yyyy-MM-dd";

const volumeObjectiveMetrics = [
  'Cost',
  'Impressions',
  'Clicks',
  'Conversions',
  'Conversion Value',
];

const rateObjectiveMetrics = [
  'CTR',
  'CVR',
  'CPI'
];

const sheetNames = {
  'form': 'Test Configurator',
  'options': 'Options',
  'registries': {
    'Label': 'Registry: Labels',
    'Experiment': 'Registry: Experiments',
    'Pre Post': 'Registry: Pre/Post'
  },
  'summary': 'Overview - Test Summary',
  'drillDowns': {
    'mvt': 'Drill Down - All Variants',
    'twoVar': 'Drill Down - Two Variants'
  },
  'metricData': {
    'rate': {
      'mvt': 'Intermediate - Rate Obj Vis',
      'twoVar': 'Intermediate - Rate Obj Vis (Two Variants)'
    },
    'volume': {
      'mvt': 'Intermediate - Volume Obj Vis',
      'twoVar': 'Intermediate - Volume Obj Vis (Two Variants)'
    },
    'visMetric': {
      'mvt': 'Intermediate - Visualisation Metric',
      'twoVar': 'Intermediate - Visualisation Metric (Two Variants)'
    }
  },
  'calculations': {
    'statTests': 'Intermediate - Significance Tests',
    'summaryStatTests': 'Intermediate - Significance Tests (Summary)',
    'power': 'Intermediate - Power Analysis',
    'summaryPowerTests': 'Intermediate - Power Analysis (Summary)'
  }
};

const sheetsWithControls = Object
  .values(sheetNames.drillDowns)
  .concat(sheetNames.summary);

const controlCellRefs = {
  testType: 'E4',
  testName: 'E5',
  objectiveMetric: 'E6',
  confidence: 'E7',
  power: 'E8'
};

const abVariantCellRefs = {
  1: 'K4',
  2: 'K5'
};

// Utility function to check if A1 reference is in control range
function isInControlRange(a1Reference) {
  return (
    Object.values(controlCellRefs).includes(a1Reference)
  )
}

// Utility function to get other sheets with controls
function getOtherSheetsWithControls(currentSheetName) {
  return sheetsWithControls.filter(function(value, index, arr){
    return value != currentSheetName;
  })
}

// Utility functon to return other drill down sheet
function getOtherDrillDownSheet(currentDrillDownSheetName) {
  // With two drill downs, output sheet always one that didn't trigger
  var drillDownSheetNames = Object.values(sheetNames.drillDowns);
  return drillDownSheetNames.filter(function(value, index, arr){
    return value != currentDrillDownSheetName;
  })[0];
}

// Utility function for finding first empty row
function getFirstEmptyRow(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues();
  var ct = 0;
  while (values[ct] && values[ct][0] != "") {
    ct++;
  }
  return ct + 1;
}

// Method to add days to date obj
Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}
