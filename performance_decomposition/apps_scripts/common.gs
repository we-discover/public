/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: A set of common objects and functions for global usage.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/


const decompositionSheetName = 'Performance Decomposition';

const controlRefs = {
  account: 'C4',
  campaignRuleType: 'C5',
  campaignRule: 'D5',
  decompMetric: 'C6',
  conversionAction: 'C7',
  initialPeriodStart: 'G4',
  initialPeriodEnd: 'I4',
  comparisonPeriodStart: 'G5',
  comparisonPeriodEnd: 'I5'  
}

// Todo: Make currency dynamic
const fmtCurrencyInt = '£#,##0';
const fmtCurrencyDec = '£#,##0.00';
const fmtPercentageInt = '0%';
const fmtPercentageDec = '0.0%';
const fmtValueInt = '#,##0';
const fmtValueDec = '#,##0.00';


const displayRefs = {
  decompMetricHeader: 'C10',
  decompMetricValues: [
    'C11:C12',
    'D17:I18',
    'M4:M11'
  ],
  independentMetricHeaders: [
    'D10',
    'E10',
    'F10',
    'G10',
    'H10',
    'I10'
  ],
  independentMetricValues: [
    'D11:D12',
    'E11:E12',
    'F11:F12',
    'G11:G12',
    'H11:H12',
    'I11:I12'
  ],  
  inverseIndicators: [
    'D19',
    'E19',
    'F19',
    'G19',
    'H19',
    'I19'
  ] 
}

const dateFormat = 'yyyy-MM-dd'

// Utility method for adding days to a date
Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}
