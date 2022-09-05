/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: A set of common objects and functions for global usage.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/

const decompositionSheetName = 'Performance Decomposition';

const campaignStatsSheetName = 'Google Ads Import: Google Ads Campaign Stats';

const controlRefs = {
  account: 'C4',
  campaignRuleType: 'C5',
  campaignRule: 'D5',
  decompMetric: 'C6',
  conversionAction: 'C7', 
  periodType: 'G4',
  comparisonType: 'I4',
  initialPeriodStart: 'G5',
  initialPeriodEnd: 'I5',
  comparisonPeriodStart: 'G6',
  comparisonPeriodEnd: 'I6'  
}

const metricFormats = {
  'Cost': fmtCurrencyInt,
  'Est. Searches': fmtValueInt,
  'Impression Share': fmtPercentageInt,
  'Impressions': fmtValueInt,
  'CTR': fmtPercentageDec,
  'Clicks': fmtValueInt,
  'CVR': fmtPercentageDec,
  'Conversions': fmtValueInt,
  'RPC': fmtCurrencyDec,
  'Conversion Value': fmtCurrencyInt,
  'ROAS': fmtPercentageInt,
  'CPA': fmtCurrencyDec
};

const currencySymbol = getCurrencySymbol();
const fmtCurrencyInt = `${currencySymbol}#,##0`;
const fmtCurrencyDec = `${currencySymbol}#,##0.00`;

const fmtPercentageInt = '0%';
const fmtPercentageDec = '0.0%';
const fmtValueInt = '#,##0';
const fmtValueDec = '#,##0.00';


const displayRefs = {
  performanceMetricHeaders: [
    'C10',
    'D10',
    'E10',
    'F10',
    'G10',
    'H10',
    'I10'
  ],
  performanceMetricValues: [
    'C11:C12',
    'D11:D12',
    'E11:E12',
    'F11:F12',
    'G11:G12',
    'H11:H12',
    'I11:I12'
  ],
  decompMetricHeader: 'C16',
  decompMetricValues: [
    'C17:C18',
    'D17:I18',
    'M3:M10'
  ],
  independentMetricHeaders: [
    'D16',
    'E16',
    'F16',
    'G16',
    'H16',
    'I16'
  ],
  independentMetricValues: [
    'D17:D18',
    'E17:E18',
    'F17:F18',
    'G17:G18',
    'H17:H18',
    'I17:I18'
  ],  
  inverseIndicators: [
    'D25',
    'E25',
    'F25',
    'G25',
    'H25',
    'I25'
  ] 
}

const dateFormat = 'yyyy-MM-dd';

// Utility method for adding days to a date
Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

/**
 * Get an array of row values for a given column in a given sheet
 * @param {SpreadsheetApp.Sheet} sheet object of sheet to search in
 * @param {String} column column to search in, in A1 notation - i.e. A, B, C, AA, AB, etc
 * @return {Array} Flat array of row values
 */
function getRowValues(sheet, column) {
  return sheet.getRange(`${column}:${column}`)
    .getValues()
    .filter(
      (row) => row.length > 0
      )
    .map(
      (row) => row[0]
    );
}


/**
 * Get the currency symbol for the account being decomposed
 * @return {String} Currency symbol for account
*/
function getCurrencySymbol() {
  // Currency symbol lookup
  const currencySymbols = {
    'AED': 'د.إ',
    'AFN': '؋',
    'ALL': 'Lek',
    'ANG': 'ƒ',
    'ARS': '$',
    'AUD': '$',
    'AWG': 'ƒ',
    'AZN': '₼',
    'BAM': 'KM',
    'BBD': '$',
    'BGN': 'лв',
    'BMD': '$',
    'BND': '$',
    'BOB': '$b',
    'BRL': 'R$',
    'BSD': '$',
    'BWP': 'P',
    'BYN': 'Br',
    'BZD': 'BZ$',
    'CAD': '$',
    'CHF': 'CHF',
    'CLP': '$',
    'CNY': '¥',
    'COP': '$',
    'CRC': '₡',
    'CUP': '₱',
    'CZK': 'Kč',
    'DKK': 'kr',
    'DOP': 'RD$',
    'EGP': '£',
    'EUR': '€',
    'FJD': '$',
    'FKP': '£',
    'GBP': '£',
    'GGP': '£',
    'GHS': '¢',
    'GIP': '£',
    'GTQ': 'Q',
    'GYD': '$',
    'HKD': '$',
    'HNL': 'L',
    'HRK': 'kn',
    'HUF': 'Ft',
    'IDR': 'Rp',
    'ILS': '₪',
    'IMP': '£',
    'INR': '₹',
    'IRR': '﷼',
    'ISK': 'kr',
    'JEP': '£',
    'JMD': 'J$',
    'JPY': '¥',
    'KGS': 'лв',
    'KHR': '៛',
    'KPW': '₩',
    'KRW': '₩',
    'KRW': '₩',
    'KYD': '$',
    'KZT': 'лв',
    'LAK': '₭',
    'LBP': '£',
    'LKR': '₨',
    'LRD': '$',
    'MKD': 'ден',
    'MNT': '₮',
    'MNT': 'د.إ',
    'MUR': '₨',
    'MXN': '$',
    'MYR': 'RM',
    'MZN': 'MT',
    'NAD': '$',
    'NGN': '₦',
    'NIO': 'C$',
    'NOK': 'kr',
    'NPR': '₨',
    'NZD': '$',
    'OMR': '﷼',
    'PAB': 'B/.',
    'PEN': 'S/.',
    'PHP': '₱',
    'PKR': '₨',
    'PLN': 'zł',
    'PYG': 'Gs',
    'QAR': '﷼',
    'RON': 'lei',
    'RSD': 'Дин.',
    'RUB': '₽',
    'SAR': '﷼',
    'SBD': '$',
    'SCR': '₨',
    'SEK': 'kr',
    'SGD': '$',
    'SHP': '£',
    'SOS': 'S',
    'SRD': '$',
    'SVC': '$',
    'SYP': '£',
    'THB': '฿',
    'TRY': '₺',
    'TTD': 'TT$',
    'TVD': '$',
    'TWD': 'NT$',
    'UAH': '₴',
    'USD': '$',
    'UYU': '$U',
    'UZS': 'лв',
    'VEF': 'Bs',
    'VND': '₫',
    'XCD': '$',
    'YER': '﷼',
    'ZAR': 'R',
    'ZWD': 'Z$',
  };

  const campaignStatsColumns = {
    accountName: "C",
    currencyCode: "D",
  }

  const spreadsheet = SpreadsheetApp.getActive();
  const accountToDecompose = spreadsheet.getSheetByName(decompositionSheetName)
      .getRange(controlRefs.account)
      .getValue();
  const campaignStatsSheet = spreadsheet.getSheetByName(campaignStatsSheetName);
  const accountsInSheet = getRowValues(campaignStatsSheet, campaignStatsColumns.accountName);
  const currenciesInSheet = getRowValues(campaignStatsSheet, campaignStatsColumns.currencyCode);
  const accountCurrencyCode = currenciesInSheet[accountsInSheet.indexOf(accountToDecompose)];
  return currencySymbols[accountCurrencyCode];
}
