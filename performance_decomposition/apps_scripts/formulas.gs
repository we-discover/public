/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: Decomposition fomulas for supported metrics.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/

// All formulas are linear, where divisions are made by setting inverse=True.
// Formulas are calculated in the same order as they appear under independentVars.
const decompFormulas = [
  
  {
    dependentVar: {
      metric: 'Conversions',
      format: fmtValueInt
    },
    independentVars: [
      {
        metric: 'Est. Searches', 
        format: fmtValueInt, 
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        format: fmtPercentageInt, 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        format: fmtPercentageDec, 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        format: fmtPercentageDec, 
        inverse: false
      }
    ]
  },

  {
    dependentVar: {
      metric: 'Conversion Value',
      format: fmtCurrencyDec
    },
    independentVars: [
      {
        metric: 'Est. Searches', 
        format: fmtValueInt, 
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        format: fmtPercentageInt, 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        format: fmtPercentageDec, 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        format: fmtPercentageDec, 
        inverse: false
      }, 
      {
        metric: 'RPC', 
        format: fmtCurrencyDec, 
        inverse: false
      }
    ]
  },

  {
    dependentVar: {
      metric: 'CPA',
      format: fmtCurrencyDec
    },
    independentVars: [
      {
        metric: 'Cost', 
        format: fmtCurrencyInt, 
        inverse: false
      },    
      {
        metric: 'Est. Searches', 
        format: fmtValueInt, 
        inverse: true
      }, 
      {
        metric: 'Impression Share', 
        format: fmtPercentageInt, 
        inverse: true
      }, 
      {
        metric: 'CTR', 
        format: fmtPercentageDec, 
        inverse: true
      }, 
      {
        metric: 'CVR', 
        format: fmtPercentageDec, 
        inverse: true
      }

    ]
  },

  {
    dependentVar: {
      metric: 'ROAS',
      format: fmtPercentageDec
    },
    independentVars: [
      {
        metric: 'Est. Searches', 
        format: fmtValueInt, 
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        format: fmtPercentageInt, 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        format: fmtPercentageDec, 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        format: fmtPercentageDec, 
        inverse: false
      },
      {
        metric: 'RPC', 
        format: fmtCurrencyDec, 
        inverse: false
      },
      {
        metric: 'Cost', 
        format: fmtCurrencyInt, 
        inverse: true
      },    
    ]
  }

]
