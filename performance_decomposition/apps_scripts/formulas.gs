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
    },
    independentVars: [
      {
        metric: 'Est. Searches',
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        inverse: false
      }
    ]
  },

  {
    dependentVar: {
      metric: 'Conversion Value',
    },
    independentVars: [
      {
        metric: 'Est. Searches', 
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        inverse: false
      }, 
      {
        metric: 'RPC', 
        inverse: false
      }
    ]
  },

  {
    dependentVar: {
      metric: 'CPA',
    },
    independentVars: [
      {
        metric: 'Cost', 
        inverse: false
      },    
      {
        metric: 'Est. Searches', 
        inverse: true
      }, 
      {
        metric: 'Impression Share', 
        inverse: true
      }, 
      {
        metric: 'CTR', 
        inverse: true
      }, 
      {
        metric: 'CVR', 
        inverse: true
      }

    ]
  },

  {
    dependentVar: {
      metric: 'ROAS',
    },
    independentVars: [
      {
        metric: 'Est. Searches', 
        inverse: false
      }, 
      {
        metric: 'Impression Share', 
        inverse: false
      }, 
      {
        metric: 'CTR', 
        inverse: false
      }, 
      {
        metric: 'CVR', 
        inverse: false
      },
      {
        metric: 'RPC', 
        inverse: false
      },
      {
        metric: 'Cost', 
        inverse: true
      },    
    ]
  }

]
