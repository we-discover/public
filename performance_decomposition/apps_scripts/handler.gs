/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: A handler class for running core processes.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/


class DecompHandler {

  constructor() {
    this.workbook = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.workbook.getSheetByName(decompositionSheetName);
    
    this.decompMetric = this.sheet.getRange(controlRefs.decompMetric).getValue();
    if (this.decompMetric === "") {
      throw new Error('Please choose a Decomposition Metric and try again.')
    }

    try {
      let matches = decompFormulas.filter(f => {return f.dependentVar.metric === this.decompMetric});
      this.formula = matches[0]
    } catch(e) {
      throw new Error(this.decompMetric + ' is not supported.')
    }
    console.log('Handler instantiated for ' + this.decompMetric);
  }

  /**
   * Set a default date range on initial and comparison period controls
   * @return {null} No return value
  */
  _setDefaultDateRange() {
      // Set default date window (-7, -14 days)
      let today = new Date();
      
      const cEnd = today.addDays(-1);
      const cStart = today.addDays(-7);
      const iEnd = today.addDays(-8);
      const iStart = today.addDays(-14);

      this.sheet.getRange(controlRefs.comparisonPeriodEnd)
        .setValue(Utilities.formatDate(cEnd, "GMT", dateFormat));
      this.sheet.getRange(controlRefs.comparisonPeriodStart)
        .setValue(Utilities.formatDate(cStart, "GMT", dateFormat));
      this.sheet.getRange(controlRefs.initialPeriodEnd)
        .setValue(Utilities.formatDate(iEnd, "GMT", dateFormat));
      this.sheet.getRange(controlRefs.initialPeriodStart)
        .setValue(Utilities.formatDate(iStart, "GMT", dateFormat));    
  }

  /**
   * Reset controls and relevant calculations on the sheet.
   * @param {boolean} valuesOnly Whether to clear controls or just values.
   * @return {null} No return value
  */
  resetSheet(valuesOnly) {
    if (!valuesOnly) {
      // Clear control values
      for (const [key, ref] of Object.entries(controlRefs)) {
        this.sheet.getRange(ref).clearContent();
      }
      this._setDefaultDateRange();
    }

    // Clear performance metric headers
    for (let i = 0; i < displayRefs.independentMetricHeaders.length; i++) {
      this.sheet.getRange(displayRefs.independentMetricHeaders[i]).clearContent();
    }
    this.sheet.getRange(displayRefs.decompMetricHeader).clearContent();

    // Clear inverse indicators
    for (let i = 0; i < displayRefs.inverseIndicators.length; i++) {
      this.sheet.getRange(displayRefs.inverseIndicators[i]).clearContent();
    }    
  }

  /**
   * Set all relevant values in the main sheet based on selected controls
   * @return {null} No return value
  */
  updateMetricReferences() {
    let decompMetricHeaderRange = this.sheet.getRange(displayRefs.decompMetricHeader);
    decompMetricHeaderRange.setValue(this.decompMetric);

    for (let i = 0; i < displayRefs.decompMetricValues.length; i++) {
      this.sheet
        .getRange(displayRefs.decompMetricValues[i])
        .setNumberFormat(this.formula.dependentVar.format);
    }

    for (let i = 0; i < this.formula.independentVars.length; i++) {
      this.sheet
        .getRange(displayRefs.independentMetricHeaders[i])
        .setValue(this.formula.independentVars[i].metric);

      let requiredFormat = metricFormats[this.formula.independentVars[i].metric];

      this.sheet
        .getRange(displayRefs.independentMetricValues[i])
        .setNumberFormat(requiredFormat);

      this.sheet
        .getRange(displayRefs.inverseIndicators[i])
        .setValue(this.formula.independentVars[i].inverse);
    }
  }

}
