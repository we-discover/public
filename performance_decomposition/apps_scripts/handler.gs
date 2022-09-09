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

    this.workbook.toast("Running background process.", "Processing...", 2);
    Utilities.sleep(1000);
  }

  /**
   * Calculate the comparison period multiplier. 
   * 1 = Previous, 2 = lagged 1, 3 = lagged 2.
  */
  _getComparisonPeriodMultiplier(comparisonType) {
    let multiplier = 1;
    let re = new RegExp('Previous Period \\(-(.*)\\)');
    let matches = re.exec(comparisonType);
    if (matches !== null) {
      multiplier += parseInt(matches[1]);
    }
    return multiplier;
  }

  /**
   * Calculate the length of the period selected. Equal length periods 
   * applied for comparison periods, except for months.
  */
  _getPeriodLength(periodType) {
    if (periodType.includes('Week')) {
      return 7;
    }

    const today = new Date();
    if (periodType === 'This Month') {
      return new Date(
        today.getFullYear(),
        today.getMonth() + 1,
        0
      ).getDate();
    }
    if (periodType === 'Last Month') {
      return new Date(
        today.getFullYear(),
        today.getMonth() - 1,
        0
      ).getDate();
    }
    
    let re = new RegExp('Last \((.*)\) Days');
    let matches = re.exec(periodType);
    return parseInt(matches[1]);
  }

  /**
   * Set custom period type
  */
  setPeriodTypeCustom() {
    let periodType = this.sheet.getRange(controlRefs.periodType).getValue();
    if (periodType !== 'Custom') {
      this.sheet.getRange(controlRefs.periodType).setValue('Custom');
    }
  }

  /**
   * Set a default date range on initial and comparison period controls
  */
  setDefaultDateRange() {
      let periodType = this.sheet.getRange(controlRefs.periodType).getValue();
      let comparisonType = this.sheet.getRange(controlRefs.comparisonType).getValue();
    
      if (periodType === 'Custom') {
        return null
      }

      // Define the period length and comparison type period multiplier
      let periodLength = this._getPeriodLength(periodType);
      let comparisonMultiplier = this._getComparisonPeriodMultiplier(comparisonType);

      // Define the end of the comparison period
      let cEnd = new Date();
      if (periodType === 'Last Week') { // End of last week
        cEnd = cEnd.addDays(-cEnd.getDay());
      } else if (periodType === 'This Week') { // End of this week
        cEnd = cEnd.addDays(-cEnd.getDay() + 7);
      } else if (periodType === 'Last Month') { // End of last month
        cEnd = new Date(cEnd.getFullYear(), cEnd.getMonth(), 1).addDays(-1);
      } else if (periodType === 'This Month') { // End of this month
        cEnd = new Date(cEnd.getFullYear(), cEnd.getMonth(), periodLength);
      } else { // Yesterday
        cEnd = cEnd.addDays(-1);
      }

      let cStart = cEnd.addDays(-(periodLength - 1));
      let iEnd = cEnd.addDays(-comparisonMultiplier * periodLength);
      let iStart = cEnd.addDays(-(((comparisonMultiplier + 1) * periodLength)-1)); 

      // Handle month varied length exceptions
      if (periodType.includes('Month')) {
        iStart = iStart.addDays(1-iStart.getDate());
        cStart = cStart.addDays(1-cStart.getDate());

        let iMonthDays = new Date(
          iStart.getFullYear(),
          iStart.getMonth() + 1,
          0
        ).addDays(-1).getDate();
        iEnd = iStart.addDays(iMonthDays);
      } 

      // Set outputs in sheet
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
  */
  resetSheet(valuesOnly) {
    if (!valuesOnly) {
      // Clear control values
      for (const [key, ref] of Object.entries(controlRefs)) {
        // Never clear date controls as it's poor UX
        if (ref.includes('G') | ref.includes('I')) {
          continue; 
        }
        this.sheet.getRange(ref).clearContent();
      }
          // Clear performance metric headers
      for (let i = 0; i < displayRefs.performanceMetricHeaders.length; i++) {
        this.sheet.getRange(displayRefs.performanceMetricHeaders[i]).clearContent();
      }
    }

    // Clear decomp metric headers
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
   * Set relevant cell formatting for performance metric values
  */
  handlePerformanceMetricSelection(updatedRef) {
    let selectedMetric = this.sheet.getRange(updatedRef).getValue();
    let metricValueRefs = displayRefs.performanceMetricValues[
      displayRefs.performanceMetricHeaders.indexOf(updatedRef)
    ];
    let requiredFormat = metricFormats[selectedMetric];
    this.sheet.getRange(metricValueRefs).setNumberFormat(requiredFormat);
  }

  /**
   * Set all relevant values in the main sheet based on selected controls
  */
  updateMetricReferences() {
    let decompMetric = this.sheet.getRange(controlRefs.decompMetric).getValue();
    if (decompMetric === "") {
      throw new Error('Please choose a Decomposition Metric and try again.')
    }

    try {
      let matches = decompFormulas.filter(f => {return f.dependentVar.metric === decompMetric});
      let formula = matches[0];
    } catch(e) {
      throw new Error(decompMetric + ' is not supported.')
    }

    let decompMetricHeaderRange = this.sheet.getRange(displayRefs.decompMetricHeader);
    decompMetricHeaderRange.setValue(decompMetric);

    for (let i = 0; i < displayRefs.decompMetricValues.length; i++) {
      this.sheet
        .getRange(displayRefs.decompMetricValues[i])
        .setNumberFormat(metricFormats[decompMetric]);
    }

    for (let i = 0; i < formula.independentVars.length; i++) {
      this.sheet
        .getRange(displayRefs.independentMetricHeaders[i])
        .setValue(formula.independentVars[i].metric);

      let requiredFormat = metricFormats[formula.independentVars[i].metric];

      this.sheet
        .getRange(displayRefs.independentMetricValues[i])
        .setNumberFormat(requiredFormat);

      this.sheet
        .getRange(displayRefs.inverseIndicators[i])
        .setValue(formula.independentVars[i].inverse);
    }
  }

}
