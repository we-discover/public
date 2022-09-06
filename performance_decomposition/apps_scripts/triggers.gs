/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: Listeners and entrypoints from the gsheet

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/


function onEdit(e) {

  var changesMade = false;

  // Reset calculations on metric change or account selection
  if ([controlRefs.decompMetric, controlRefs.account].includes(e.range.getA1Notation())) {
    var handler = new DecompHandler();
    handler.resetSheet(true);
    changesMade = true;
  }

  // Nullify campaign filter criterion if show all
  if (e.range.getA1Notation() === controlRefs.campaignRuleType) {
    var handler = new DecompHandler();
    if (handler.sheet.getRange(controlRefs.campaignRuleType).getValue() === 'Show all') {
      handler.sheet.getRange(controlRefs.campaignRule).setValue('');
    }
    changesMade = true;
  }

  // Handle periodType and comparisonType date changes
  if ([controlRefs.periodType, controlRefs.comparisonType].includes(e.range.getA1Notation())) {
    var handler = new DecompHandler();
    handler.setDefaultDateRange();
    changesMade = true;
  }

  // Handle date range changes
  var dateRangeRefs = [
    controlRefs.initialPeriodEnd,
    controlRefs.initialPeriodStart,
    controlRefs.comparisonPeriodEnd,
    controlRefs.comparisonPeriodStart
  ]
  if (dateRangeRefs.includes(e.range.getA1Notation())) {
    var handler = new DecompHandler();
    handler.setPeriodTypeCustom();
    changesMade = true;
  }  

  // Handle performance metric selection
  if (displayRefs.performanceMetricHeaders.includes(e.range.getA1Notation())) {
    if (e.range.getValue() !== '') {
      var handler = new DecompHandler();
      handler.handlePerformanceMetricSelection(e.range.getA1Notation());
      changesMade = true;
    }
  }

  if (changesMade) {
    handler.workbook.toast("Background processes completed.", "Finished" );
  }
}


function handleRunCommand() {
  var handler = new DecompHandler();
  handler.resetSheet(true);
  handler.updateMetricReferences();
  handler.workbook.toast("Background processes completed.", "Finished" );
} 


function handleResetCommand() {
  var handler = new DecompHandler(true);
  handler.resetSheet();
  handler.workbook.toast("Background processes completed.", "Finished" );
} 
