/*
    Name:        WeDiscover - Performance Decomposition, Google Apps Script

    Description: Listeners and entrypoints from the gsheet

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2022-08-01

    Contact:     scripts@we-discover.com
*/


function onEdit(e) {
  // Reset calculations on metric change
  if (e.range.getA1Notation() === controlRefs.decompMetric) {
    var handler = new DecompHandler();
    handler.resetSheet(true);
  }
  if (e.range.getA1Notation() === controlRefs.campaignRuleType) {
    var handler = new DecompHandler();
    if (handler.sheet.getRange(controlRefs.campaignRuleType).getValue() === 'Show All') {
      handler.sheet.getRange(controlRefs.campaignRule).setValue('');
    }
  }
}


function handleRunCommand() {
  var handler = new DecompHandler();
  handler.resetSheet(true);
  handler.updateMetricReferences();
} 


function handleResetCommand() {
  var handler = new DecompHandler();
  handler.resetSheet();
} 
