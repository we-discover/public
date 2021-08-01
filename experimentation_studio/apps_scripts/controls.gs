/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: Handler to make changes to controls apply globally and
                 populate with context-specific defaults.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

function handleControlChanges(e) {

  Logger.log("Info: Running handleControlChanges");
  var workbook = SpreadsheetApp.getActiveSpreadsheet();
  workbook.toast('👋 Updating views...', 'Status', 60);
  var triggeringSheet = e.source.getActiveSheet();

  // Handle global control change
  if (isInControlRange(e.range.getA1Notation())) {
    var options = workbook.getSheetByName(sheetNames.options);
    var otherSheetsWithControls = getOtherSheetsWithControls(triggeringSheet.getName());

    // Update testName before propagating if testType is changed
    if (e.range.getA1Notation() === controlCellRefs.testType) {
      var testType = e.range.getValue();
      var firstTestNameRef = 'C2';
      if (testType === 'Experiment') {
        firstTestNameRef = 'B2';
      }
      var testRegistry = workbook.getSheetByName(sheetNames.registries[testType]);
      var firstTestName = testRegistry.getRange(firstTestNameRef).getValue();
      triggeringSheet.getRange(controlCellRefs.testName).setValue(firstTestName);
    }

    for (var i in otherSheetsWithControls) {
      var outputSheet = workbook.getSheetByName(otherSheetsWithControls[i]);
      for (const prop in controlCellRefs) {
        var ref = controlCellRefs[prop];
        var triggeringSheetValue = triggeringSheet.getRange(ref).getValue();
        outputSheet.getRange(ref).setValue(triggeringSheetValue);
      }
    }

    // Update Two Variant variants if testName or testType updated
    if ([controlCellRefs.testName, controlCellRefs.testType].includes(e.range.getA1Notation())) {
      var twoVarSheet = workbook.getSheetByName(sheetNames.drillDowns.twoVar)
      var firstVariantName = options.getRange('AD2').getValue();
      var secondVariantName = options.getRange('AD3').getValue();
      twoVarSheet.getRange(abVariantCellRefs[1]).setValue(firstVariantName);
      twoVarSheet.getRange(abVariantCellRefs[2]).setValue(secondVariantName);
    }
  }
}
