/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: Process to update the test summary table when appropriate.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

const summaryTableRefs = {
  'Variant': 'C12:C17',
  'CTR': 'F12:F17',
  'CVR': 'G12:G17',
  'CPI': 'H12:H17',
  'Cost': 'I12:I17',
  'Impressions': 'J12:J17',
  'Clicks': 'K12:K17',
  'Conversions': 'L12:L17',
  'Conversion Value': 'M12:M17'
};

const calcsCellRefs = {
  objectiveType: 'D3',
  objectiveMetric: 'E3',
  outcomes: 'AI3:AJ8'
};

const applicableControlChanges = [
  controlCellRefs.confidence,
  controlCellRefs.testName,
  controlCellRefs.testType
];


function updateSummaryTable(e) {
  Logger.log('Run updateSummaryTable');

  var workbook = SpreadsheetApp.getActiveSpreadsheet();

  if (applicableControlChanges.includes(e.range.getA1Notation())) {
    workbook.toast('📊 Computing test summary...', 'Status', 60);

    var summarySheet = workbook.getSheetByName(sheetNames.summary);
    var calcSheet = workbook.getSheetByName(sheetNames.calculations.summaryStatTests);

    const metrics = rateObjectiveMetrics.concat(volumeObjectiveMetrics);

    for (var i in metrics) {
      var metric = metrics[i];
      var metricType = 'Volume';
      if (rateObjectiveMetrics.includes(metric)) {
        metricType = 'Rate';
      }
      var tableOutputs = [];

      calcSheet.getRange(calcsCellRefs.objectiveType).setValue(metricType);
      calcSheet.getRange(calcsCellRefs.objectiveMetric).setValue(metric);

      var outcomes = calcSheet.getRange(calcsCellRefs.outcomes).getValues();
      for (var i in outcomes) {
        if (outcomes[i][0]) { // Is Best
          tableOutputs.push([true]);
          continue;
        }
        if (outcomes[i][1]) { // Is Worst
          tableOutputs.push([false]);
          continue;
        }
        tableOutputs.push(['']);
      }

      summarySheet.getRange(summaryTableRefs[metric]).setValues(tableOutputs);
    }

  }

  workbook.toast('Update complete! 🚀', 'Status', 3);
}
