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
  outcomes: 'AI3:AJ8',
  values: {
    'Rate': 'AQ15:AV15',
    'Volume': 'AI13:AN13'
  }
};

const powerCellRefs = {
  bestHasPower: 'C19',
  worstHasPower: 'G19'
};

const applicableControlChanges = [
  controlCellRefs.confidence,
  controlCellRefs.power,
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
    var powerSheet = workbook.getSheetByName(sheetNames.calculations.summaryPowerTests);

    const metrics = rateObjectiveMetrics.concat(volumeObjectiveMetrics);

    for (var i in metrics) {
      var metric = metrics[i];

      var metricType = 'Volume';
      if (rateObjectiveMetrics.includes(metric)) {
        metricType = 'Rate';
      }

      calcSheet.getRange(calcsCellRefs.objectiveType).setValue(metricType);
      calcSheet.getRange(calcsCellRefs.objectiveMetric).setValue(metric);

      var tableOutputs = {values: [], colours: []};
      var outcomes = calcSheet.getRange(calcsCellRefs.outcomes).getValues();
      var values = calcSheet.getRange(calcsCellRefs.values[metricType]).getValues();

      for (var i in outcomes) {
        tableOutputs.values.push([values[0][i]]);

        if (outcomes[i][0]) { // Is Best
          var bestHasPower = powerSheet.getRange(powerCellRefs.bestHasPower).getValue();
          if (bestHasPower) {
            tableOutputs.colours.push([colours.lightGreen]);
            continue;
          }
        }

        if (outcomes[i][1]) { // Is Worst
          var worstHasPower = powerSheet.getRange(powerCellRefs.worstHasPower).getValue();
          if (worstHasPower) {
            tableOutputs.colours.push([colours.lightRed]);
            continue;
          }
        }

        tableOutputs.colours.push(['']);
      }

      var outputRange = summarySheet.getRange(summaryTableRefs[metric]);
      outputRange.setValues(tableOutputs.values);
      outputRange.setBackgrounds(tableOutputs.colours);
    }

  }
  workbook.toast('Update complete! 🚀', 'Status', 3);
}
