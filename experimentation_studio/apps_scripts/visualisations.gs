/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: Responsible for dynamically updating visualisations and cell
                 formatting based on the type of test or specific metrics
                 chosen.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

// Callback for cell formats
function toggleCellFormats(e) {
  var workbook = SpreadsheetApp.getActiveSpreadsheet();

  var optionsSheet = workbook.getSheetByName(sheetNames.options);
  var currencySymbol = optionsSheet.getRange('AF2').getValue();
  var currencyFormat = currencySymbol + '#,##0';

  // Control global sheet cell currency formats for test selection
  if ([controlCellRefs.testName, controlCellRefs.testType].includes(e.range.getA1Notation())) {
    Logger.log('Run toggleCellFormats for fixed');

    for (var i in fixedCurrencyRanges) {
      var targetSheetName = getProp(sheetNames, fixedCurrencyRanges[i].sheetNameRef);
      var targetSheet = workbook.getSheetByName(targetSheetName);

      for (var j in fixedCurrencyRanges[i].ranges) {
        var currencyRange = fixedCurrencyRanges[i].ranges[j];
        targetSheet.getRange(currencyRange).setNumberFormat(currencyFormat);
      }
    }
  }

  // Control cell formats for toggle on drill down views
  if (isInControlRange(e.range.getA1Notation())) {
    Logger.log('Run toggleCellFormats for objective');

    for (var i in sheetNames.drillDowns) {
      var outputSheet = workbook.getSheetByName(sheetNames.drillDowns[i]);
      var selectedObjectiveMetric = outputSheet
        .getRange(controlCellRefs.objectiveMetric)
        .getValue();

      if (['Cost', 'Conversion Value'].includes(selectedObjectiveMetric)) {
        outputSheet.getRange("E28:J33").setNumberFormat(currencyFormat);
        continue;
      }
      if (rateObjectiveMetrics.includes(selectedObjectiveMetric)) {
        outputSheet.getRange("F28:J33").setNumberFormat('0.0%');
        continue;
      }

      outputSheet.getRange("E28:J33").setNumberFormat('#,##0');
    }
  }
}


// Callback to control showing variants
function toggleVisualisationRecords(e) {
  Logger.log('Run toggleVisualisationRecords');
  var handler = new MetricVisualisationHandler(e);
  handler.showRelevantVariantsOnVisMetricChart();
}


// Callback for visualisation toggle on drill down views
function toggleCoreVisualisations(e) {
  Logger.log('Run toggleCoreVisualisations');
  var handler = new MetricVisualisationHandler(e);
  handler.clearObjectiveMetricCharts();
  handler.prepareDisplayArea();
  handler.addSelectedObjectiveMetricCharts();
}


class MetricVisualisationHandler {

  constructor(event) {
    this.chartProps = new ChartPropertyFactory();
    this.workbook = SpreadsheetApp.getActiveSpreadsheet();

    this.triggeringSheet = event.source.getActiveSheet();
    var triggeringSheetName = this.triggeringSheet.getName();
    if (triggeringSheetName === sheetNames.summary) {
      triggeringSheetName = sheetNames.drillDowns.mvt;
      this.triggeringSheet = this.workbook.getSheetByName(triggeringSheetName);
    }

    this.sheetTypes = ['mvt', 'twoVar'];
    if (triggeringSheetName === sheetNames.drillDowns.twoVar) {
      this.sheetTypes = ['twoVar', 'mvt'];
    }

    this.testType = this.triggeringSheet.getRange('E4').getValue();
    this.testName = this.triggeringSheet.getRange('E5').getValue();
    this.metric = this.triggeringSheet.getRange('E6').getValue();
    this.numVariants = this._getNumberOfVariants();

    this.outputSheets = {
      'mvt': this.workbook.getSheetByName(sheetNames.drillDowns.mvt),
      'twoVar': this.workbook.getSheetByName(sheetNames.drillDowns.twoVar)
    }

    this.objType = 'rate';
    if (volumeObjectiveMetrics.includes(this.metric)) {
      this.objType = 'volume';
    }
  }

  clearObjectiveMetricCharts() {
    for (var i in this.sheetTypes) {
      var targetSheet = this.outputSheets[this.sheetTypes[i]];
      var charts = targetSheet.getCharts();
      for (var i in charts) {
        // Rough method for determining if chart is one to replace
        if (charts[i].getOptions().get('title').includes('Distribution')) {
          targetSheet.removeChart(charts[i]);
        };
      }
    }
  }

  addSelectedObjectiveMetricCharts() {
    if (this.objType === 'volume') {
      this._addCombinedVolumeDistributionChart();
      this._addVariantVolumeDistributionCharts();
    }
    if (this.objType === 'rate') {
      this._addRateTimeseriesChart();
    }
  }

  prepareDisplayArea() {
    if (this.numVariants < 4 || this.objType === 'rate') {
      this.outputSheets.mvt.hideRows(54, 4);
    } else {
      this.outputSheets.mvt.showRows(54, 4);
    }
  }

  _getNumberOfVariants() {
    if (this.testType != 'Label') {
      return 2;
    }
    var numVariants = 0;
    var registrySheet = this.workbook.getSheetByName(sheetNames.registries[this.testType]);
    var labelVariantsData = registrySheet.getRange('C:D').getValues();
    for (var i = 0; i < labelVariantsData.length; i++) {
      if (labelVariantsData[i][0] === this.testName) {
        numVariants++;
      }
    }
    return numVariants;
  }

  // Visualisation metric chart

  showRelevantVariantsOnVisMetricChart() {
    var visMetricData = this.workbook.getSheetByName(sheetNames.metricData.visMetric.mvt);
    visMetricData.showColumns(3, this.numVariants);
    if (this.numVariants < 6) {
      visMetricData.hideColumns(this.numVariants + 3, 6 - this.numVariants);
    }
  }

  // Rate metric charts

  _defineRateSeriesProperties(sheetType) {
    var properties = {};
    var variantsToPlot = this.numVariants;
    if (sheetType === 'twoVar') {
      variantsToPlot = 2;
    }
    for (var i = 0; i < variantsToPlot; i++) {
      properties[i] = {
        lineWidth: this.chartProps.thinestLine,
        visibleInLegend: false,
        color: colours.variants[i]
      };
      properties[i + variantsToPlot] = properties[i];
      properties[i + (2 * variantsToPlot)] = {
        lineWidth: this.chartProps.thickestLine,
        lineDashStyle: [10, 2],
        color: colours.variants[i]
      };
    }
    return properties;
  }

  _addRateTimeseriesChart() {
    for (var i in this.sheetTypes) {
      var dataSheet = this.workbook.getSheetByName(
        sheetNames.metricData[this.objType][this.sheetTypes[i]]
      );
      var targetSheet = this.outputSheets[this.sheetTypes[i]];
      var yPosition = this.chartProps.yPosition[targetSheet.getName()];
      targetSheet.insertChart(
        targetSheet.newChart()
          .setPosition(1, 1, this.chartProps.xPosition, yPosition)
          .setChartType(Charts.ChartType.LINE)
          .addRange(dataSheet.getRange('B7:B747')) // Date Index
          .addRange(dataSheet.getRange('AA7:AR747')) // Rates + Bounds
          .setNumHeaders(this.chartProps.numHeaders)
          .setOption('height', 770)
          .setOption('width', this.chartProps.fullWidth)
          .setOption('title', 'Distribution Of ' + this.metric + ' For All Variants Over Time')
          .setOption('titleTextStyle', this.chartProps.headerTextStyle)
          .setOption('hAxis', this.chartProps.getAxisProperties('Date Index', true, true))
          .setOption('vAxis', this.chartProps.getAxisProperties(this.metric, true, true))
          .setOption('series', this._defineRateSeriesProperties(this.sheetTypes[i]))
          .setOption('curveType', this.chartProps.curveType)
          .setOption('legend', this.chartProps.legend)
          .build()
      )
    }
  }

  // Volume metric charts

  _addCombinedVolumeDistributionChart() {
    for (var i in this.sheetTypes) {
      var dataSheet = this.workbook.getSheetByName(
        sheetNames.metricData[this.objType][this.sheetTypes[i]]
      );
      var targetSheet = this.outputSheets[this.sheetTypes[i]];
      var yPosition = this.chartProps.yPosition[targetSheet.getName()];
      targetSheet.insertChart(
        targetSheet.newChart()
          .setPosition(1, 1, this.chartProps.xPosition, yPosition)
          .setChartType(Charts.ChartType.LINE)
          .addRange(dataSheet.getRange('Y3:AE104'))
          .setNumHeaders(this.chartProps.numHeaders)
          .setOption('width', this.chartProps.fullWidth)
          .setOption('title', 'Distribution Of Daily ' + this.metric + ' For All Variants')
          .setOption('titleTextStyle', this.chartProps.headerTextStyle)
          .setOption('hAxis', this.chartProps.getAxisProperties(this.metric, false, true))
          .setOption('vAxis', this.chartProps.getAxisProperties('Probability Density', true, true))
          .setOption('colors', colours.variants)
          .setOption('lineWidth', this.chartProps.thickLine)
          .setOption('curveType', this.chartProps.curveType)
          .setOption('legend', this.chartProps.legend)
          .build()
      )
    }
  }

  _addVariantVolumeDistributionCharts() {
    for (var i in this.sheetTypes) {
      var variantsToPlot = this.numVariants;
      if (this.sheetTypes[i] === 'twoVar') {
        variantsToPlot = 2;
      }
      for (var v = 0; v < variantsToPlot; v++) {
        this._addVariantVolumeDistributionChart(this.sheetTypes[i], v);
      }
    }
  }

  _addVariantVolumeDistributionChart(sheetType, variantNumber) {
    var variantRowIndex = Math.floor(variantNumber / 3);
    var variantColIndex = variantNumber - (variantRowIndex * 3)
    var xOffset =  variantColIndex * 528;
    var yOffset = (variantRowIndex + 1) * 400;

    var dataSheet = this.workbook.getSheetByName(
        sheetNames.metricData[this.objType][sheetType]
    );

    var targetSheet = this.outputSheets[sheetType];
    var yPosition = this.chartProps.yPosition[targetSheet.getName()];
    targetSheet.insertChart(
        targetSheet.newChart()
        .setPosition(1, 1, this.chartProps.xPosition + xOffset, yPosition + yOffset)
        .setChartType(Charts.ChartType.COMBO)
        .addRange(dataSheet.getRange('K3:K104')) // Objective Metric
        .addRange(dataSheet.getRange(3, 11 + variantNumber + 1, 101)) // Frequency
        .addRange(dataSheet.getRange(3, 25 + variantNumber + 1, 101)) // Probability
        .setNumHeaders(this.chartProps.numHeaders)
        .setOption('width', this.chartProps.thirdWidth)
        .setOption('title', 'Distribution Of Variant ' +  (variantNumber + 1) + ' Daily ' + this.metric)
        .setOption('titleTextStyle', this.chartProps.headerTextStyle)
        .setOption('hAxis', this.chartProps.getAxisProperties(this.metric, false, true))
        .setOption('vAxes', {
          0: this.chartProps.getAxisProperties('Frequency', true, true),
          1: this.chartProps.getAxisProperties('Probability Density', true, true)})
        .setOption('series', {
          0: {
            color: colours.grey,
            type: 'bars',
            targetAxisIndex: 0
          },
          1: {
            color: colours.variants[variantNumber],
            type: 'line',
            curveType: this.chartProps.curveType,
            targetAxisIndex: 1
          }
        })
        .setOption('lineWidth', this.chartProps.regularLine)
        .setOption('legend', this.chartProps.legend)
        .build()
    )
  }
}


class ChartPropertyFactory {

  // Defines all default properties or sub properties
  constructor() {
    this.fullWidth = 1566;
    this.thirdWidth = 506;

    this.yPosition = {}
    this.yPosition[sheetNames.drillDowns.mvt] = 1850;
    this.yPosition[sheetNames.drillDowns.twoVar] = 1750;

    this.xPosition = 70;

    this.numHeaders = 1;
    this.curveType = 'function';

    this.greyGridLines = {
      color: colours.lightGrey
    }
    this.noGridLines = {
      color: colours.white
    }

    this.headerTextStyle = {
      color: colours.navy,
      fontSize: 14
    }
    this.hiddenTextStyle = {
      color: colours.white
    }

    this.legend = {
      position: 'bottom',
      textStyle: this.textStyle
    }

    this.thinestLine = 1;
    this.thinLine = 2;
    this.regularLine = 3;
    this.thickLine = 4;
    this.thickestLine = 6;
  }

  getAxisProperties(title, showGrid, showMarks) {
    var gridlines = this.noGridLines;
    if (showGrid) {
      gridlines = this.greyGridLines;
    }
    var textStyle = this.hiddenTextStyle;
    if (showMarks) {
      textStyle = this.headerTextStyle;
    }
    return {
      title: title,
      titleTextStyle: this.headerTextStyle,
      textStyle,
      gridlines
    }
  }

}
