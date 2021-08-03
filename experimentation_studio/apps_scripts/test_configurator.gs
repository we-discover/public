/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: The logic that enables the test configuration form to function.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

// Callback on test addition form
function handleTestAdditionFormEdits(e) {
  if (e.range.getA1Notation() === 'C4') {
    var handler = new FormHandler();
    handler.clearAddedFields()
    if (handler.testType === 'Label') {
      handler.structureFormForLabels();
    }
    if (handler.testType === 'Experiment') {
      handler.structureFormForExperiment();
    }
    if (handler.testType === 'Pre Post') {
      handler.structureFormForPrepost();
    }

  }
}

// Handles form submission
function addTestFromConfiguratorForm() {
  var handler = new FormHandler();
  handler.workbook.toast('👋 Submitting test config...', 'Status', 15);
  if (handler.testType === 'Label') {
    handler.addLabelTest();
  }
  if (handler.testType === 'Experiment') {
    handler.addExperimentTest();
  }
  if (handler.testType === 'Pre Post') {
    handler.addPrepostTest();
  }
  handler.tearDownForm();
  handler.workbook.toast('Config submitted! ✅', 'Status', 3);
}

// Handles form submission
function resetTestConfiguratorForm() {
  var handler = new FormHandler();
  handler.tearDownForm();
}


class FormHandler {

  constructor() {
    this.workbook = SpreadsheetApp.getActiveSpreadsheet();
    this.form = this.workbook.getSheetByName(sheetNames.form);
    this.testType = this.form.getRange('C4').getValue();
    this.output = this.workbook.getSheetByName(sheetNames.registries[this.testType]);
  }

  tearDownForm() {
    this.form.getRange('C3:C4').setValue('');
    this.form.getRange('B5:C11').setValue('');
    this.form.getRange('B5:C11').clearDataValidations();
  }

  // === Dynamic form structure

  _getDataValidationRule(field) {
    var options = this.workbook.getSheetByName(sheetNames.options);
    if (field === 'Label Type') {
      return SpreadsheetApp.newDataValidation()
        .requireValueInRange(options.getRange('D2:D4'))
        .build();
    }
    if (field === 'Variants') {
      return SpreadsheetApp.newDataValidation()
        .requireValueInRange(options.getRange('J3:J7'))
        .build();
    }

  }

  structureFormForLabels() {
    this.form.getRange('B5').setValue('Start Date');
    this.form.getRange('C5').setValue(
      Utilities.formatDate(new Date(), "GMT", defaultDateFormat)
    );

    this.form.getRange('B6').setValue('Label Type')
    this.form.getRange('C6').setValue('ads');
    this.form.getRange('C6').setDataValidation(
      this._getDataValidationRule('Label Type')
    );

    this.form.getRange('B7').setValue('Variants');
    this.form.getRange('C7').setValue('3');
    this.form.getRange('C7').setDataValidation(
      this._getDataValidationRule('Variants')
    );
  }

  structureFormForExperiment() {
    this.form.getRange('B5').setValue('Start Date');
    this.form.getRange('C5').setValue(
      Utilities.formatDate(new Date(), "GMT", defaultDateFormat)
    );
  }

  structureFormForPrepost() {
    this.form.getRange('B5').setValue('Label Type')
    this.form.getRange('C5').setValue('ads');
    this.form.getRange('C5').setDataValidation(
      this._getDataValidationRule('Label Type')
    );

    this.form.getRange('B6').setValue('Pre Start Date');
    this.form.getRange('C6').setValue(
      Utilities.formatDate((new Date()).addDays(-14), "GMT", defaultDateFormat)
    );
    this.form.getRange('B7').setValue('Pre End Date');
    this.form.getRange('C7').setValue(
      Utilities.formatDate((new Date()).addDays(-7), "GMT", defaultDateFormat)
    );

    this.form.getRange('B8').setValue('Post Start Date');
    this.form.getRange('C8').setValue(
      Utilities.formatDate((new Date()).addDays(-6), "GMT", defaultDateFormat)
    );
    this.form.getRange('B9').setValue('Post End Date');
    this.form.getRange('C9').setValue(
      Utilities.formatDate(new Date(), "GMT", defaultDateFormat)
    );
  }

  clearAddedFields() {
    this.form.getRange('B5:C11').setValue('');
    this.form.getRange('B5:C11').clearDataValidations();
  }

  // === Form submision handlers

  _getNextTestId() {
    // Test ID increments for each addition
    var idRef = 'B2:B'
    if (this.testType === 'Experiment') {
      idRef = 'A2:A'
    }
    var currentId = Math.max.apply(null, this.output.getRange(idRef).getValues());
    return (currentId || 0) + 1;
  }

  addLabelTest() {
    var inputs = {
      testName:   this.form.getRange('C3').getValue(),
      startDate:  this.form.getRange('C5').getValue(),
      labelType:  this.form.getRange('C6').getValue(),
      variants:   this.form.getRange('C7').getValue()
    }
    if ((inputs.testName || '').length < 1) {
      throw new Error('Test name cannot be blank or undefined')
    }
    var testId = this._getNextTestId();
    for (var i = 1; i <= inputs.variants; i++) {
      var outputRow = getFirstEmptyRow(this.output);
      this.output.getRange(outputRow, 1).setValue(inputs.labelType);
      this.output.getRange(outputRow, 2).setValue(testId);
      this.output.getRange(outputRow, 3).setValue(inputs.testName)
      this.output.getRange(outputRow, 4).setValue(i);
      this.output.getRange(outputRow, 5).setValue('Variant ' + i);
      this.output.getRange(outputRow, 7).setValue(inputs.startDate);
      this.output.getRange(outputRow, 9).setValue(true);
    }
  }

  addExperimentTest() {
    var inputs = {
      testName:   this.form.getRange('C3').getValue(),
      startDate:  this.form.getRange('C5').getValue(),
    }
    if ((inputs.testName || '').length < 1) {
      throw new Error('Test name cannot be blank or undefined')
    }
    var outputRow = getFirstEmptyRow(this.output);
    this.output.getRange(outputRow, 1).setValue(this._getNextTestId());
    this.output.getRange(outputRow, 2).setValue(inputs.testName)
    this.output.getRange(outputRow, 4).setValue(inputs.startDate);
    this.output.getRange(outputRow, 6).setValue(true);
  }

  addPrepostTest() {
    var inputs = {
      testName:       this.form.getRange('C3').getValue(),
      labelType:      this.form.getRange('C5').getValue(),
      preStartDate:   this.form.getRange('C6').getValue(),
      preEndDate:     this.form.getRange('C7').getValue(),
      postStartDate:  this.form.getRange('C8').getValue(),
      postEndDate:    this.form.getRange('C9').getValue()
    }
    if ((inputs.testName || '').length < 1) {
      throw new Error('Test name cannot be blank or undefined')
    }
    var variants = ['Pre', 'Post'];
    var testId = this._getNextTestId();
    for (var i = 0; i < variants.length; i++) {
      var outputRow = getFirstEmptyRow(this.output);
      this.output.getRange(outputRow, 1).setValue(inputs.labelType);
      this.output.getRange(outputRow, 2).setValue(testId);
      this.output.getRange(outputRow, 3).setValue(inputs.testName)
      this.output.getRange(outputRow, 4).setValue(variants[i]);
      if (variants[i] === 'Pre') {
        this.output.getRange(outputRow, 6).setValue(inputs.preStartDate);
        this.output.getRange(outputRow, 7).setValue(inputs.preEndDate);
      }
      if (variants[i] === 'Post') {
        this.output.getRange(outputRow, 6).setValue(inputs.postStartDate);
        this.output.getRange(outputRow, 7).setValue(inputs.postEndDate);
      }
      this.output.getRange(outputRow, 8).setValue(true);
    }
  }

}
