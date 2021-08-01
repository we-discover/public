/*
    Name:        WeDiscover - Experimentation Studio, Google Apps Script

    Description: Configuration of simple trigger logic to control implement
                 that is non-native to a GSheet.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.1

    Released:    2021-07-31

    Contact:     scripts@we-discover.com
*/

function onEdit(e) {

  const activeSheetName = e.source.getActiveSheet().getName();

  // For any edit on a sheet with controls
  if (
    sheetsWithControls.includes(activeSheetName) &&
    isInControlRange(e.range.getA1Notation())
  ) {
    handleControlChanges(e); // Toasts update starting
    toggleVisualisationRecords(e);
    toggleCellFormats(e);
    toggleCoreVisualisations(e);
    updateSummaryTable(e); // Toasts update ending
  }

  // Handle test addition form interactions
  if (activeSheetName === sheetNames.form) {
    handleTestAdditionFormEdits(e);
  }

}
