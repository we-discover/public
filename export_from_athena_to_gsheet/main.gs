
/**
     * Config structure for queries to be exported to given gsheets.
     * This export structure was suitable for the task for which this
     * script was originally developed but there's a lot of opportunity
     * for adapting and extending this.
*/
const exportConfigs = [
    {
        name: '', // Name, used as identifier and export sheet name
        gsheetId: '', // GSheet ID, used for output destination
        database: '', // Athena database to run query against
        query: '' // Query to run against database
    }
]

/**
    * Loop through all configurations and export to Athena.
    * Neglectful error handling, skips failures with fairly useful logs
*/
function main() {
  for (var i = 0; i < exportConfigs.length; i++) {
    exportFromAthenaToSheet(exportConfigs[i]);
  }
}

/**
    * Run the specified export from Athena to a GSheet
    * @param {string} config - JSON object with keys as outlined above.
*/
function exportFromAthenaToSheet(config) {

    var gsheetId = config.gsheetId;
    var database = config.database;
    var query = config.query;

    Logger.log('### Running export: ' + config.name);

    // Connect to output sheet and worksheet
    try {
        var spreadsheet = SpreadsheetApp.openById(gsheetId);
        var outputSheetName = "Data Import: " + config.name;
        var outputSheet = spreadsheet.getSheetByName(outputSheetName);
        if (outputSheet === null) {
            outputSheet = spreadsheet.insertSheet(outputSheetName, 99);
        }
        outputSheet.clear();
        Logger.log('Connected to GSheet: ' + gsheetId);
    } catch (e) {
        return Logger.log('Failed to connect to GSheet: ' + gsheetId + '. Error: ' + e)
    }

    // Execute query against specified Athena database
    try {
        var execuctionResponse = Athena.runQuery(database, query);
        var queryId = execuctionResponse.QueryExecutionId;
        Logger.log('Running query: ' + queryId);
    } catch (e) {
        return Logger.log('Failed to run query. Error: ' + e)
    }

    // Wait for query to have a final state and try to load results
    try {
        Athena.waitForQueryEndState(queryId);
        Logger.log('Loading results: ' + queryId);
        var results = Athena.getQueryResults(queryId);
    } catch (e) {
        return Logger.log('Failed to load results: ' + queryId + '. Error: ' + e)
    }

    // Write results to output sheet
    try {
        var outputRange = outputSheet.getRange(1, 1, results.length, results[0].length);
        outputRange.setValues(results);
        Logger.log('Exported results to GSheet: ' + gsheetId);
    } catch (e) {
        return Logger.log('Failed to export results: ' + gsheetId   + '. Error: ' + e)
    }

    Logger.log('### Finished export: ' + config.name);

};
