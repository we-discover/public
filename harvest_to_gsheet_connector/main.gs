/*
    Name:        WeDiscover - Harvest API to GSheet Exporter

    Description: A utility script to export data from the Harvest API to a GSheet

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.0

    Released:    2022-02-01

    Contact:     scripts@we-discover.com
*/

/**
     * Config structure for configuring data  be exported to given gsheets.
     * This export structure was suitable for the task for which this
     * script was originally developed but there's a lot of opportunity
     * for adapting and extending this.
     * 
     * Currently populated for all bulk endp
     * 
     * name:    Name, used as identifier and export sheet name
     * path:    The Harvest API endpoint to request data from (see https://help.getharvest.com/api-v2/)
*/
const exportConfigs = [
    {
        name: 'Data Import - Users', 
        path: 'users',
    },
    {
        name: 'Data Import - Roles', 
        path: 'roles',
    },    
    {
        name: 'Data Import - Clients', 
        path: 'clients',
    },    
    {
        name: 'Data Import - Tasks', 
        path: 'tasks',
    },
    {
        name: 'Data Import - Projects', 
        path: 'projects',
    },    
    {
        name: 'Data Import - Task Assignments', 
        path: 'task_assignments',
    },    
    {
        name: 'Data Import - User Assignments', 
        path: 'user_assignments',
    },        
           
    {
        name: 'Data Import - Time Entries', 
        path: 'time_entries',
    }        
]


/**
    * Loop through all configurations and export to GSheet.
    * Basic bulk export of all data returned through the API, no 
    * querystring filtering supported, but would be a straightforward
    * extension. Neglectful error handling, skips failures with 
    * somewhat useful logs
*/
function main() {
  for (var i = 0; i < exportConfigs.length; i++) {
    exportHarvestDataFromConfig(exportConfigs[i]);
  }
}


/**
    * Run the specified export from Harvest to a GSheet
    * @param {object} config - JSON object with keys as outlined above.
*/
function exportHarvestDataFromConfig(config) {

  var userProperties = PropertiesService.getUserProperties();

  const settings = {
    protocol: "https:",
    hostname: "api.harvestapp.com",
    version: "v2",
    headers: {
      "User-Agent": "WeDiscover Google Apps Script",
      "Authorization": "Bearer " + userProperties.getProperty('HARVEST_ACCESS_TOKEN'),
      "Harvest-Account-ID": userProperties.getProperty('HARVEST_ACCOUNT_ID')
    }
  }

  try {
    Logger.log('### Running export: ' + config.name);
    
    var response = UrlFetchApp.fetch(
      `${settings.protocol}//${settings.hostname}/${settings.version}/${config.path}`, {
          "method": "get",
          "headers": settings.headers
      }
    );
    Logger.log('Exporting data from Harvest API.');

    var responseContent = JSON.parse(response.getContentText());
    const isPaged = Object.keys(responseContent).includes('page');

    var results = [];

    if (!isPaged) {
      var results = responseContent;
      if (!Array.isArray(results)) {
        results = [results];
      }
    }

    if (isPaged) {
      var isLast = responseContent.links.next === null;
      results = responseContent[config.path];

      while(!isLast) {
        Logger.log(responseContent.links.next);

        var response = UrlFetchApp.fetch(responseContent.links.next, {
          "method": "get", 
          "headers": settings.headers
          }
        );
        responseContent = JSON.parse(response.getContentText());
        results = results.concat(responseContent[config.path]);
        isLast = responseContent.links.next === null;
      }
    }

    writeJsonToGsheet(config.name, results);
    Logger.log('### Completed export: ' + config.name);

  } catch(e) {
    Logger.log('### Failed export: ' + config.name);
    console.log(e);
  } 
}


/**
    * Function to load response into the objective sheet
    * @param {string} outputSheetName - The name of the intended output sheet.
    * @param {array} jsonResults - an array of JSON results returned from the API.
*/
function writeJsonToGsheet(outputSheetName, jsonResults){
  var spreadsheet = SpreadsheetApp.getActive();
  var outputSheet = spreadsheet.getSheetByName(outputSheetName);
  if (outputSheet === null) {
    outputSheet = spreadsheet.insertSheet(outputSheetName, 99);
  }
  outputSheet.clear();
  Logger.log('Connected and cleared output sheet: ' + outputSheetName);

  Logger.log('Flattening data');
  var flatJsonResults = [];
  var rowValues = []
  for (var i = 0; i < jsonResults.length; i++) {
    flatJsonResults.push(Object.flatten(jsonResults[i]));
    rowValues.push(Object.values(flatJsonResults[i]));
  }

  Logger.log('Writing data to: ' + outputSheetName);
  var columnHeaders = [Object.keys(flatJsonResults[0])];
  outputSheet.getRange(1,1, columnHeaders.length, columnHeaders[0].length)
    .setValues(columnHeaders);
    
  outputSheet.getRange(2, 1, rowValues.length, columnHeaders[0].length)
    .setValues(rowValues);
}


/**
    * Flatten method for objects
    * Adapted from method authored by: https://stackoverflow.com/users/1048572/bergi
*/
Object.flatten = function(data) {
    var result = {};
    function recurse (cur, prop) {
        if (prop === 'external_reference') {
            
        } else if (Object(cur) !== cur) {
            result[prop] = cur;
        } else if (Array.isArray(cur)) {
            result[prop] = cur.join(', ');
        } else {
            var isEmpty = true;
            for (var p in cur) {
                isEmpty = false;
                recurse(cur[p], prop ? prop+"."+p : p);
            }
            if (isEmpty && prop)
                result[prop] = {};
        }
    }
    recurse(data, "");
    return result;
}

