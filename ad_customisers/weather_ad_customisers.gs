/*
    Name:        WeDiscover - Weather Ad Customisers, Google Apps Script

    Description: A function to get latest weather data.

    License:     https://github.com/we-discover/public/blob/master/LICENSE

    Version:     1.0.0

    Released:    2021-09-10

    Contact:     scripts@we-discover.com
*/

function getLatestData() {
  // API Key EDIT ME
  const key = "xxx";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const locationSheet = spreadsheet.getSheetByName("Location");
  const liveSheet = spreadsheet.getSheetByName("Live Data");
  const historySheet = spreadsheet.getSheetByName("History");

  const location = locationSheet.getRange("A2").getValue();
  const mainCell = liveSheet.getRange("A2");

  let apiURL = `https://api.openweathermap.org/data/2.5/weather?q=${location}&appid=${key}`;

  const resText = UrlFetchApp.fetch(apiURL).getContentText();
  console.log(resText);
  const resJSON = JSON.parse(resText);
  const mainWeather = resJSON["weather"][0]["main"];
  mainCell.setValue(mainWeather);

  historySheet.appendRow([new Date(),location, mainWeather]);
}
