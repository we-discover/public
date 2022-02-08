# Harvest API exporter to Google Sheets

This Google Apps Script lets you query a bulk Harvest API endpoint and load the results to a GSheet.

While this is a fairly crude application, this is a situation we encountered with an urgent priority
and so we thought we'd share the love to help anyone in a similar situation.



### Usage

1. Create a new [Google Apps Script Project](https://script.google.com/home).
2. Copy the contents of each file - with a .gs extension - in this folder to a corresponding file.
3. Set your Harvest API  as the appropriate userProperties by running `setUserProperties` in 
`properties.gs`. This is one of the most secure ways to handle credentials in Google Apps Scripts.
(see [credentials](#credentials) for more detail).
4. Populate the `exportConfigs` in `main.gs` with details of the endpoints and destinations sheets 
that you wish export (see [export configuration](#export-configuration) for more details).
5. (Optional) Schedule this script to run on a regular basis.



### Credentials

Credentials are managed using using the Google Apps Scripts 
[PropertiesService](https://developers.google.com/apps-script/reference/properties). In this script,
these properties can be set by following the instructions in `properties.gs`. They can then be 
removed from whatever code is saved in your project.

The credentials required are listed below:

Credential                  | Description
----------------------------| -------------------------------------------------------------------------
`HARVEST_ACCESS_TOKEN`      | Your PAT, generated in the Harvest settings tab.
`HARVEST_ACCOUNT_ID`        | The Harvest account ID that you wish to export data from.



### Export configuration 

The object `exportConfigs` in `main.gs` is the section of this script in which you can specify the
endpoints that you wish to export data from and where the outputs should be written.

This export structure was suitable for the task for which this script was originally developed but 
there's a lot of opportunity for adapting and extending this. If you do, get in touch, we'd love to
hear if you've made some cool improvements here. Importing only new data and using queryparams to
export sub-sets of data are probable next steps.

The schema for a specific configuration is as follows, comments are available in `main.gs` that show
what is expected for these values.

```json
{
    "name": STRING,
    "path": STRING
}
```
