# AWS Athena exporter to Google Sheets

![](/assets/export_from_athena_to_gsheet__1.png)

This Google Apps Script lets you query Athena directly and export the results to a GSheet.

While this is a fairly niche use case, this is a situation we encountered for a client that was not
trivial to solve so we thought we'd share the love to help anyone in a similar situation.


## Usage

1. Create a new [Google Apps Script Project](https://script.google.com/home).
2. Copy the contents of each file - with a .gs` extension - in this folder to a corresponding file.
3. Set your AWS creddentials as the appropriate userProperties by running `setUserProperties` in 
`properties.gs`. This is one of the most secure ways to handle credentials in Google Apps Scripts.
(see [credentials](#credentials) for more detail).
4. Populate the `exportConfigs` in `main.gs` with details of the queries and destinations that you
wish export (see [export configuration](#export-configuration) for more details).
5. (Optional) Schedule this script to run on a regular basis and keep whatever you need in GSheets 
freshly up to date with data from Athena!

## Credentials

Credentials are managed using using the Google Apps Scripts 
[PropertiesService](https://developers.google.com/apps-script/reference/properties). In this script,
these properties can be set by following the instructions in `properties.gs`. They can then be 
removed from whatever code is saved in your project.

The credentials required are listed below:

Credential                  | Description
----------------------------| -------------------------------------------------------------------------
`AWS_ACCESS_KEY`            | Your AWS access key.
`AWS_SECRET_KEY`            | Your AWS secret key.
`AWS_REGION`                | The region in which your AWS Athena service is hosted.
`ATHENA_S3_OUTPUT_LOCATION` | `s3://aws-athena-query-results-{account_id}-{region}`


## Export configuration 

The object `exportConfigs` in `main.gs` is the section of this script in which you can specify the
queries that you wish to run against your specified Athena database and where the outputs should be
written.

This export structure was suitable for the task for which this script was originalyl develope but 
there's a lot of opportunity for adapting and extending this. If you do, get in touch, we'd love to
hear if you've made some cool improvements here.

The schema for a specific configuration is as follows, comments are available in `main.gs` that show
what is expected for these values.

```json
{
    "name": STRING,
    "gsheetId": STRING,
    "database": STRING,
    "query": STRING
}
```
