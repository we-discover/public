
// Forked from: https://github.com/canyousayyes/datastudio-aws-athena-connector

var Athena = (function() {

    var userProperties = PropertiesService.getUserProperties();
    var S3OutputLocation = userProperties.getProperty('ATHENA_S3_OUTPUT_LOCATION');

    return {

        init: function Athena (S3OutputLocation) {
            if(S3OutputLocation == undefined) {
                throw "Error: No S3 output location defined for Athena";
            }
            S3OutputLocation = S3OutputLocation;
        },

        /**
         * Submit a query to AWS Athena.
         * @param {string} database Datebase to run this query.
         * @param {string} query The query string.
         * @return {Object} {"QueryExecutionId": "string"}
        */
        runQuery: function(database, query) {
            var payload = {
                'ClientRequestToken': uuidv4(),
                'QueryExecutionContext': {
                    'Database': database
                },
                'QueryString': query,
                'ResultConfiguration': {
                    'OutputLocation': S3OutputLocation
                }
            };
            return AWS.post('athena', 'AmazonAthena.StartQueryExecution', payload);
        },

        /**
         * Wait until an Athena query to reach a terminal state.
         * This is a blocking function that continuously pull the state of the query.
         * If the query finished without any errors, the function will return nothing.
         * Otherwise an exception will be thrown.
         * @param  {string} queryExecutionId The submitted execution ID.
        */
        waitForQueryEndState: function(queryExecutionId) {
            var payload = {
                'QueryExecutionId': queryExecutionId
            };
            while (1) {
                var result = AWS.post('athena', 'AmazonAthena.GetQueryExecution', payload);
                var state = result.QueryExecution.Status.State.toLowerCase();
                switch (state) {
                    case 'succeeded':
                        return true;
                    case 'failed':
                        throw new Error(result.QueryExecution.Status.StateChangeReason || 'Unknown query error');
                    case 'cancelled':
                        throw new Error('Query cancelled');
                }
                Utilities.sleep(5000);
            }
        },

        /**
         * Fetch all rows from a submitted AWS Athena query.
         * @param  {string} queryExecutionId The submitted execution ID.
         * @return {Array} Array of rows, in the form of { val0, val1, ... }, includes header
        */
        getQueryResults: function(queryExecutionId) {
            var rows = [];
            var nextToken = null;
            while (1) {
                var payload = {
                    'QueryExecutionId': queryExecutionId
                };
                if (nextToken) {
                    payload.NextToken = nextToken;
                }
                var result = AWS.post('athena', 'AmazonAthena.GetQueryResults', payload);
                result.ResultSet.Rows.forEach(function (row) {
                    var newRow = [];
                    row.Data.forEach(function (data, index) {
                        newRow.push(data.VarCharValue);
                    });
                    rows.push(newRow);
                });
                nextToken = result.NextToken;
                if (!nextToken) {
                    break;
                }
            }
            return rows
        }

    };

})();
