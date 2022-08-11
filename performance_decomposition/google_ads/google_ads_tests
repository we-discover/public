/** Note: To run the test suite, you must rename the main() function to something else so that this main() can run **/

function main() {
 runAllTests(); 
}

/******************* TEST CASES ********************************************************************************************/
// 2 element Array:
// * First element is an array containing arguments for function
// * Second element is expected result
const getHeadersCases = [
  [`
SELECT 
  segments.date, 
  customer.id,
  customer.descriptive_name,
  customer.currency_code,
  campaign.id,
  campaign.name, 
  metrics.cost_micros, 
  metrics.search_impression_share,
  metrics.impressions, 
  metrics.clicks
FROM 
  campaign 
WHERE 
  segments.date BETWEEN
`,
  ['segments.date', 'customer.id', 'customer.descriptive_name', 'customer.currency_code', 'campaign.id', 'campaign.name', 'metrics.cost_micros', 'metrics.search_impression_share', 'metrics.impressions', 'metrics.clicks']],
  [`
SELECT
  campaign.name
FROM
  campaign
WHERE
  segments.date DURING LAST_7_DAYS
  `,
  ['campaign.name']],
];

const initialiseReportsCases = [
    [
        [
            {
                name: "BASE_METRICS_PERFORMANCE_QUERY",
                query: "SELECT metrics.impressions, metrics.clicks FROM campaign WHERE segments.date BETWEEN '2022-01-01' AND '2021-02-01'",
                sheetName: "Google Ads Import: Google Ads Campaign Stats"
            },
            {
                name: "CONVERSION_PERFORMANCE_QUERY",
                query: "SELECT metrics.all_conversions, metrics.all_conversions_value FROM campaign WHERE segments.date BETWEEN '2022-01-01' AND '2021-02-01'",
                sheetName: "Google Ads Import: Google Ads Campaign Conv. Stats"
            },
        ],
        {
            "Google Ads Import: Google Ads Campaign Stats": {
                "headerOrder": ['metrics.impressions', 'metrics.clicks'],
                "metrics.impressions": [],
                "metrics.clicks": [],
            },
            "Google Ads Import: Google Ads Campaign Conv. Stats": {
                "headerOrder": ['metrics.all_conversions', 'metrics.all_conversions_value'],
                "metrics.all_conversions": [],
                "metrics.all_conversions_value": [],
            }
        }
    ]
];

/******************* TEST SUITE ********************************************************************************************/

// Keep track of no. of tests passed and failed
let tally = {"passed": 0, "failed": 0}

/**
 * Basic assert function
 * @param {function} func function to be tests
 * @param {Any} args arguments for func
 * @param {String} expectedResult value the test should return
 * @return {String} whether test passed or failed
 */
// TODO: Make log/warning not show function arguments inside square brackets
function assert(func, args, expectedResult) {
  const actualResult = func(...[args]);
  
  if(actualResult === expectedResult && typeof actualResult === typeof expectedResult) {
    console.log(`Test passed: ${func.name}(${args}) = ${expectedResult}`);
    return "passed";
  }
  
  else if (typeof actualResult === typeof expectedResult && typeof actualResult === 'object') {
    if (JSON.stringify(actualResult) === JSON.stringify(expectedResult)) {
        console.log(`Test passed: ${func.name}(${args}) = ${expectedResult}`);
        return "passed";
    }
    
    else if (JSON.stringify(actualResult) !== JSON.stringify(expectedResult)) {
    console.warn(`Test failed: ${func.name}(${args}) =/= ${expectedResult}\n\n\n
Expected: ${expectedResult}\n\n
Actual: ${actualResult}`);
    return "failed";    }
  }
  
  else if (typeof actualResult !== typeof expectedResult) {
    console.warn(`Test failed: mismatched types.
Expected: ${typeof expectedResult}
Actual: ${actualResult}`);
    
    return "failed";
  }
  
  else if(actualResult !== expectedResult) {

    console.warn(`Test failed: ${func.name}(`, args, `) =/= ${expectedResult}
Expected: ${expectedResult}
Actual: ${actualResult}`);
    return "failed";
  }
  
}

function runMultipleTests(func, cases) {
  Logger.log("** Testing " + func.name + " **");

  for (let i = 0; i < cases.length; i++) {
    const funcArgs = cases[i][0]
    const expectedResult = cases[i][1];

    let result = assert(func, funcArgs, expectedResult);
    tally[result]++
  }

  Logger.log("** Finished testing " + func.name + " **");
}
/**
 * Function to run multiple test cases using runMultipleTests
 */
function runAllTests() {
  runMultipleTests(getHeaders, getHeadersCases);
  runMultipleTests(initialiseReports, initialiseReportsCases);

  Logger.log(`Passed: ${tally["passed"]}, Failed: ${tally["failed"]}`);

  return null;
}
