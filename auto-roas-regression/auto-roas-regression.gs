/*
    Name:           WeDiscover - ROAS Profit Maximiser
    Description:    This script uses Quadratic Regression to analyse historical campaign
                    performance data (via the Traffic Simulator) to determine the specific 
                    ROAS (Return On Ad Spend) target that maximises expected profit.

                    It offers flexible configuration options including:
                    1. Multi-Campaign Support: Enter a list of names to run in batch
                    2. "Run All" Mode: Leave the list empty to process ALL eligible campaigns
                    3. No Data Filtering: All simulator points are used (Threshold = 0)
                    4. Portfolio Support: Can update Shared Bidding Strategies

    License:        https://github.com/we-discover/public/blob/master/LICENSE
    Version:        1.0.0
    Released:       2026-01-15
    Author:         Nathan Ifill (@nathanifill)
    Contact:        scripts@we-discover.com
*/

/**
 * ==============================================================================
 * CONFIGURATION (THE CONTROL PANEL)
 * ==============================================================================
 */

const CONFIG = {
  // 1. CAMPAIGN NAMES
  // Enter specific campaign names inside the brackets, separated by commas.
  // Example: ['Campaign A', 'Campaign B']
  // LEAVE EMPTY [] to run across ALL campaigns with available data.
  campaignNames: [], 
  
  // 2. UPDATE MODE
  // true  = ACTION MODE: Will actually update the Target ROAS in Google Ads.
  // false = READ ONLY: Will only log the maths and recommendation to the console.
  updateCampaigns: false
};

/**
 * ==============================================================================
 * MAIN SCRIPT LOGIC
 * ==============================================================================
 */

function main() {
  logHeader("STARTING BATCH ANALYSIS");

  // --- STEP 1: FETCH SIMULATION DATA FOR ALL RELEVANT CAMPAIGNS ---
  // We query the 'campaign_simulation' report. This report inherently only 
  // returns campaigns that have enough data to generate a simulation.
  
  const query = `
    SELECT 
      campaign.id, 
      campaign.name,
      campaign_simulation.target_roas_point_list.points 
    FROM campaign_simulation 
    WHERE 
      campaign_simulation.type = 'TARGET_ROAS'
  `;

  const report = AdsApp.search(query);
  let processedCount = 0;

  while (report.hasNext()) {
    const row = report.next();
    const currentName = row.campaign.name;

    // --- STEP 2: FILTERING ---
    // If the user provided specific names, check if this campaign is in the list.
    // If the list is empty [], we assume the user wants to process everything.

    if (CONFIG.campaignNames.length > 0 && CONFIG.campaignNames.indexOf(currentName) === -1) {
      continue; 
    }

    // --- STEP 3: PROCESS THE INDIVIDUAL CAMPAIGN ---

    processSingleCampaign(row);
    processedCount++;
  }

  if (processedCount === 0) {
    Logger.log("\n[INFO] No eligible campaigns found matching your criteria.");
    Logger.log("Note: Google only generates simulations for campaigns with sufficient conversion data.");
  } else {
    logHeader(`BATCH COMPLETE: Processed ${processedCount} Campaign(s)`);
  }
}

/**
 * CORE LOGIC: Analyses and updates a single campaign found in the loop.
 */

function processSingleCampaign(row) {
  const campaignName = row.campaign.name;
  logHeader(`ANALYSING: "${campaignName}"`);

  const points = row.campaignSimulation.targetRoasPointList.points;
  
  // --- CALCULATE PROFIT ---
  // Convert raw metrics into "Profit" using a 1.4x value multiplier (margin).
  // Formula: (Conversion Value * 1.4) - Cost
  
  let data = [];

  // Log table header only once per campaign
  Logger.log("ROAS (x)".padEnd(12) + "| ADJ. PROFIT (y)");
  Logger.log("".padEnd(30, "-"));

  for (const point of points) {
    const roas = point.targetRoas;
    const cost = point.costMicros / 1000000;
    const value = point.biddableConversionsValue;
    const profit = (value * 1.0) - cost;
    
    // HARD CODED: No filtering (Min Threshold = 0)
    data.push({ x: roas, y: profit });

    // Minimal logging for batch view
    Logger.log(
      roas.toFixed(2).padEnd(12) + 
      "| " + formatCurrency(profit)
    );
  }

  // Need at least 3 points for a curve
  if (data.length < 3) {
    Logger.log("[SKIPPED] Not enough datapoints to calculate a curve (Need 3+).");
    return;
  }

  // --- RUN THE MATHS (REGRESSION) ---
  
  const coeffs = fitQuadratic(data);
  const a = coeffs.a;
  
  // Safety Check: If curve is U-shaped (positive 'a'), profit increases forever.
  if (a >= 0) {
    Logger.log("[SKIPPED] Data shows profit increasing indefinitely (U-Curve). Unsafe to optimise.");
    return;
  }

  // Calculate Peak (Vertex)
  const b = coeffs.b;
  const c = coeffs.c;
  const optimalRoas = -b / (2 * a);
  const maxProfit = (a * (optimalRoas * optimalRoas)) + (b * optimalRoas) + c;

  // Rounding
  const cleanRoas = parseFloat(optimalRoas.toFixed(2));

  // Sanity Check: Ensure ROAS isn't negative or absurd
  if (cleanRoas <= 0) {
      Logger.log(`[SKIPPED] Calculated ROAS (${cleanRoas}) is invalid.`);
      return;
  }

  // --- RESULTS ---
  
  Logger.log("".padEnd(30, "-"));
  Logger.log(`RECOMMENDATION: ROAS ${cleanRoas} (Exp. Profit: ${formatCurrency(maxProfit)})`);

  // --- APPLY CHANGES ---
  
  if (CONFIG.updateCampaigns) {
    applyTargetRoas(row.campaign.id, cleanRoas);
  } else {
    Logger.log("STATUS:         [READ ONLY]");
  }
}


/**
 * ==============================================================================
 * HELPER FUNCTIONS (UTILITIES)
 * ==============================================================================
 */

function formatCurrency(amount) {
  return amount.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

function logHeader(title) {
  Logger.log("\n" + "".padEnd(60, "="));
  Logger.log(title);
  Logger.log("".padEnd(60, "="));
}

// Least Squares fitting for quadratic curve
function fitQuadratic(data) {
  let s4 = 0, s3 = 0, s2 = 0, s1 = 0, s0 = 0;
  let sy = 0, sxy = 0, sx2y = 0;
  
  for (const p of data) {
    const x = p.x; const y = p.y; const x2 = x * x;
    s4 += x2 * x2; s3 += x2 * x; s2 += x2; s1 += x; s0 += 1;
    sy += y; sxy += x * y; sx2y += x2 * y;
  }
  
  const A = [[s4, s3, s2], [s3, s2, s1], [s2, s1, s0]];
  const B = [sx2y, sxy, sy];
  
  return solve3x3(A, B);
}

function solve3x3(A, B) {
  const det = (m) => m[0][0]*(m[1][1]*m[2][2]-m[1][2]*m[2][1]) - m[0][1]*(m[1][0]*m[2][2]-m[1][2]*m[2][0]) + m[0][2]*(m[1][0]*m[2][1]-m[1][1]*m[2][0]);
  const D = det(A);
  
  if (Math.abs(D) < 1e-9) throw new Error("Maths Error: Singular Matrix");
  
  const createMatrix = (col, vec) => {
    let m = JSON.parse(JSON.stringify(A));
    for(let i=0; i<3; i++) m[i][col] = vec[i];
    return m;
  };
  
  return { a: det(createMatrix(0, B)) / D, b: det(createMatrix(1, B)) / D, c: det(createMatrix(2, B)) / D };
}

/**
 * ==============================================================================
 * ADVANCED UPDATE: HANDLES PORTFOLIO STRATEGIES
 * ==============================================================================
 */

function applyTargetRoas(campaignId, roas) {
  const campaignIterator = AdsApp.campaigns().withIds([campaignId]).get();
  
  if (campaignIterator.hasNext()) {
    const campaign = campaignIterator.next();
    const portfolioStrategy = campaign.bidding().getStrategy(); 

    if (portfolioStrategy) {
      Logger.log(`INFO: Linked to Portfolio Strategy: "${portfolioStrategy.getName()}"`);
      
      try {
        updatePortfolioViaMutate(portfolioStrategy, roas);
      } catch (e) {
        Logger.log(`[ERROR] Failed to update Portfolio Strategy: ${e.message}`);
      }

    } else {
      const strategy = campaign.getBiddingStrategyType();
      
      if (strategy === 'TARGET_ROAS' || strategy === 'MAXIMIZE_CONVERSION_VALUE') {
         campaign.bidding().setTargetRoas(roas);
         Logger.log(`SUCCESS: Standard Campaign Target ROAS updated to ${roas.toFixed(2)}`);
      } else {
        Logger.log(`ERROR: Campaign uses '${strategy}', which cannot be updated this way.`);
      }
    }
  }
}

/**
 * Sends a direct JSON command to the Google Ads API.
 */

function updatePortfolioViaMutate(portfolioStrategy, roas) {
  const strategyId = portfolioStrategy.getId();
  const type = portfolioStrategy.getType();
  const customerId = AdsApp.currentAccount().getCustomerId().replace(/-/g, '');
  
  const resourceNameString = `customers/${customerId}/biddingStrategies/${strategyId}`;

  // 1. Prepare the inner operation
  
  let innerOperation = {
    "update": { 
      "resourceName": resourceNameString
    },
    "update_mask": "" 
  };

  if (type === 'TARGET_ROAS') {
    innerOperation.update.targetRoas = { "targetRoas": roas };
    innerOperation.update_mask = "target_roas.target_roas";
  } else if (type === 'MAXIMIZE_CONVERSION_VALUE') {
    innerOperation.update.maximizeConversionValue = { "targetRoas": roas };
    innerOperation.update_mask = "maximize_conversion_value.target_roas";
  } else {
    Logger.log(`[ERROR] Portfolio type '${type}' is not supported by this script.`);
    return;
  }

  // 2. WRAP IT: 'bidding_strategy_operation' (snake_case)
  
  const payload = {
    "bidding_strategy_operation": innerOperation
  };

  try {
    // 3. Send Mutate
    
    const response = AdsApp.mutate(payload);
    
    // 4. Validate Response
    
    if (response.isSuccessful()) {
       Logger.log(`SUCCESS: Portfolio Target ROAS updated to ${roas}`);
    } else {
       Logger.log(`[API FAILURE] The request was rejected.`);
       Logger.log(`Errors: ${response.getErrorMessages().join(", ")}`);
    }
    
  } catch (e) {
    Logger.log(`[CRITICAL API ERROR] ${e.toString()}`);
  }
}
