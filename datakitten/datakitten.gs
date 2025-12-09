 /**
 * 🐱 DataKitten
 * * * Description: Monitors hourly Google Ads performance using statistical anomaly detection.
 * * * Setup: Add Slack Webhook URL in SLACK_WEBHOOK_URL below. Schedule to run Hourly.
 */


const CONFIG = {
  // 1. Webhook URL is loaded from the secure constant above
  SLACK_WEBHOOK_URL: '', 


  // Sensitivity: How many standard deviations (SD) before we alert?
  // 2.5 = Approx 1 in 80 events (Balanced)
  SENSITIVITY_THRESHOLD: 2.5, 
  
  // Conversion Lag (Days):
  // Applies ONLY to metrics marked with 'useLag: true' (Conversions, CPA).
  // 0 = Check Today's data for Conversions (if you're using online conversions).
  // 1 = Check Yesterday's data for Conversions (Safe for OCI).
  CONVERSION_LAG: 1,

  CLIENT_NAME: AdsApp.currentAccount().getName(),

  METRICS: [
    { 
      id: 'impressions',
      name: 'Impr.', 
      field: 'metrics.impressions', 
      jsonKey: 'impressions',
      useLag: false, // Real-time
      isCurrency: false,
      isPercentage: false,
      definition: "Count of how often your ad was shown.",
      action: "Check for budget caps or seasonality."
    },
    { 
      id: 'clicks',
      name: 'Clicks', 
      field: 'metrics.clicks', 
      jsonKey: 'clicks',
      useLag: false, // Real-time
      isCurrency: false,
      isPercentage: false,
      definition: "How many times users clicked your ads.",
      action: "Check CTR or ad relevance."
    },
    { 
      id: 'ctr',
      name: 'CTR', 
      field: 'metrics.ctr', 
      jsonKey: 'ctr',
      useLag: false, // Real-time
      isCurrency: false,
      isPercentage: true,
      definition: "Click-through rate (Clicks / Impr).",
      action: "Check ad copy or creative fatigue."
    },
    { 
      id: 'cost',
      name: 'Cost', 
      field: 'metrics.cost_micros', 
      jsonKey: 'costMicros',
      useLag: false, // Real-time
      isCurrency: true,
      isPercentage: false,
      definition: "Total spend accumulated.",
      action: "Check for runaway bids or broad match spikes."
    },
    { 
      id: 'avgCpc',
      name: 'Avg. CPC', 
      field: 'metrics.average_cpc', 
      jsonKey: 'averageCpc',
      useLag: false, // Real-time
      isCurrency: true,
      isPercentage: false,
      definition: "Average amount paid per click.",
      action: "Check competition levels or Quality Score."
    },
    { 
      id: 'conversions',
      name: 'Conversions', 
      field: 'metrics.conversions', 
      jsonKey: 'conversions',
      useLag: true, // LAGGED (Wait for OCI)
      isCurrency: false,
      isPercentage: false,
      definition: "Count of conversion actions.",
      action: "Check tracking tags or site checkout health."
    },
    { 
      id: 'cpa',
      name: 'Cost / Conv.', 
      field: 'metrics.cost_per_conversion', 
      jsonKey: 'costPerConversion',
      useLag: true, // LAGGED (Wait for OCI)
      isCurrency: true,
      isPercentage: false,
      definition: "Average cost per conversion (CPA).",
      action: "Check CVR drops or CPC spikes."
    },
    { 
      id: 'absTop',
      name: 'Search abs. top impr. rate', 
      field: 'metrics.absolute_top_impression_percentage', 
      jsonKey: 'absoluteTopImpressionPercentage',
      useLag: false, // Real-time
      isCurrency: false,
      isPercentage: true,
      definition: "% of your ads that appeared as the very first result (Position #1).",
      action: "Check Quality Score or new competitors."
    }
  ]
};

function main() {
  Logger.log("🐱 DataKitten is waking up...");
  Logger.log("------------------------------------------");

  const accountTimeZone = AdsApp.currentAccount().getTimeZone();
  
  // --- 1. Define Timelines ---
  
  // Timeline A: Real-Time (Today, Previous Hour)
  // Used for: Cost, Clicks, Impressions
  const realTimeDateObj = new Date();
  realTimeDateObj.setHours(realTimeDateObj.getHours() - 1);
  const realTimeDateStr = Utilities.formatDate(realTimeDateObj, accountTimeZone, "yyyy-MM-dd");
  
  // Timeline B: Lagged (Today - Lag Days, Previous Hour)
  // Used for: Conversions, CPA
  const laggedDateObj = new Date();
  laggedDateObj.setHours(laggedDateObj.getHours() - 1);
  laggedDateObj.setDate(laggedDateObj.getDate() - CONFIG.CONVERSION_LAG);
  const laggedDateStr = Utilities.formatDate(laggedDateObj, accountTimeZone, "yyyy-MM-dd");

  // Shared Hour (e.g., 2pm)
  const targetHourStr = Utilities.formatDate(realTimeDateObj, accountTimeZone, "H");
  const targetHour = parseInt(targetHourStr, 10);

  Logger.log(`🌍 Account Timezone: ${accountTimeZone}`);
  Logger.log(`🕒 Analysing Hour: ${targetHour}:00 - ${targetHour}:59`);
  Logger.log(`📅 Real-Time Date: ${realTimeDateStr} (For Traffic/Spend)`);
  Logger.log(`📅 Lagged Date:    ${laggedDateStr} (For Conversions - Lag: ${CONFIG.CONVERSION_LAG} days)`);

  // --- 2. Fetch Data (Double Fetch) ---
  
  Logger.log("------------------------------------------");

  // Fetch Real-Time Data & History (Baseline excludes Today)
  const rtCurrent = fetchSpecificDateStats(targetHour, realTimeDateStr);
  const rtHistory = fetchHistoryStats(targetHour, realTimeDateStr);

  // Fetch Lagged Data & History (Baseline excludes Lagged Date onwards)
  const lagCurrent = fetchSpecificDateStats(targetHour, laggedDateStr);
  const lagHistory = fetchHistoryStats(targetHour, laggedDateStr);

  if ((!rtCurrent || rtCurrent.impressions === 0) && (!lagCurrent || lagCurrent.impressions === 0)) {
    Logger.log("😴 No data found for either timeline. Kitten is going back to sleep.");
    return;
  }

  // --- 3. Merge Data based on Config ---
  
  const finalCurrent = {};
  const finalHistory = {};

  CONFIG.METRICS.forEach(metric => {
    if (metric.useLag) {
        // Use Lagged Data source
        finalCurrent[metric.id] = lagCurrent[metric.id];
        finalHistory[metric.id] = lagHistory[metric.id];
    } else {
        // Use Real-Time Data source
        finalCurrent[metric.id] = rtCurrent[metric.id];
        finalHistory[metric.id] = rtHistory[metric.id];
    }
  });

  // --- 4. Run Analysis ---
  
  Logger.log("------------------------------------------");
  checkAnomalies(finalCurrent, finalHistory);
}

// --- ANALYSIS LOGIC ---

function checkAnomalies(current, history) {
  CONFIG.METRICS.forEach(metric => {
    const historicalValues = history[metric.id];
    
    // Safety check for data volume
    if (!historicalValues || historicalValues.length < 5) {
      Logger.log(`⚠️ Not enough historical data to analyse ${metric.name} (${historicalValues ? historicalValues.length : 0} points).`);
      return;
    }

    const mean = getMean(historicalValues);
    const sd = getStdDev(historicalValues, mean);
    let currentValue = current[metric.id] || 0;
    
    // Skip if flatline
    if (sd === 0) {
      if (mean === 0) {
         Logger.log(`ℹ️ ${metric.name}: All historical values are 0. Skipping.`);
      } else {
         Logger.log(`ℹ️ ${metric.name}: Standard Deviation is 0 (Stable). Skipping.`);
      }
      return;
    }

    const zScore = (currentValue - mean) / sd;
    const formattedCurrent = formatValue(currentValue, metric);
    const formattedMean = formatValue(mean, metric);
    
    Logger.log(`📊 Analysing ${metric.name}:`);
    Logger.log(`   Current: ${formattedCurrent} | 30-Day Mean: ${formattedMean} | SD: ${sd.toFixed(2)}`);
    Logger.log(`   Z-Score: ${zScore.toFixed(2)}`);

    if (Math.abs(zScore) > CONFIG.SENSITIVITY_THRESHOLD) {
      const direction = zScore > 0 ? "High Spike 📈" : "Significant Drop 📉";
      Logger.log(`🚨 MEOW ALERT! ${metric.name}: ${direction} (${zScore.toFixed(1)} SDs from norm)`);
      
      sendSlackAlert(metric, currentValue, mean, zScore);
      
    } else {
      Logger.log(`✅ Normal behaviour.`);
    }
    Logger.log("------------------------------------------");
  });
}

// --- SLACK ALERT GENERATOR ---

function sendSlackAlert(metric, currentVal, normalVal, zScore) {
  const isDrop = zScore < 0;
  const emoji = isDrop ? "📉" : "📈";
  const severity = Math.abs(zScore) > 3 ? "🚨 URGENT" : "⚠️ Warning";
  const accountId = AdsApp.currentAccount().getCustomerId().replace(/-/g, "");
  const deepLink = `https://ads.google.com/aw/overview?__u=${accountId}`;

  // Add context to the alert title if it's a lagged metric
  const timeContext = metric.useLag ? `(Lagged Data: ${CONFIG.CONVERSION_LAG}d ago)` : "(Real-Time Data)";

  const slackPayload = {
    "blocks": [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": `${severity}: ${CONFIG.CLIENT_NAME} Anomaly`,
          "emoji": true
        }
      },
      {
        "type": "section",
        "fields": [
          {
            "type": "mrkdwn",
            "text": `*Metric:*\n${metric.name} ${metric.useLag ? "🕒" : "⚡"}`
          },
          {
            "type": "mrkdwn",
            "text": `*Deviation:*\n${zScore.toFixed(1)}σ (Sigma)`
          }
        ]
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `${emoji} *Current Value:* ${formatValue(currentVal, metric)}\nTypically, we expect around *${formatValue(normalVal, metric)}* at this time of day.\n_${timeContext}_`
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "context",
        "elements": [
          {
            "type": "mrkdwn",
            "text": `🎓 *What is this?* ${metric.definition}`
          }
        ]
      },
      {
        "type": "section",
        "text": {
            "type": "mrkdwn",
            "text": `💡 *Recommended Action:* ${metric.action}`
        },
        "accessory": {
          "type": "button",
          "text": {
            "type": "plain_text",
            "text": "View Account",
            "emoji": true
          },
          "url": deepLink,
          "action_id": "button-action"
        }
      }
    ]
  };

  if (CONFIG.SLACK_WEBHOOK_URL && CONFIG.SLACK_WEBHOOK_URL.length > 0) {
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(slackPayload)
    };
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, options);
    Logger.log("✅ Slack alert sent successfully.");
  } else {
    Logger.log("⚠️ Slack Webhook URL is missing in CONFIG. Alert not sent.");
  }
}

// --- DATA FETCHING ---

function fetchHistoryStats(hour, targetDateStr) {
  Logger.log(`📥 Fetching BASELINE history for hour: ${hour}...`);
  Logger.log(`   (Excluding dates >= ${targetDateStr})`);
  
  const query = `
    SELECT 
      segments.date,
      segments.hour,
      metrics.impressions, 
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.conversions,
      metrics.cost_per_conversion,
      metrics.absolute_top_impression_percentage
    FROM ad_group
    WHERE segments.date DURING LAST_30_DAYS
    AND segments.hour = ${hour}
    AND metrics.impressions > 0
  `;

  const report = AdsApp.search(query);
  const dailyData = {}; 

  while (report.hasNext()) {
    const row = report.next();
    const date = row.segments.date;

    // Filter out the Test Date and any dates NEWER than it (the lag gap)
    if (date >= targetDateStr) {
      continue; 
    }

    if (!dailyData[date]) {
      dailyData[date] = { imps: 0, clicks: 0, cost: 0, weightedCtr: 0, weightedCpc: 0, conv: 0, weightedCpa: 0, weightedAbsTop: 0 };
    }

    // Process Metrics
    const imps = parseInt(row.metrics.impressions || 0);
    const clicks = parseInt(row.metrics.clicks || 0);
    const cost = (parseInt(row.metrics.costMicros || 0)) / 1000000;
    const ctr = parseFloat(row.metrics.ctr || 0);
    const avgCpc = (parseInt(row.metrics.averageCpc || 0)) / 1000000;
    const conv = parseFloat(row.metrics.conversions || 0);
    const cpa = (parseInt(row.metrics.costPerConversion || 0)) / 1000000;
    const absTop = parseFloat(row.metrics.absoluteTopImpressionPercentage || 0);

    dailyData[date].imps += imps;
    dailyData[date].clicks += clicks;
    dailyData[date].cost += cost;
    dailyData[date].weightedCtr += (ctr * imps);
    dailyData[date].weightedCpc += (avgCpc * clicks);
    dailyData[date].conv += conv;
    dailyData[date].weightedCpa += (cpa * conv);
    dailyData[date].weightedAbsTop += (absTop * imps);
  }

  const finalData = { 'impressions': [], 'clicks': [], 'ctr': [], 'cost': [], 'avgCpc': [], 'conversions': [], 'cpa': [], 'absTop': [] };
  
  const sortedDates = Object.keys(dailyData).sort().reverse(); 

  for (const date of sortedDates) {
    const d = dailyData[date];
    
    const finalCtr = d.imps > 0 ? (d.weightedCtr / d.imps) : 0;
    const finalCpc = d.clicks > 0 ? (d.weightedCpc / d.clicks) : 0;
    const finalCpa = d.conv > 0 ? (d.weightedCpa / d.conv) : 0;
    const finalAbsTop = d.imps > 0 ? (d.weightedAbsTop / d.imps) : 0;
    
    finalData['impressions'].push(d.imps);
    finalData['clicks'].push(d.clicks);
    finalData['ctr'].push(finalCtr);
    finalData['cost'].push(d.cost);
    finalData['avgCpc'].push(finalCpc);
    finalData['conversions'].push(d.conv);
    finalData['cpa'].push(finalCpa);
    finalData['absTop'].push(finalAbsTop);
  }
  
  return finalData;
}

function fetchSpecificDateStats(hour, dateStr) {
  // Use simple single-day filter to avoid BAD_VALUE errors with DURING syntax
  
  const query = `
    SELECT 
      segments.hour,
      metrics.impressions, 
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.conversions,
      metrics.cost_per_conversion,
      metrics.absolute_top_impression_percentage
    FROM ad_group
    WHERE segments.date = '${dateStr}'
    AND segments.hour = ${hour}
    AND metrics.impressions > 0
  `;

  const report = AdsApp.search(query);
  
  let totalImps = 0;
  let totalClicks = 0;
  let totalCost = 0;
  let weightedCtr = 0;
  let weightedCpc = 0;
  let totalConv = 0;
  let weightedCpa = 0;
  let weightedAbsTop = 0;
  
  while (report.hasNext()) {
    const row = report.next();
    
    const imps = parseInt(row.metrics.impressions || 0);
    const clicks = parseInt(row.metrics.clicks || 0);
    const cost = (parseInt(row.metrics.costMicros || 0)) / 1000000;
    const ctr = parseFloat(row.metrics.ctr || 0);
    const avgCpc = (parseInt(row.metrics.averageCpc || 0)) / 1000000;
    const conv = parseFloat(row.metrics.conversions || 0);
    const cpa = (parseInt(row.metrics.costPerConversion || 0)) / 1000000;
    const absTop = parseFloat(row.metrics.absoluteTopImpressionPercentage || 0);
    
    totalImps += imps;
    totalClicks += clicks;
    totalCost += cost;
    weightedCtr += (ctr * imps);
    weightedCpc += (avgCpc * clicks);
    totalConv += conv;
    weightedCpa += (cpa * conv);
    weightedAbsTop += (absTop * imps);
  }

  const finalCtr = totalImps > 0 ? (weightedCtr / totalImps) : 0;
  const finalCpc = totalClicks > 0 ? (weightedCpc / totalClicks) : 0;
  const finalCpa = totalConv > 0 ? (weightedCpa / totalConv) : 0;
  const finalAbsTop = totalImps > 0 ? (weightedAbsTop / totalImps) : 0;

  return {
    'impressions': totalImps,
    'clicks': totalClicks,
    'ctr': finalCtr,
    'cost': totalCost,
    'avgCpc': finalCpc,
    'conversions': totalConv,
    'cpa': finalCpa,
    'absTop': finalAbsTop
  };
}

// --- MATHS & FORMATTING ---

function getMean(data) {
  if (data.length === 0) return 0;
  const sum = data.reduce((a, b) => a + b, 0);
  return sum / data.length;
}

function getStdDev(data, mean) {
  if (data.length === 0) return 0;
  const squareDiffs = data.map(x => Math.pow(x - mean, 2));
  const avgSquareDiff = squareDiffs.reduce((a, b) => a + b, 0) / data.length;
  return Math.sqrt(avgSquareDiff);
}

function formatValue(value, metricConfig) {
  if (metricConfig.isCurrency) return "£" + value.toFixed(2);
  if (metricConfig.isPercentage) return (value * 100).toFixed(2) + "%";
  return value.toFixed(2);
}
