// V2 Data Loader - loads all data, normalizes, parses KPI targets
const fs = require('fs');
const path = require('path');

function loadDataV2(dataDir) {
  // Portfolio master data
  const portfolio = JSON.parse(fs.readFileSync(path.join(dataDir, 'primechange_portfolio_analysis.json'), 'utf8'));

  // Analysis data (1-7)
  function loadAnalysis(n) {
    try { return JSON.parse(fs.readFileSync(path.join(dataDir, 'analysis_' + n + '_data.json'), 'utf8')); }
    catch(e) { return null; }
  }
  const analyses = {};
  for (var i = 1; i <= 7; i++) { analyses[i] = loadAnalysis(i); }

  // Revenue data
  var revenueData = {};
  try { revenueData = JSON.parse(fs.readFileSync(path.join(dataDir, 'hotel_revenue_data.json'), 'utf8')); }
  catch(e) { console.warn('hotel_revenue_data.json not found'); }

  // Hotel details (individual analysis files)
  var hotelFiles = fs.readdirSync(dataDir).filter(function(f) {
    return f.endsWith('_analysis.json') && !f.includes('portfolio');
  });
  var hotelDetails = {};
  hotelFiles.forEach(function(f) {
    try {
      var d = JSON.parse(fs.readFileSync(path.join(dataDir, f), 'utf8'));
      var key = d.hotel_key || f.replace('_analysis.json', '');
      hotelDetails[key] = d;
    } catch(e) { /* skip */ }
  });

  // Parse KPI targets
  var kpiTargets = parseKPITargets(portfolio.kpi_framework);

  // Parse per-hotel targets
  var perHotelTargets = {};
  if (portfolio.kpi_framework && portfolio.kpi_framework.per_hotel_targets) {
    portfolio.kpi_framework.per_hotel_targets.forEach(function(h) {
      perHotelTargets[h.hotel] = h;
    });
  }

  // Get regression slope for revenue calculations
  var revparSlope = 0;
  if (analyses[6] && analyses[6].regression_results && analyses[6].regression_results.score_vs_revpar) {
    revparSlope = analyses[6].regression_results.score_vs_revpar.slope;
  }

  return {
    portfolio: portfolio,
    meta: portfolio.report_metadata,
    pov: portfolio.portfolio_overview,
    cleanDive: portfolio.cleaning_deep_dive,
    priMatrix: portfolio.priority_matrix,
    actionPlans: portfolio.action_plans,
    crossRec: portfolio.cross_cutting_recommendations,
    kpi: portfolio.kpi_framework,
    roi: portfolio.roi_estimation,
    analyses: analyses,
    revenueData: revenueData,
    hotelDetails: hotelDetails,
    kpiTargets: kpiTargets,
    perHotelTargets: perHotelTargets,
    revparSlope: revparSlope
  };
}

// Parse portfolio_targets strings into numeric values
function parseKPITargets(kpiFramework) {
  if (!kpiFramework || !kpiFramework.portfolio_targets) return {};

  var targets = {};
  kpiFramework.portfolio_targets.forEach(function(t) {
    var currentVal = parseTargetString(t.current);
    var targetVal = parseTargetString(t.target);
    targets[t.kpi] = {
      label: t.kpi,
      current: currentVal,
      target: targetVal,
      deadline: t.deadline,
      raw: t
    };
  });
  return targets;
}

// Convert target strings like "8.89", "2.0%以下", "83.4%以上" to numbers
function parseTargetString(s) {
  if (s == null) return null;
  s = String(s);
  // Remove non-numeric suffixes
  var cleaned = s.replace(/[%以下以上点]+/g, '').trim();
  var val = parseFloat(cleaned);
  return isNaN(val) ? null : val;
}

module.exports = { loadDataV2, parseKPITargets, parseTargetString };
