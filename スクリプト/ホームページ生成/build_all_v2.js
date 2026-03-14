// PRIMECHANGE V2 Dashboard - Build Orchestrator
// Generates 7 pages in ホームページV2/
const fs = require('fs');
const path = require('path');

// Foundation modules
const { loadDataV2 } = require('./lib/data-loader');
const { esc, writeCommonCSS, copyAssets } = require('./lib/common-v2');
const { calcDeltas } = require('./lib/delta-engine');
const { calcAllRevenueOpportunities, formatYen } = require('./lib/revenue-calc');
const { analyzeCS, getKeywordFrequency } = require('./lib/cs-analyzer');

// Page generators
const { buildIndex } = require('./lib/page-index');
const { buildHotelDashboard } = require('./lib/page-hotel-dashboard');
const { buildCleaningStrategy } = require('./lib/page-cleaning-strategy');
const { buildDeepAnalysis } = require('./lib/page-deep-analysis');
const { buildRevenueImpact } = require('./lib/page-revenue-impact');
const { buildActionPlans } = require('./lib/page-action-plans');
const { buildESDashboard } = require('./lib/page-es-dashboard');

// Deep analysis renderers (reused from V1)
// Already required by page-deep-analysis.js

// Paths
const DATA_DIR = path.join(__dirname, '..', '..', 'データ', '分析結果JSON');
const OUTPUT_DIR = path.join(__dirname, '..', '..', 'ホームページV2');

console.log('=== PRIMECHANGE V2 Dashboard Build ===');
console.log('Data source: ' + DATA_DIR);
console.log('Output: ' + OUTPUT_DIR);
console.log('');

// Create output directories
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
['styles', 'scripts', 'data'].forEach(function(d) {
  var dir = path.join(OUTPUT_DIR, d);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ============================================================
// Phase 1: Load all data
// ============================================================
console.log('[Phase 1] Loading data...');
var data = loadDataV2(DATA_DIR);
console.log('  Portfolio: ' + (data.meta ? data.meta.total_hotels : '?') + ' hotels');
console.log('  Analyses: ' + Object.keys(data.analyses).filter(function(k) { return data.analyses[k]; }).length + '/7 loaded');
console.log('  Revenue data: ' + Object.keys(data.revenueData).length + ' hotels');
console.log('  Hotel details: ' + Object.keys(data.hotelDetails).length + ' files');
console.log('  RevPAR slope: ' + data.revparSlope);

// ============================================================
// Phase 1b: Process hotel reviews (same as V1)
// ============================================================
console.log('[Phase 1b] Processing reviews...');
var hotelDetailsProcessed = {};
var allReviewsCompact = {};
var allDates = [];

var hotelFiles = fs.readdirSync(DATA_DIR).filter(function(f) {
  return f.endsWith('_analysis.json') && !f.includes('portfolio');
});

hotelFiles.forEach(function(file) {
  var key = file.replace('_analysis.json', '');
  var d = JSON.parse(fs.readFileSync(path.join(DATA_DIR, file), 'utf8'));
  var allComments = (d.comments || []).map(function(c) {
    var date = c.date || '';
    if (date) allDates.push(date);
    return {
      site: c.site, rating_10pt: c.rating_10pt, date: date,
      comment: (c.translated || c.comment || '').slice(0, 300),
      good: (c.translated_good || c.good || '').slice(0, 200),
      bad: (c.translated_bad || c.bad || '').slice(0, 200)
    };
  });
  hotelDetailsProcessed[key] = {
    total_reviews: d.total_reviews, overall_avg_10pt: d.overall_avg_10pt,
    high_count: d.high_count, high_rate: d.high_rate,
    mid_count: d.mid_count, mid_rate: d.mid_rate,
    low_count: d.low_count, low_rate: d.low_rate,
    site_stats: d.site_stats, distribution: d.distribution,
    comments: allComments.slice(0, 30)
  };
  allReviewsCompact[key] = allComments.map(function(c) {
    return { s: c.site, r: c.rating_10pt, d: c.date, c: c.comment, g: c.good, b: c.bad };
  });
});

allDates.sort();
var dateMin = allDates.length > 0 ? allDates[0] : '';
var dateMax = allDates.length > 0 ? allDates[allDates.length - 1] : '';
var totalAllReviews = allDates.length;
console.log('  Reviews: ' + totalAllReviews + ' (' + dateMin + ' ~ ' + dateMax + ')');

// ============================================================
// Phase 2: Calculate derived data
// ============================================================
console.log('[Phase 2] Calculating derived data...');

// Portfolio summary
var CLEANING_KEYWORDS = ['清掃', '汚れ', 'ゴミ', '髪の毛', 'シミ', 'カビ', 'ほこり', '埃', '汚い', '不潔', '臭い', 'におい', '匂い', 'ホコリ', 'しみ', 'かび', 'ごみ'];
function hasCleaningIssue(text) {
  for (var i = 0; i < CLEANING_KEYWORDS.length; i++) {
    if (text.indexOf(CLEANING_KEYWORDS[i]) !== -1) return true;
  }
  return false;
}

var ptRev = 0, ptSum = 0, ptHigh = 0, ptClean = 0;
Object.keys(allReviewsCompact).forEach(function(key) {
  allReviewsCompact[key].forEach(function(r) {
    ptRev++;
    var score = parseFloat(r.r) || 0;
    ptSum += score;
    if (score >= 8) ptHigh++;
    if (hasCleaningIssue((r.c || '') + (r.g || '') + (r.b || ''))) ptClean++;
  });
});

var portfolioAvg = ptRev > 0 ? Math.round(ptSum / ptRev * 100) / 100 : 0;
var portfolioHighRate = ptRev > 0 ? Math.round(ptHigh / ptRev * 1000) / 10 : 0;
var portfolioCleanRate = ptRev > 0 ? Math.round(ptClean / ptRev * 1000) / 10 : 0;

var portfolioSummary = {
  total_hotels: Object.keys(allReviewsCompact).length,
  total_reviews: ptRev,
  avg_score: portfolioAvg,
  high_rate: portfolioHighRate,
  cleaning_issue_rate: portfolioCleanRate,
  cleaning_issue_count: ptClean
};
console.log('  Portfolio: avg=' + portfolioAvg + ' high=' + portfolioHighRate + '% clean=' + portfolioCleanRate + '%');

// Revenue opportunities
var revenueOps = calcAllRevenueOpportunities(data);
var totalOpportunity = 0;
Object.keys(revenueOps).forEach(function(k) { totalOpportunity += revenueOps[k].monthlyLoss; });
console.log('  Revenue opportunity: ¥' + formatYen(totalOpportunity) + '/月');

// Delta calculation
var deltas = calcDeltas(OUTPUT_DIR, portfolioSummary);
console.log('  Deltas: ' + (deltas.hasDeltas ? deltas.alerts.length + ' alerts' : 'no previous data'));

// CS analysis
console.log('  Running CS 6-axis analysis...');
var csResults = analyzeCS(data.hotelDetails);
var keywordFreq = getKeywordFrequency(data.hotelDetails);
console.log('  CS: ' + Object.keys(csResults).length + ' hotels, ' + keywordFreq.length + ' keywords');

// ============================================================
// Phase 3: Write shared data files
// ============================================================
console.log('[Phase 3] Writing data files...');

function writeJSON(filename, jsonData) {
  fs.writeFileSync(path.join(OUTPUT_DIR, 'data', filename), JSON.stringify(jsonData), 'utf8');
  console.log('  data/' + filename);
}

writeJSON('hotel-reviews-all.json', allReviewsCompact);
writeJSON('hotel-details.json', hotelDetailsProcessed);
writeJSON('hotel-ranked.json', data.pov.hotels_ranked);
writeJSON('tier-color.json', { '優秀': '#10B981', '良好': '#3B82F6', '概ね良好': '#F59E0B', '要改善': '#EF4444' });
writeJSON('portfolio-summary.json', portfolioSummary);
writeJSON('deltas.json', deltas);

var buildDate = new Date().toISOString().slice(0, 10);
var buildMetaData = {
  build_date: buildDate,
  data_range: { min: dateMin, max: dateMax },
  total_reviews: totalAllReviews,
  snapshot_id: buildDate,
  version: 'V2'
};
writeJSON('build-meta.json', buildMetaData);

// Write inline data JS (for file:// protocol compatibility - fetch() doesn't work with file://)
var inlineDataJS = [
  'window.__REVIEWS_DATA__ = ' + JSON.stringify(allReviewsCompact) + ';',
  'window.__BUILD_META__ = ' + JSON.stringify(buildMetaData) + ';',
  'window.__DELTAS_DATA__ = ' + JSON.stringify(deltas) + ';'
].join('\n');
fs.writeFileSync(path.join(OUTPUT_DIR, 'data', 'inline-data.js'), inlineDataJS, 'utf8');
console.log('  data/inline-data.js');

// Write assets
writeCommonCSS(OUTPUT_DIR);
copyAssets(OUTPUT_DIR);

// ============================================================
// Phase 4: Generate pages
// ============================================================
console.log('[Phase 4] Generating pages...');

function writePage(filename, content) {
  fs.writeFileSync(path.join(OUTPUT_DIR, filename), content, 'utf8');
  console.log('  ' + filename);
}

// 1. Index (Portal)
var indexHtml = buildIndex(data, deltas, revenueOps);
writePage('index.html', indexHtml);

// 2. Hotel Dashboard
var dashboardHtml = buildHotelDashboard(data, revenueOps);
writePage('hotel-dashboard.html', dashboardHtml);

// 3. Cleaning Strategy
var cleaningHtml = buildCleaningStrategy(data, revenueOps);
writePage('cleaning-strategy.html', cleaningHtml);

// 4. Deep Analysis
var deepResult = buildDeepAnalysis(data, csResults, keywordFreq);
if (typeof deepResult === 'object' && deepResult.html) {
  writePage('deep-analysis.html', deepResult.html);
  if (deepResult.snapshotContent) {
    writeJSON('deep-analysis-content.json', deepResult.snapshotContent);
  }
} else {
  writePage('deep-analysis.html', deepResult);
}

// 5. Revenue Impact
var revenueResult = buildRevenueImpact(data, revenueOps);
if (typeof revenueResult === 'object' && revenueResult.html) {
  writePage('revenue-impact.html', revenueResult.html);
  if (revenueResult.snapshotContent) {
    writeJSON('revenue-impact-content.json', revenueResult.snapshotContent);
  }
} else {
  writePage('revenue-impact.html', revenueResult);
}

// 6. Action Plans
var actionResult = buildActionPlans(data, revenueOps);
if (typeof actionResult === 'object' && actionResult.html) {
  writePage('action-plans.html', actionResult.html);
  if (actionResult.snapshotContent) {
    writeJSON('action-plans-content.json', actionResult.snapshotContent);
  }
  if (actionResult.actionStatus) {
    writeJSON('action-status.json', actionResult.actionStatus);
  }
} else {
  writePage('action-plans.html', actionResult);
}

// 7. ES Dashboard (new!)
var esHtml = buildESDashboard(data);
writePage('es-dashboard.html', esHtml);

// ============================================================
// Phase 5: Snapshot management
// ============================================================
console.log('[Phase 5] Managing snapshots...');

var snapshotDir = path.join(OUTPUT_DIR, 'data', 'snapshots', buildDate);
if (!fs.existsSync(snapshotDir)) fs.mkdirSync(snapshotDir, { recursive: true });

// Copy snapshot files
['hotel-reviews-all.json', 'hotel-details.json', 'build-meta.json', 'portfolio-summary.json'].forEach(function(f) {
  var src = path.join(OUTPUT_DIR, 'data', f);
  if (fs.existsSync(src)) {
    fs.copyFileSync(src, path.join(snapshotDir, f));
  }
});

// Copy content HTML fragments to snapshot
['deep-analysis-content.json', 'revenue-impact-content.json', 'action-plans-content.json'].forEach(function(f) {
  var src = path.join(OUTPUT_DIR, 'data', f);
  if (fs.existsSync(src)) fs.copyFileSync(src, path.join(snapshotDir, f));
});

console.log('  data/snapshots/' + buildDate + '/');

// Update snapshot index
var snapshotIndexPath = path.join(OUTPUT_DIR, 'data', 'snapshot-index.json');
var snapshotIndex = [];
try { snapshotIndex = JSON.parse(fs.readFileSync(snapshotIndexPath, 'utf8')); } catch(e) {}

var existingIdx = snapshotIndex.findIndex(function(s) { return s.id === buildDate; });
var snapshotEntry = {
  id: buildDate,
  date: buildDate,
  total_reviews: ptRev,
  avg_score: portfolioAvg,
  high_rate: portfolioHighRate,
  cleaning_issue_rate: portfolioCleanRate,
  data_range: { min: dateMin, max: dateMax },
  content_files: ['deep-analysis', 'revenue-impact', 'action-plans'],
  version: 'V2'
};

if (existingIdx >= 0) {
  snapshotIndex[existingIdx] = snapshotEntry;
} else {
  snapshotIndex.push(snapshotEntry);
}
snapshotIndex.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
writeJSON('snapshot-index.json', snapshotIndex);

// Update inline-data.js with snapshot index (now finalized)
inlineDataJS += '\nwindow.__SNAPSHOT_INDEX__ = ' + JSON.stringify(snapshotIndex) + ';';
fs.writeFileSync(path.join(OUTPUT_DIR, 'data', 'inline-data.js'), inlineDataJS, 'utf8');
console.log('  data/inline-data.js (updated with snapshot index)');

// ============================================================
// Done!
// ============================================================
console.log('');
console.log('=== Build Complete! ===');
console.log('Output: ' + OUTPUT_DIR);
console.log('Pages: index.html, hotel-dashboard.html, cleaning-strategy.html,');
console.log('       deep-analysis.html, revenue-impact.html, action-plans.html,');
console.log('       es-dashboard.html');
console.log('');
console.log('To view: open ' + path.join(OUTPUT_DIR, 'index.html'));
