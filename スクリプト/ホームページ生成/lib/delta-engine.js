// V2 Delta Engine - snapshot diff calculation
const fs = require('fs');
const path = require('path');

function calcDeltas(outputDir, currentSummary) {
  var deltas = { hasDeltas: false, alerts: [], metrics: {} };

  // Load snapshot index to find previous snapshot
  var snapshotIndexPath = path.join(outputDir, 'data', 'snapshot-index.json');
  if (!fs.existsSync(snapshotIndexPath)) return deltas;

  var snapshots;
  try { snapshots = JSON.parse(fs.readFileSync(snapshotIndexPath, 'utf8')); }
  catch(e) { return deltas; }

  if (!snapshots || snapshots.length < 1) return deltas;

  // Previous snapshot is the last one in the index
  var prev = snapshots[snapshots.length - 1];
  if (!prev) return deltas;

  // Load previous portfolio summary
  var prevSummaryPath = path.join(outputDir, 'data', 'snapshots', prev.id, 'portfolio-summary.json');
  var prevSummary;
  try { prevSummary = JSON.parse(fs.readFileSync(prevSummaryPath, 'utf8')); }
  catch(e) { return deltas; }

  deltas.hasDeltas = true;
  deltas.previousDate = prev.id;
  deltas.previousSummary = prevSummary;

  // Calculate metric deltas
  var metrics = {};
  var fields = ['avg_score', 'high_rate', 'cleaning_issue_rate', 'total_reviews'];
  fields.forEach(function(f) {
    var curr = currentSummary[f];
    var prevVal = prevSummary[f];
    if (curr != null && prevVal != null) {
      metrics[f] = {
        current: curr,
        previous: prevVal,
        delta: Math.round((curr - prevVal) * 100) / 100
      };
    }
  });
  deltas.metrics = metrics;

  // Generate alerts based on thresholds
  var alerts = [];
  if (metrics.avg_score) {
    if (metrics.avg_score.delta <= -0.1) {
      alerts.push({
        type: 'danger',
        icon: '&#9888;&#65039;',
        title: 'スコア低下警告',
        message: 'ポートフォリオ平均スコアが ' + metrics.avg_score.delta.toFixed(2) + 'pt 低下（' + metrics.avg_score.previous + ' → ' + metrics.avg_score.current + '）',
        severity: 'red'
      });
    } else if (metrics.avg_score.delta >= 0.05) {
      alerts.push({
        type: 'improvement',
        icon: '&#128994;',
        title: 'スコア改善',
        message: 'ポートフォリオ平均スコアが +' + metrics.avg_score.delta.toFixed(2) + 'pt 改善（' + metrics.avg_score.previous + ' → ' + metrics.avg_score.current + '）',
        severity: 'green'
      });
    }
  }

  if (metrics.cleaning_issue_rate) {
    if (metrics.cleaning_issue_rate.delta >= 2.0) {
      alerts.push({
        type: 'danger',
        icon: '&#128680;',
        title: '清掃クレーム率上昇',
        message: '清掃クレーム率が +' + metrics.cleaning_issue_rate.delta.toFixed(1) + '%上昇（' + metrics.cleaning_issue_rate.previous + '% → ' + metrics.cleaning_issue_rate.current + '%）',
        severity: 'red'
      });
    }
  }

  if (metrics.total_reviews) {
    var reviewDelta = metrics.total_reviews.delta;
    if (reviewDelta > 0) {
      alerts.push({
        type: 'info',
        icon: '&#128172;',
        title: '新規口コミ',
        message: '+' + reviewDelta + '件の新規口コミを取得',
        severity: 'blue'
      });
    }
  }

  deltas.alerts = alerts;
  return deltas;
}

module.exports = { calcDeltas };
