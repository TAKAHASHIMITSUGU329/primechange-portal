#!/bin/bash
# 既存バックアップからスナップショットを生成する（1回限りのスクリプト）
# 使い方: bash スクリプト/最新化/convert_backup_to_snapshot.sh [backup_dir] [date]
# 例:     bash スクリプト/最新化/convert_backup_to_snapshot.sh backup_2026-03-12 2026-03-12

set -e
BASE_DIR="/Users/mitsugutakahashi/ホテル口コミ"
JSON_DIR="$BASE_DIR/データ/分析結果JSON"
OUTPUT_DIR="$BASE_DIR/ホームページ"

BACKUP_NAME="${1:-backup_2026-03-12}"
SNAPSHOT_DATE="${2:-2026-03-12}"

BACKUP_DIR="$JSON_DIR/$BACKUP_NAME"
SNAPSHOT_DIR="$OUTPUT_DIR/data/snapshots/$SNAPSHOT_DATE"

if [ ! -d "$BACKUP_DIR" ]; then
  echo "ERROR: バックアップディレクトリが見つかりません: $BACKUP_DIR"
  exit 1
fi

echo "=== バックアップからスナップショット生成 ==="
echo "  バックアップ: $BACKUP_DIR"
echo "  スナップショット日付: $SNAPSHOT_DATE"
echo "  出力先: $SNAPSHOT_DIR"

# Node.jsで変換実行
node -e "
var fs = require('fs');
var path = require('path');

var backupDir = '$BACKUP_DIR';
var snapshotDir = '$SNAPSHOT_DIR';
var snapshotDate = '$SNAPSHOT_DATE';
var outputDir = '$OUTPUT_DIR';

// Create snapshot directory
if (!fs.existsSync(snapshotDir)) fs.mkdirSync(snapshotDir, { recursive: true });

// Load hotel analysis files from backup
var files = fs.readdirSync(backupDir).filter(function(f) {
  return f.endsWith('_analysis.json') && !f.includes('portfolio') && !f.startsWith('analysis_');
});

var hotelDetails = {};
var allReviewsCompact = {};
var allDates = [];

var CLEANING_KEYWORDS = ['清掃', '汚れ', 'ゴミ', '髪の毛', 'シミ', 'カビ', 'ほこり', '埃', '汚い', '不潔', '臭い', 'におい', '匂い', 'ホコリ', 'しみ', 'かび', 'ごみ'];

files.forEach(function(file) {
  var key = file.replace('_analysis.json', '');
  var data = JSON.parse(fs.readFileSync(path.join(backupDir, file), 'utf8'));
  var allComments = (data.comments || []).map(function(c) {
    var d = c.date || '';
    if (d) allDates.push(d);
    return { site: c.site, rating_10pt: c.rating_10pt, date: d,
      comment: (c.translated || c.comment || '').slice(0, 300),
      good: (c.translated_good || c.good || '').slice(0, 200),
      bad: (c.translated_bad || c.bad || '').slice(0, 200) };
  });
  hotelDetails[key] = {
    total_reviews: data.total_reviews, overall_avg_10pt: data.overall_avg_10pt,
    high_count: data.high_count, high_rate: data.high_rate,
    mid_count: data.mid_count, mid_rate: data.mid_rate,
    low_count: data.low_count, low_rate: data.low_rate,
    site_stats: data.site_stats, distribution: data.distribution,
    comments: allComments.slice(0, 30)
  };
  allReviewsCompact[key] = allComments.map(function(c) {
    return { s: c.site, r: c.rating_10pt, d: c.date, c: c.comment, g: c.good, b: c.bad };
  });
});

allDates.sort();
var dateMin = allDates.length > 0 ? allDates[0] : '';
var dateMax = allDates.length > 0 ? allDates[allDates.length - 1] : '';
var totalReviews = allDates.length;

// Calculate portfolio KPIs
var portfolioScoreSum = 0, portfolioHighCount = 0, portfolioCleanCount = 0;
Object.keys(allReviewsCompact).forEach(function(key) {
  allReviewsCompact[key].forEach(function(r) {
    var score = parseFloat(r.r) || 0;
    portfolioScoreSum += score;
    if (score >= 8) portfolioHighCount++;
    var text = (r.c || '') + (r.g || '') + (r.b || '');
    for (var i = 0; i < CLEANING_KEYWORDS.length; i++) {
      if (text.indexOf(CLEANING_KEYWORDS[i]) !== -1) { portfolioCleanCount++; break; }
    }
  });
});

var avgScore = totalReviews > 0 ? Math.round(portfolioScoreSum / totalReviews * 100) / 100 : 0;
var highRate = totalReviews > 0 ? Math.round(portfolioHighCount / totalReviews * 1000) / 10 : 0;
var cleanRate = totalReviews > 0 ? Math.round(portfolioCleanCount / totalReviews * 1000) / 10 : 0;

// Write snapshot files
fs.writeFileSync(path.join(snapshotDir, 'hotel-reviews-all.json'), JSON.stringify(allReviewsCompact), 'utf8');
fs.writeFileSync(path.join(snapshotDir, 'hotel-details.json'), JSON.stringify(hotelDetails), 'utf8');
fs.writeFileSync(path.join(snapshotDir, 'build-meta.json'), JSON.stringify({
  build_date: snapshotDate,
  data_range: { min: dateMin, max: dateMax },
  total_reviews: totalReviews,
  snapshot_id: snapshotDate
}), 'utf8');
fs.writeFileSync(path.join(snapshotDir, 'portfolio-summary.json'), JSON.stringify({
  total_hotels: Object.keys(allReviewsCompact).length,
  total_reviews: totalReviews,
  avg_score: avgScore,
  high_rate: highRate,
  cleaning_issue_rate: cleanRate,
  cleaning_issue_count: portfolioCleanCount
}), 'utf8');

console.log('  hotel-reviews-all.json: ' + Object.keys(allReviewsCompact).length + 'ホテル, ' + totalReviews + '件');
console.log('  hotel-details.json: ' + Object.keys(hotelDetails).length + 'ホテル');
console.log('  build-meta.json: ' + dateMin + ' 〜 ' + dateMax);
console.log('  portfolio-summary.json: avg=' + avgScore + ', high=' + highRate + '%, clean=' + cleanRate + '%');

// Update snapshot-index.json
var indexPath = path.join(outputDir, 'data', 'snapshot-index.json');
var index = [];
try { index = JSON.parse(fs.readFileSync(indexPath, 'utf8')); } catch(e) {}

var existingIdx = index.findIndex(function(s) { return s.id === snapshotDate; });
var entry = {
  id: snapshotDate,
  date: snapshotDate,
  total_reviews: totalReviews,
  avg_score: avgScore,
  high_rate: highRate,
  cleaning_issue_rate: cleanRate,
  data_range: { min: dateMin, max: dateMax }
};

if (existingIdx >= 0) {
  index[existingIdx] = entry;
} else {
  index.push(entry);
}
index.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

fs.writeFileSync(indexPath, JSON.stringify(index), 'utf8');
console.log('  snapshot-index.json updated: ' + index.length + '件のスナップショット');
"

echo ""
echo "=== 完了 ==="
echo "  スナップショット: $SNAPSHOT_DIR"
ls -la "$SNAPSHOT_DIR/"
