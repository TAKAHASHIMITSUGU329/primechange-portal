// V2 CS Analyzer - 6-axis keyword classification + NPS estimation

// CS 6 axes with keywords
var CS_AXES = {
  '接客態度': ['接客', 'スタッフ', 'フロント', '対応', '態度', '丁寧', '親切', '笑顔', '挨拶', 'サービス', '従業員', 'ホスピタリティ', '気配り', '不愛想', '無愛想'],
  '立地': ['立地', '駅', 'アクセス', '近い', '便利', '徒歩', '周辺', 'コンビニ', '繁華街', '交通', '場所', 'ロケーション'],
  '朝食': ['朝食', '朝ごはん', '朝ご飯', 'モーニング', 'ブレックファスト', 'バイキング', 'ビュッフェ', '食事', '料理', 'パン', 'おかず'],
  '設備': ['設備', '部屋', 'ベッド', 'バス', 'トイレ', 'シャワー', 'エアコン', 'テレビ', 'Wi-Fi', 'wifi', 'WiFi', 'アメニティ', '枕', 'タオル', '冷蔵庫', '家具', '古い', '新しい', '綺麗', 'きれい', 'リノベ', '改装'],
  '清掃': ['清掃', '掃除', '汚れ', 'ゴミ', '髪の毛', 'シミ', 'カビ', 'ほこり', '埃', '汚い', '不潔', '臭い', 'におい', '匂い', 'ホコリ', '清潔'],
  'コスパ': ['コスパ', 'コストパフォーマンス', '価格', '値段', '料金', '安い', '高い', 'お得', 'リーズナブル', '割高', '値段相応', '妥当']
};

function analyzeCS(hotelDetails) {
  var results = {};

  Object.keys(hotelDetails).forEach(function(key) {
    var hotel = hotelDetails[key];
    var reviews = hotel.comments || hotel.reviews || [];
    if (reviews.length === 0) return;

    var axisScores = {};
    var axisCounts = {};
    Object.keys(CS_AXES).forEach(function(axis) {
      axisScores[axis] = { positive: 0, negative: 0, total: 0 };
      axisCounts[axis] = 0;
    });

    var promoters = 0, detractors = 0, total = reviews.length;

    reviews.forEach(function(r) {
      var score = parseFloat(r.rating_10pt || r.r || 0);
      var comment = r.translated || r.comment || r.c || '';
      var goodText = r.translated_good || r.good || r.g || '';
      var badText = r.translated_bad || r.bad || r.b || '';
      var text = comment + ' ' + goodText + ' ' + badText;

      // NPS calculation (10-point scale: 9-10=promoter, 1-6=detractor)
      if (score >= 9) promoters++;
      else if (score <= 6) detractors++;

      // Axis keyword matching
      Object.keys(CS_AXES).forEach(function(axis) {
        var keywords = CS_AXES[axis];
        var matched = false;
        for (var i = 0; i < keywords.length; i++) {
          if (text.indexOf(keywords[i]) !== -1) {
            matched = true;
            break;
          }
        }
        if (matched) {
          axisCounts[axis]++;
          // Determine positive/negative from good/bad text or score
          var inGood = false, inBad = false;
          for (var j = 0; j < keywords.length; j++) {
            if (goodText && goodText.indexOf(keywords[j]) !== -1) inGood = true;
            if (badText && badText.indexOf(keywords[j]) !== -1) inBad = true;
          }
          if (inGood && !inBad) axisScores[axis].positive++;
          else if (inBad && !inGood) axisScores[axis].negative++;
          else if (score >= 8) axisScores[axis].positive++;
          else if (score <= 5) axisScores[axis].negative++;
          axisScores[axis].total++;
        }
      });
    });

    var nps = total > 0 ? Math.round((promoters / total - detractors / total) * 100) : 0;

    results[key] = {
      nps: nps,
      promoters: promoters,
      detractors: detractors,
      passives: total - promoters - detractors,
      totalReviews: total,
      axisScores: axisScores,
      axisCounts: axisCounts
    };
  });

  return results;
}

// Get keyword frequency from all reviews across hotels
function getKeywordFrequency(hotelDetails) {
  var freq = {};

  Object.keys(hotelDetails).forEach(function(key) {
    var reviews = hotelDetails[key].comments || hotelDetails[key].reviews || [];
    reviews.forEach(function(r) {
      var text = (r.translated || r.comment || r.c || '') + ' ' + (r.translated_good || r.good || r.g || '') + ' ' + (r.translated_bad || r.bad || r.b || '');
      Object.keys(CS_AXES).forEach(function(axis) {
        CS_AXES[axis].forEach(function(kw) {
          if (text.indexOf(kw) !== -1) {
            if (!freq[kw]) freq[kw] = { keyword: kw, axis: axis, count: 0 };
            freq[kw].count++;
          }
        });
      });
    });
  });

  return Object.values(freq).sort(function(a, b) { return b.count - a.count; });
}

module.exports = { analyzeCS, getKeywordFrequency, CS_AXES };
