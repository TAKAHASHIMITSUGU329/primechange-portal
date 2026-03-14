// V2 Revenue Calculator - opportunity loss calculations
// Formula: (targetScore - currentScore) * RevPAR_slope * roomCount * 30 days

function calcRevenueLoss(currentScore, targetScore, revparSlope, roomCount) {
  if (!currentScore || !targetScore || !revparSlope || !roomCount) return 0;
  var gap = targetScore - currentScore;
  if (gap <= 0) return 0;
  // RevPAR change per score point * room count * 30 days
  return Math.round(gap * revparSlope * roomCount * 30);
}

// Calculate monthly improvement opportunity for a hotel
function calcMonthlyOpportunity(hotelRevData, currentScore, targetScore, revparSlope) {
  if (!hotelRevData || !currentScore || !targetScore || !revparSlope) return 0;
  var gap = targetScore - currentScore;
  if (gap <= 0) return 0;
  return Math.round(gap * revparSlope * (hotelRevData.room_count || 0) * 30);
}

// Calculate revenue opportunities for all hotels
function calcAllRevenueOpportunities(data) {
  var opportunities = {};
  var hotelsRanked = data.pov && data.pov.hotels_ranked ? data.pov.hotels_ranked : [];
  var portfolioTargetScore = 8.89; // Default target

  if (data.kpiTargets && data.kpiTargets['ポートフォリオ平均スコア']) {
    portfolioTargetScore = data.kpiTargets['ポートフォリオ平均スコア'].target || 8.89;
  }

  hotelsRanked.forEach(function(h) {
    var key = h.key;
    var revData = data.revenueData[key];
    var currentScore = h.avg || h.avg_score || 0;

    // Use per-hotel target if available, otherwise portfolio target
    var targetScore = portfolioTargetScore;
    var perHotelTarget = findPerHotelTarget(data.perHotelTargets, h.name);
    if (perHotelTarget) {
      targetScore = perHotelTarget.target_avg || targetScore;
    }

    var monthlyLoss = 0;
    if (revData) {
      monthlyLoss = calcRevenueLoss(currentScore, targetScore, data.revparSlope, revData.room_count);
    }

    opportunities[key] = {
      name: h.name,
      key: key,
      currentScore: currentScore,
      targetScore: targetScore,
      gap: Math.round((targetScore - currentScore) * 100) / 100,
      monthlyLoss: monthlyLoss,
      roomCount: revData ? revData.room_count : 0,
      actualRevenue: revData ? revData.actual_revenue : 0,
      occupancyRate: revData ? revData.occupancy_rate : 0,
      priority: perHotelTarget ? perHotelTarget.priority : null
    };
  });

  return opportunities;
}

// Find per-hotel target by name
function findPerHotelTarget(targets, hotelName) {
  if (!targets) return null;
  var keys = Object.keys(targets);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i] === hotelName) return targets[keys[i]];
  }
  return null;
}

// Format yen amount for display
function formatYen(amount) {
  if (Math.abs(amount) >= 10000) {
    return Math.round(amount / 10000) + '万';
  }
  return amount.toLocaleString();
}

module.exports = { calcRevenueLoss, calcMonthlyOpportunity, calcAllRevenueOpportunities, formatYen };
