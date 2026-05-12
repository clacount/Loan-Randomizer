const test = require('node:test');
const assert = require('node:assert/strict');

global.window = global;
global.localStorage = {
  getItem() { return null; },
  setItem() {}
};

require('../src/utils/loanCategoryUtils.js');
require('../src/services/fairnessEngineService.js');

test('fairness engine classes are registered behind the service facade', () => {
  assert.equal(typeof global.GlobalFairnessEngine, 'function');
  assert.equal(typeof global.OfficerLaneFairnessEngine, 'function');
  assert.equal(typeof global.FairnessEngineService.getFairnessEngineEvaluator, 'undefined');
});

test('officer_lane does not fail routing when no mortgage-only lane exists', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'F1',
        totalLoans: 10,
        totalAmount: 100,
        consumerLoanCount: 10,
        consumerAmount: 100,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: {}
      },
      {
        officer: 'F2',
        totalLoans: 10,
        totalAmount: 100,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 10,
        mortgageAmount: 100,
        typeBreakdown: {}
      }
    ]
  });

  assert.equal(evaluation.metrics.flexVariance.maxCountVariancePercent, 0);
  assert.equal(evaluation.metrics.flexVariance.maxAmountVariancePercent, 0);
  assert.equal(evaluation.overallResult, 'PASS');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'flex_lane_dollar_variance');
});

test('officer_lane consumer variance includes flex officers that receive consumer loans', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'Ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'Peter', eligibility: { consumer: true, mortgage: true } },
      { name: 'Jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'Kash', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'Ashley',
        totalLoans: 3,
        totalAmount: 18300,
        consumerLoanCount: 3,
        consumerAmount: 18300,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Auto: 3 }
      },
      {
        officer: 'Peter',
        totalLoans: 8,
        totalAmount: 65600,
        consumerLoanCount: 8,
        consumerAmount: 65600,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Auto: 8 }
      },
      {
        officer: 'Jade',
        totalLoans: 2,
        totalAmount: 16100,
        consumerLoanCount: 2,
        consumerAmount: 16100,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Personal: 2 }
      },
      {
        officer: 'Kash',
        totalLoans: 5,
        totalAmount: 250000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 5,
        mortgageAmount: 250000,
        typeBreakdown: { HELOC: 5 }
      }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.ok(
    ['consumer_lane_count_variance', 'consumer_lane_dollar_variance'].includes(evaluation.statusMetricDescriptor?.key),
    `Expected consumer lane variance descriptor, got ${evaluation.statusMetricDescriptor?.key}`
  );
  assert.ok(evaluation.metrics.consumerVariance.maxAmountVariancePercent > 20);
  assert.match(evaluation.summaryItems.join(' | '), /Ashley/i);
  assert.match(evaluation.summaryItems.join(' | '), /Peter/i);
  assert.match(evaluation.summaryItems.join(' | '), /Jade/i);
});

test('officer_lane does not let flex-only imbalance force REVIEW when consumer and mortgage lanes pass', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: false,
    officers: [
      { name: 'Ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'Peter', eligibility: { consumer: true, mortgage: true } },
      { name: 'Jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'Kash', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'Peter',
        totalLoans: 3,
        totalAmount: 25700,
        consumerLoanCount: 3,
        consumerAmount: 25700,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { 'Credit Card': 1, Personal: 1, Auto: 1 }
      },
      {
        officer: 'Jade',
        totalLoans: 3,
        totalAmount: 71000,
        consumerLoanCount: 2,
        consumerAmount: 25000,
        mortgageLoanCount: 1,
        mortgageAmount: 46000,
        typeBreakdown: { Collateralized: 1, 'Credit Card': 1, HELOC: 1 }
      },
      {
        officer: 'Ashley',
        totalLoans: 3,
        totalAmount: 23800,
        consumerLoanCount: 3,
        consumerAmount: 23800,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Collateralized: 1, Personal: 2 }
      },
      {
        officer: 'Kash',
        totalLoans: 1,
        totalAmount: 38500,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 1,
        mortgageAmount: 38500,
        typeBreakdown: { HELOC: 1 }
      }
    ]
  });

  assert.equal(evaluation.metrics.consumerVariance.maxCountVariancePercent <= 15, true);
  assert.equal(evaluation.metrics.consumerVariance.maxAmountVariancePercent <= 20, true);
  assert.equal(evaluation.metrics.mortgageVariance.maxCountVariancePercent <= 15, true);
  assert.equal(evaluation.metrics.mortgageVariance.maxAmountVariancePercent <= 20, true);
  assert.equal(evaluation.metrics.flexVariance.maxAmountVariancePercent > 25, true);
  assert.equal(evaluation.overallResult, 'ADVISORY');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'flex_lane_dollar_variance');
  assert.equal(evaluation.statusMetricDescriptor?.contextLabel, 'Flex support monitoring');
  assert.equal(evaluation.roleAwareFlags?.flexVarianceNonBlocking, true);
  assert.match(evaluation.notes.join(' '), /does not independently force REVIEW/i);
});

test('officer_lane ignores flex officers with no consumer participation when evaluating consumer lane variance', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: false,
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: true, mortgage: true } },
      { name: 'peter', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 3,
        totalAmount: 23800,
        consumerLoanCount: 3,
        consumerAmount: 23800,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Auto: 1, Personal: 2 }
      },
      {
        officer: 'jade',
        totalLoans: 5,
        totalAmount: 137200,
        consumerLoanCount: 3,
        consumerAmount: 25700,
        mortgageLoanCount: 2,
        mortgageAmount: 111500,
        typeBreakdown: { Personal: 1, Collateralized: 2, HELOC: 2 }
      },
      {
        officer: 'kash',
        totalLoans: 2,
        totalAmount: 90500,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 2,
        mortgageAmount: 90500,
        typeBreakdown: { HELOC: 2 }
      },
      {
        officer: 'peter',
        totalLoans: 4,
        totalAmount: 117000,
        consumerLoanCount: 2,
        consumerAmount: 47333.33,
        mortgageLoanCount: 2,
        mortgageAmount: 69666.67,
        typeBreakdown: { 'Credit Card': 1, Personal: 1, HELOC: 2 }
      }
    ]
  });

  assert.equal(evaluation.metrics.consumerVariance.maxCountVariancePercent, 12.5);
  assert.equal(Number(evaluation.metrics.consumerVariance.maxAmountVariancePercent.toFixed(1)), 24.3);
  assert.equal(evaluation.overallResult, 'ADVISORY');
  assert.ok(
    ['flex_lane_count_variance', 'flex_lane_dollar_variance'].includes(evaluation.statusMetricDescriptor?.key),
    `Expected flex variance descriptor, got ${evaluation.statusMetricDescriptor?.key}`
  );
});

test('officer_lane ignores mortgage-lane review pressure during consumer-only runs', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: true,
    currentRunHasMortgageLoans: false,
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: false, mortgage: true } },
      { name: 'peter', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 3,
        totalAmount: 24000,
        consumerLoanCount: 3,
        consumerAmount: 24000,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Auto: 2, Personal: 1 }
      },
      {
        officer: 'jade',
        totalLoans: 3,
        totalAmount: 26000,
        consumerLoanCount: 3,
        consumerAmount: 26000,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Collateralized: 1, Personal: 2 }
      },
      {
        officer: 'kash',
        totalLoans: 4,
        totalAmount: 420000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 4,
        mortgageAmount: 420000,
        typeBreakdown: { 'First Mortgage': 4 }
      },
      {
        officer: 'peter',
        totalLoans: 1,
        totalAmount: 90000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 1,
        mortgageAmount: 90000,
        typeBreakdown: { HELOC: 1 }
      }
    ]
  });

  assert.equal(evaluation.metrics.consumerVariance.maxCountVariancePercent <= 15, true);
  assert.equal(evaluation.metrics.consumerVariance.maxAmountVariancePercent <= 20, true);
  assert.equal(evaluation.metrics.mortgageVariance.maxCountVariancePercent, 0);
  assert.equal(evaluation.metrics.mortgageVariance.maxAmountVariancePercent, 0);
  assert.equal(evaluation.overallResult, 'PASS');
  assert.notEqual(evaluation.statusMetricDescriptor?.key, 'mortgage_lane_dollar_variance');
});

test('officer_lane ignores consumer-lane review pressure during mortgage-only runs', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: false,
    currentRunHasMortgageLoans: true,
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: false, mortgage: true } },
      { name: 'peter', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 3,
        totalAmount: 35900,
        consumerLoanCount: 3,
        consumerAmount: 35900,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Personal: 2, Auto: 1 }
      },
      {
        officer: 'jade',
        totalLoans: 3,
        totalAmount: 13600,
        consumerLoanCount: 3,
        consumerAmount: 13600,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Personal: 3 }
      },
      {
        officer: 'kash',
        totalLoans: 2,
        totalAmount: 622500,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 2,
        mortgageAmount: 622500,
        typeBreakdown: { 'First Mortgage': 1, 'Home Refi': 1 }
      },
      {
        officer: 'peter',
        totalLoans: 2,
        totalAmount: 71000,
        consumerLoanCount: 2,
        consumerAmount: 71000,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Collateralized: 1, Personal: 1 }
      }
    ]
  });

  assert.equal(Number(evaluation.metrics.consumerVariance.maxAmountVariancePercent.toFixed(1)), 47.6);
  assert.equal(evaluation.metrics.mortgageVariance.maxCountVariancePercent, 0);
  assert.equal(evaluation.metrics.mortgageVariance.maxAmountVariancePercent, 0);
  assert.notEqual(evaluation.overallResult, 'REVIEW');
  assert.notEqual(evaluation.statusMetricDescriptor?.key, 'consumer_lane_dollar_variance');
});

test('officer_lane excludes flex officers without current-run mortgage participation from low-volume mortgage policy checks', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: true,
    currentRunHasMortgageLoans: true,
    currentRunLaneParticipationByOfficer: {
      ashley: { consumerLoanCount: 1, consumerAmount: 9000, mortgageLoanCount: 0, mortgageAmount: 0, totalLoans: 1, totalAmount: 9000 },
      jade: { consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 0, totalLoans: 0, totalAmount: 0 },
      kash: { consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 450000, totalLoans: 1, totalAmount: 450000 }
    },
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 3,
        totalAmount: 27000,
        consumerLoanCount: 3,
        consumerAmount: 27000,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Personal: 3 }
      },
      {
        officer: 'jade',
        totalLoans: 5,
        totalAmount: 410000,
        consumerLoanCount: 3,
        consumerAmount: 30000,
        mortgageLoanCount: 2,
        mortgageAmount: 380000,
        typeBreakdown: { Personal: 3, 'First Mortgage': 2 }
      },
      {
        officer: 'kash',
        totalLoans: 4,
        totalAmount: 1560000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 4,
        mortgageAmount: 1560000,
        typeBreakdown: { 'First Mortgage': 4 }
      }
    ]
  });

  assert.equal(evaluation.overallResult, 'PASS');
  assert.notEqual(evaluation.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');
  assert.equal(evaluation.roleAwareFlags?.lowVolumeMortgageFlexParticipationFilterApplied, true);
});

test('officer_lane treats one-loan deviation from average as acceptable count balance when dollar variance is the primary issue', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: true,
    currentRunHasMortgageLoans: true,
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'peter', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 4,
        totalAmount: 109800,
        consumerLoanCount: 4,
        consumerAmount: 109800,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Collateralized: 4 }
      },
      {
        officer: 'jade',
        totalLoans: 4,
        totalAmount: 194000,
        consumerLoanCount: 2,
        consumerAmount: 54000,
        mortgageLoanCount: 2,
        mortgageAmount: 140000,
        typeBreakdown: { Personal: 2, 'First Mortgage': 1, HELOC: 1 }
      },
      {
        officer: 'peter',
        totalLoans: 5,
        totalAmount: 131700,
        consumerLoanCount: 3,
        consumerAmount: 79020,
        mortgageLoanCount: 2,
        mortgageAmount: 52680,
        typeBreakdown: { Collateralized: 1, Personal: 4 }
      },
      {
        officer: 'kash',
        totalLoans: 5,
        totalAmount: 660000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 5,
        mortgageAmount: 660000,
        typeBreakdown: { 'First Mortgage': 3, HELOC: 2 }
      }
    ]
  });

  assert.equal(evaluation.metrics.consumerVariance.maxCountVariancePercent > 15, true);
  assert.equal(evaluation.metrics.consumerVariance.oneLoanTargetDeviationToleranceApplied, true);
  assert.equal(evaluation.metrics.consumerVariance.countDistributionPass, true);
  assert.equal(Number(evaluation.metrics.consumerVariance.maxAmountVariancePercent.toFixed(1)), 23.0);
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');
});

test('homogeneous HELOC support pool can PASS with specialized thresholds and weighted optimization target', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } },
      { name: 'F3', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    optimizationMetrics: {
      helocWeightedVariancePercent: 12.6
    },
    officerStats: [
      { officer: 'F1', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'F2', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F3', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'M1', totalLoans: 3, totalAmount: 300, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 3, mortgageAmount: 300, typeBreakdown: { HELOC: 3 } }
    ]
  });

  assert.ok(evaluation.metrics.flexVariance.maxCountVariancePercent > 15, 'Would fail old generic 15% count threshold.');
  assert.equal(evaluation.overallResult, 'PASS');
  assert.equal(evaluation.roleAwareFlags.helocOnlySupportThresholdsApplied, true);
  assert.equal(evaluation.metrics.helocWeightedVariancePercent, 12.6);
  assert.equal(evaluation.statusMetricDescriptor?.key, 'heloc_weighted_variance');
  assert.equal(evaluation.statusMetricDescriptor?.label, 'Weighted HELOC variance');
  assert.equal(evaluation.statusMetricDescriptor?.valuePercent, 12.6);
  assert.match(evaluation.summaryItems.join(' | '), /HELOC-only support thresholds applied/i);
});

test('fairness metrics keep HELOC weighted variance null when optimization metric is absent', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } }
    ]
  });

  assert.equal(global.FairnessEngineService.isHomogeneousHelocSupportPool({
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    hasConsumerLoans: false,
    loanTypeNames: ['heloc']
  }), true);
  assert.equal(evaluation.roleAwareFlags.helocOnlySupportThresholdsApplied, true);
  assert.equal(evaluation.metrics.helocWeightedVariancePercent, null);
  assert.equal(evaluation.roleAwareFlags.helocWeightedMetricUnavailable, true);
  assert.notEqual(evaluation.statusMetricDescriptor?.key, 'heloc_weighted_variance');
  assert.match(evaluation.notes.join(' '), /Weighted HELOC optimization metric was unavailable/i);
});

test('HELOC support pool detection ignores mortgage-only officers excluded from HELOC assignment', () => {
  assert.equal(global.FairnessEngineService.isHomogeneousHelocSupportPool({
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true }, excludeHeloc: true }
    ],
    hasConsumerLoans: false,
    loanTypeNames: ['heloc']
  }), false);
});

test('homogeneous HELOC support pool remains REVIEW when flex variance exceeds 25%', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } },
      { name: 'F3', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    optimizationMetrics: {
      helocWeightedVariancePercent: 18
    },
    officerStats: [
      { officer: 'F1', totalLoans: 4, totalAmount: 400, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 4, mortgageAmount: 400, typeBreakdown: { HELOC: 4 } },
      { officer: 'F2', totalLoans: 0, totalAmount: 0, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} },
      { officer: 'F3', totalLoans: 0, totalAmount: 0, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} },
      { officer: 'M1', totalLoans: 4, totalAmount: 400, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 4, mortgageAmount: 400, typeBreakdown: { HELOC: 4 } }
    ]
  });

  assert.ok(evaluation.metrics.flexVariance.maxCountVariancePercent > 25);
  assert.equal(evaluation.overallResult, 'REVIEW');
});

test('homogeneous HELOC support pool remains REVIEW when weighted optimization target fails', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } },
      { name: 'F3', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    optimizationMetrics: {
      helocWeightedVariancePercent: 27
    },
    officerStats: [
      { officer: 'F1', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'F2', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F3', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'M1', totalLoans: 3, totalAmount: 300, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 3, mortgageAmount: 300, typeBreakdown: { HELOC: 3 } }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
});

test('homogeneous HELOC support pool remains REVIEW when M leadership policy fails', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } },
      { name: 'F3', eligibility: { consumer: true, mortgage: true } },
      { name: 'M1', eligibility: { consumer: false, mortgage: true } }
    ],
    optimizationMetrics: {
      helocWeightedVariancePercent: 12
    },
    officerStats: [
      { officer: 'F1', totalLoans: 3, totalAmount: 300, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 3, mortgageAmount: 300, typeBreakdown: { HELOC: 3 } },
      { officer: 'F2', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'F3', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
});

test('mortgage leadership descriptor value uses mortgage loan-count share, not mortgage dollar share', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'M1',
        totalLoans: 2,
        totalAmount: 900000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 2,
        mortgageAmount: 900000,
        typeBreakdown: { 'Home Refi': 2 }
      },
      {
        officer: 'F1',
        totalLoans: 3,
        totalAmount: 100000,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 3,
        mortgageAmount: 100000,
        typeBreakdown: { 'Home Refi': 3 }
      },
      {
        officer: 'F2',
        totalLoans: 0,
        totalAmount: 0,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: {}
      }
    ]
  });

  const expectedCountSharePercent = (2 / 5) * 100;
  const mortgageRoutingSharePercent = (900000 / 1000000) * 100;

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_leadership_policy');
  assert.equal(evaluation.statusMetricDescriptor?.valuePercent, expectedCountSharePercent);
  assert.notEqual(evaluation.statusMetricDescriptor?.valuePercent, mortgageRoutingSharePercent);
});

test('officer_lane excludes flex officers without current-run consumer participation in mixed runs', () => {
  const evaluation = global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    currentRunHasConsumerLoans: true,
    currentRunHasMortgageLoans: true,
    currentRunLaneParticipationByOfficer: {
      ashley: { consumerLoanCount: 1, consumerAmount: 86000, mortgageLoanCount: 0, mortgageAmount: 0, totalLoans: 1, totalAmount: 86000 },
      jade: { consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 0, totalLoans: 0, totalAmount: 0 },
      peter: { consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 115000, totalLoans: 2, totalAmount: 115000 },
      kash: { consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 450500, totalLoans: 1, totalAmount: 450500 }
    },
    officers: [
      { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
      { name: 'jade', eligibility: { consumer: true, mortgage: true } },
      { name: 'peter', eligibility: { consumer: true, mortgage: true } },
      { name: 'kash', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      {
        officer: 'ashley',
        totalLoans: 4,
        totalAmount: 109800,
        consumerLoanCount: 4,
        consumerAmount: 109800,
        mortgageLoanCount: 0,
        mortgageAmount: 0,
        typeBreakdown: { Collateralized: 4 }
      },
      {
        officer: 'jade',
        totalLoans: 4,
        totalAmount: 194000,
        consumerLoanCount: 2,
        consumerAmount: 54000,
        mortgageLoanCount: 2,
        mortgageAmount: 140000,
        typeBreakdown: { Personal: 2, 'First Mortgage': 1, HELOC: 1 }
      },
      {
        officer: 'peter',
        totalLoans: 7,
        totalAmount: 246700,
        consumerLoanCount: 5,
        consumerAmount: 131700,
        mortgageLoanCount: 2,
        mortgageAmount: 115000,
        typeBreakdown: { Collateralized: 1, Personal: 4, HELOC: 2 }
      },
      {
        officer: 'kash',
        totalLoans: 6,
        totalAmount: 1110500,
        consumerLoanCount: 0,
        consumerAmount: 0,
        mortgageLoanCount: 6,
        mortgageAmount: 1110500,
        typeBreakdown: { 'First Mortgage': 4, HELOC: 2 }
      }
    ]
  });

  assert.equal(evaluation.metrics.consumerVariance.maxCountVariancePercent, 0);
  assert.equal(evaluation.metrics.consumerVariance.maxAmountVariancePercent, 0);
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');
});
