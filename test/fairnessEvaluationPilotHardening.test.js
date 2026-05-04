const test = require('node:test');
const assert = require('node:assert/strict');

global.window = global;
global.localStorage = { getItem() { return null; }, setItem() {} };

require('../src/utils/loanCategoryUtils.js');
require('../src/services/fairnessEngineService.js');
require('../src/services/fairnessReviewWorkflowService.js');

function evaluateGlobal(officerStats) {
  return global.FairnessEngineService.evaluateFairness({
    engineType: 'global',
    officers: officerStats.map((entry) => ({ name: entry.officer, eligibility: { consumer: true, mortgage: false } })),
    officerStats
  });
}

function evaluateOfficerLane({ officers, officerStats, optimizationMetrics = {} }) {
  return global.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers,
    officerStats,
    optimizationMetrics
  });
}

test('calculateMaxSpread returns highest minus lowest and tolerates empty entries', () => {
  assert.equal(global.FairnessEngineService.calculateMaxSpread([], 'totalLoans'), 0);
  assert.equal(global.FairnessEngineService.calculateMaxSpread([
    { totalLoans: 2 },
    { totalLoans: 1 },
    { totalLoans: 0 }
  ], 'totalLoans'), 2);
});

test('global one-loan count spread does not REVIEW solely due to count variance', () => {
  const evaluation = evaluateGlobal([
    { officer: 'A', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
    { officer: 'B', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
    { officer: 'C', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} }
  ]);

  assert.equal(evaluation.metrics.maxCountVariancePercent, 50);
  assert.equal(evaluation.overallResult, 'PASS');
  assert.equal(evaluation.roleAwareFlags.globalOneLoanSpreadToleranceApplied, true);
  assert.match(evaluation.notes.join(' '), /one-loan spread tolerance/i);
});

test('global 2/1/1 count spread does not REVIEW solely due to count variance', () => {
  const evaluation = evaluateGlobal([
    { officer: 'A', totalLoans: 2, totalAmount: 100, consumerLoanCount: 2, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 2 } },
    { officer: 'B', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
    { officer: 'C', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } }
  ]);

  assert.equal(evaluation.overallResult, 'PASS');
  assert.equal(evaluation.roleAwareFlags.globalOneLoanSpreadToleranceApplied, true);
});

test('global count spread above one loan remains eligible for REVIEW', () => {
  const evaluation = evaluateGlobal([
    { officer: 'A', totalLoans: 3, totalAmount: 100, consumerLoanCount: 3, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 3 } },
    { officer: 'B', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} },
    { officer: 'C', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} }
  ]);

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'global_count_variance');
});

test('global high dollar variance still REVIEWs when count spread tolerance applies', () => {
  const evaluation = evaluateGlobal([
    { officer: 'A', totalLoans: 1, totalAmount: 1000, consumerLoanCount: 1, consumerAmount: 1000, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
    { officer: 'B', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
    { officer: 'C', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} }
  ]);

  assert.equal(evaluation.roleAwareFlags.globalOneLoanSpreadToleranceApplied, true);
  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'global_dollar_variance');
});

test('officer-lane consumer, mortgage, and flex one-loan spreads do not fail solely from count variance', () => {
  const consumerEvaluation = evaluateOfficerLane({
    officers: [
      { name: 'C1', eligibility: { consumer: true, mortgage: false } },
      { name: 'C2', eligibility: { consumer: true, mortgage: false } },
      { name: 'C3', eligibility: { consumer: true, mortgage: false } }
    ],
    officerStats: [
      { officer: 'C1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
      { officer: 'C2', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
      { officer: 'C3', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} }
    ]
  });
  assert.equal(consumerEvaluation.overallResult, 'PASS');
  assert.equal(consumerEvaluation.roleAwareFlags.consumerOneLoanSpreadToleranceApplied, true);

  const mortgageEvaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'M2', eligibility: { consumer: false, mortgage: true } },
      { name: 'M3', eligibility: { consumer: false, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'M2', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'M3', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 100, typeBreakdown: {} }
    ]
  });
  assert.equal(mortgageEvaluation.overallResult, 'PASS');
  assert.equal(mortgageEvaluation.roleAwareFlags.mortgageOneLoanSpreadToleranceApplied, true);

  const flexEvaluation = evaluateOfficerLane({
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } },
      { name: 'F3', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
      { officer: 'F2', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F3', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 50, mortgageLoanCount: 0, mortgageAmount: 50, typeBreakdown: {} }
    ]
  });
  assert.equal(flexEvaluation.overallResult, 'PASS');
  assert.equal(flexEvaluation.roleAwareFlags.flexOneLoanSpreadToleranceApplied, true);
});

test('minimal flex participation does not cause REVIEW solely from flex lane count variance', () => {
  const evaluation = evaluateOfficerLane({
    officers: [
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 1, consumerAmount: 100, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: { Personal: 1 } },
      { officer: 'F2', totalLoans: 0, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 50, mortgageLoanCount: 0, mortgageAmount: 50, typeBreakdown: {} }
    ]
  });

  assert.equal(evaluation.overallResult, 'PASS');
  assert.equal(evaluation.roleAwareFlags.flexMinimalParticipationToleranceApplied, true);
  assert.match(evaluation.notes.join(' '), /below the minimum volume/i);
});

test('flex mortgage policy violations still REVIEW', () => {
  const evaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { 'First Mortgage': 1 } }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');
});

test('HELOC-only missing weighted metric falls back without automatic REVIEW', () => {
  const evaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } }
    ]
  });

  assert.equal(evaluation.metrics.helocWeightedVariancePercent, null);
  assert.equal(evaluation.roleAwareFlags.helocWeightedMetricUnavailable, true);
  assert.equal(evaluation.overallResult, 'PASS');
  assert.match(evaluation.notes.join(' '), /Weighted HELOC optimization metric was unavailable/i);
});

test('HELOC-only support still REVIEWs when leadership policy fails', () => {
  const evaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F1', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'F2', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_leadership_policy');
});

test('HELOC-only policy failure descriptor takes priority over passing weighted metric', () => {
  const evaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } },
      { name: 'F2', eligibility: { consumer: true, mortgage: true } }
    ],
    optimizationMetrics: { helocWeightedVariancePercent: 8 },
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F1', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } },
      { officer: 'F2', totalLoans: 2, totalAmount: 200, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 2, mortgageAmount: 200, typeBreakdown: { HELOC: 2 } }
    ]
  });

  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.metrics.helocWeightedVariancePercent, 8);
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_leadership_policy');
});

test('ADVISORY does not require manager confirmation and ranks between PASS and REVIEW', () => {
  assert.equal(global.FairnessReviewService.shouldRequireManagerConfirmation({ status: 'ADVISORY' }), false);
  assert.equal(global.FairnessReviewService.shouldRequireManagerConfirmation({ status: 'REVIEW' }), true);

  const { selectedAttempt } = global.FairnessReviewService.selectBestFairnessAttempt([
    { attemptNumber: 1, status: 'REVIEW', metrics: { maxAmountVariancePercent: 1 } },
    { attemptNumber: 2, status: 'ADVISORY', metrics: { maxAmountVariancePercent: 24 } },
    { attemptNumber: 3, status: 'PASS', metrics: { maxAmountVariancePercent: 30 } }
  ]);
  assert.equal(selectedAttempt.attemptNumber, 3);
});

test('single-MLO expected variance is protected while true policy failures still REVIEW', () => {
  const protectedEvaluation = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 3, totalAmount: 300, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 3, mortgageAmount: 300, typeBreakdown: { HELOC: 3 } },
      { officer: 'F1', totalLoans: 0, totalAmount: 0, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 0, mortgageAmount: 0, typeBreakdown: {} }
    ]
  });

  assert.notEqual(protectedEvaluation.statusMetricDescriptor?.key, 'mortgage_lane_count_variance');
  assert.equal(protectedEvaluation.roleAwareFlags.hasSingleMortgageOnlyOfficer, true);
  assert.equal(protectedEvaluation.roleAwareFlags.mortgageVarianceExpected, true);
  assert.match(protectedEvaluation.chartAnnotations.mortgageNote, /one-MLO/i);

  const policyFailure = evaluateOfficerLane({
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { HELOC: 1 } },
      { officer: 'F1', totalLoans: 1, totalAmount: 100, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100, typeBreakdown: { 'First Mortgage': 1 } }
    ]
  });

  assert.equal(policyFailure.overallResult, 'REVIEW');
  assert.equal(policyFailure.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');
});
