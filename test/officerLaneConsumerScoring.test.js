const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function makeElement(tagName = 'div') {
  return {
    tagName,
    style: {},
    dataset: {},
    className: '',
    classList: { add() {}, remove() {}, toggle() { return false; }, contains() { return false; } },
    children: [],
    textContent: '',
    innerHTML: '',
    appendChild(child) { this.children.push(child); return child; },
    append() {},
    replaceChildren(...children) { this.children = [...children]; },
    remove() {},
    setAttribute() {},
    removeAttribute() {},
    addEventListener() {},
    removeEventListener() {},
    querySelector() { return makeElement(); },
    querySelectorAll() { return []; },
    closest() { return null; },
    getContext() {
      return {
        clearRect() {},
        fillRect() {},
        fillText() {},
        beginPath() {},
        moveTo() {},
        arc() {},
        closePath() {},
        fill() {},
        measureText() { return { width: 100 }; },
        textAlign: 'left',
        font: '',
        fillStyle: ''
      };
    },
    toDataURL() { return 'data:image/png;base64,stub'; },
    focus() {},
    click() {},
    value: '',
    checked: false,
    disabled: false,
    open: false
  };
}

function loadAppContext() {
  const localStorageStore = {};
  const elementsById = {};
  const documentStub = {
    getElementById(id) {
      if (!elementsById[id]) {
        elementsById[id] = makeElement('div');
      }
      return elementsById[id];
    },
    querySelector() { return makeElement(); },
    querySelectorAll() { return []; },
    createElement(tagName) { return makeElement(tagName); },
    addEventListener() {},
    body: makeElement('body'),
    documentElement: { dataset: { theme: 'dark' } }
  };

  const context = {
    console,
    window: null,
    document: documentStub,
    localStorage: {
      getItem(key) { return localStorageStore[key] ?? null; },
      setItem(key, value) { localStorageStore[key] = String(value); }
    },
    setTimeout,
    clearTimeout,
    Math,
    Date,
    JSON,
    Map,
    Set,
    Object,
    Array,
    Number,
    String,
    Boolean,
    Promise,
    Error,
    parseFloat,
    parseInt,
    isNaN,
    Intl,
    location: { hostname: 'localhost', protocol: 'file:' },
    confirm: () => false,
    prompt: () => null,
    alert: () => {}
  };

  context.window = context;
  context.global = context;

  vm.createContext(context);

  const root = path.resolve(__dirname, '..');
  [
    'src/utils/loanCategoryUtils.js',
    'src/services/fairnessEngines/globalFairnessEngine.js',
    'src/services/fairnessEngines/officerLaneFairnessEngine.js',
    'src/services/fairnessEngineService.js',
    'src/services/focusWeightSettingsService.js',
    'src/services/mortgageFocusRoutingService.js',
    'src/services/officerLaneOptimizationService.js',
    'src/services/linkedLoanGroupService.js',
    'src/bootstrap/initApp.js'
  ].forEach((relativePath) => {
    const source = fs.readFileSync(path.join(root, relativePath), 'utf8');
    vm.runInContext(source, context, { filename: relativePath });
  });

  return context;
}

test('officer-lane consumer scoring prefers consumer-only catch-up over flex subtype flattening in segmented lanes', () => {
  const context = loadAppContext();
  context.FairnessEngineService.setSelectedFairnessEngine('officer_lane');

  const officersByName = {
    ashley: { name: 'ashley', eligibility: { consumer: true, mortgage: false }, weights: { consumer: 1, mortgage: 0 } },
    jade: { name: 'jade', eligibility: { consumer: true, mortgage: true }, weights: { consumer: 0.7, mortgage: 0.3 } },
    peter: { name: 'peter', eligibility: { consumer: true, mortgage: true }, weights: { consumer: 0.5, mortgage: 0.5 } },
    kash: { name: 'kash', eligibility: { consumer: false, mortgage: true }, weights: { consumer: 0, mortgage: 1 } }
  };
  const officerLoanTotals = { ashley: 3, jade: 4, peter: 5, kash: 2 };
  const officerAmountTotals = { ashley: 23800, jade: 108000, peter: 131700, kash: 94500 };
  const officerActiveSessions = { ashley: 1, jade: 1, peter: 1, kash: 1 };
  const runAssignmentCounts = { ashley: 0, jade: 0, peter: 0, kash: 0 };
  const runTypeAssignmentCounts = { ashley: {}, jade: {}, peter: {}, kash: {} };
  const officerTypeCounts = {
    ashley: { Collateralized: 3 },
    jade: { Personal: 2, 'First Mortgage': 1, HELOC: 1 },
    peter: { Collateralized: 1, Personal: 4 },
    kash: { 'First Mortgage': 1, HELOC: 1 }
  };

  const decision = context.chooseOfficerForLoan(
    officersByName,
    officerLoanTotals,
    officerTypeCounts,
    officerAmountTotals,
    officerActiveSessions,
    runAssignmentCounts,
    runTypeAssignmentCounts,
    {},
    { name: 'app4', type: 'Collateralized', amountRequested: 86000 }
  );

  assert.equal(decision.selectedOfficer, 'ashley');
  assert.equal(decision.scoredOfficers[0].officer, 'ashley');
  assert.equal(decision.scoredOfficers[0].segmentedConsumerLane, true);
});

test('officer-lane final assignment keeps flex officers out of consumer-lane variance when they received no consumer loans in the run', () => {
  const context = loadAppContext();
  context.Math.random = () => 0;
  context.FairnessEngineService.setSelectedFairnessEngine('officer_lane');

  const officers = [
    { name: 'ashley', eligibility: { consumer: true, mortgage: false }, weights: { consumer: 1, mortgage: 0 } },
    { name: 'jade', eligibility: { consumer: true, mortgage: true }, weights: { consumer: 0.7, mortgage: 0.3 } },
    { name: 'peter', eligibility: { consumer: true, mortgage: true }, weights: { consumer: 0.5, mortgage: 0.5 } },
    { name: 'kash', eligibility: { consumer: false, mortgage: true }, weights: { consumer: 0, mortgage: 1 } }
  ];
  const loans = [
    { name: 'app1', type: 'Collateralized', amountRequested: 86000 },
    { name: 'app2', type: 'HELOC', amountRequested: 62000 },
    { name: 'app3', type: 'HELOC', amountRequested: 53000 },
    { name: 'app4', type: 'First Mortgage', amountRequested: 450500 }
  ];
  const runningTotals = {
    officers: {
      ashley: {
        loanCount: 3,
        totalAmountRequested: 23800,
        typeCounts: { Collateralized: 3 },
        activeSessionCount: 0,
        eligibility: { consumer: true, mortgage: false },
        weights: { consumer: 1, mortgage: 0 }
      },
      jade: {
        loanCount: 4,
        totalAmountRequested: 194000,
        typeCounts: { Personal: 2, 'First Mortgage': 1, HELOC: 1 },
        activeSessionCount: 0,
        eligibility: { consumer: true, mortgage: true },
        weights: { consumer: 0.7, mortgage: 0.3 }
      },
      peter: {
        loanCount: 5,
        totalAmountRequested: 131700,
        typeCounts: { Collateralized: 1, Personal: 4 },
        activeSessionCount: 0,
        eligibility: { consumer: true, mortgage: true },
        weights: { consumer: 0.5, mortgage: 0.5 }
      },
      kash: {
        loanCount: 5,
        totalAmountRequested: 660000,
        typeCounts: { 'First Mortgage': 3, HELOC: 2 },
        activeSessionCount: 0,
        eligibility: { consumer: false, mortgage: true },
        weights: { consumer: 0, mortgage: 1 }
      }
    }
  };

  const result = context.assignLoans(officers, loans, runningTotals);
  const assignmentsByLoan = Object.fromEntries(result.loanAssignments.map((entry) => [entry.loan.name, entry.officers[0]]));

  assert.equal(assignmentsByLoan.app1, 'ashley');
  assert.equal(result.fairnessEvaluation?.metrics?.consumerVariance?.maxAmountVariancePercent, 0);
  assert.equal(result.fairnessEvaluation?.metrics?.consumerVariance?.maxCountVariancePercent, 0);
});

test('officer-lane downgrades tiny single-loan same-lane residual imbalance from REVIEW to ADVISORY', () => {
  const context = loadAppContext();
  context.Math.random = () => 0;
  context.FairnessEngineService.setSelectedFairnessEngine('officer_lane');

  const officers = [
    { name: 'ashley', eligibility: { consumer: true, mortgage: false }, weights: { consumer: 1, mortgage: 0 } },
    { name: 'jade', eligibility: { consumer: true, mortgage: false }, weights: { consumer: 1, mortgage: 0 } }
  ];
  const loans = [
    { name: 'app1', type: 'Personal', amountRequested: 6000 }
  ];
  const runningTotals = {
    officers: {
      ashley: {
        loanCount: 6,
        totalAmountRequested: 120000,
        typeCounts: { Personal: 6 },
        activeSessionCount: 0,
        eligibility: { consumer: true, mortgage: false },
        weights: { consumer: 1, mortgage: 0 }
      },
      jade: {
        loanCount: 4,
        totalAmountRequested: 60000,
        typeCounts: { Personal: 4 },
        activeSessionCount: 0,
        eligibility: { consumer: true, mortgage: false },
        weights: { consumer: 1, mortgage: 0 }
      }
    }
  };

  const result = context.assignLoans(officers, loans, runningTotals);
  const assignmentsByLoan = Object.fromEntries(result.loanAssignments.map((entry) => [entry.loan.name, entry.officers[0]]));

  assert.equal(assignmentsByLoan.app1, 'jade');
  assert.equal(result.fairnessEvaluation?.statusMetricDescriptor?.key, 'consumer_lane_dollar_variance');
  assert.equal(result.fairnessEvaluation?.overallResult, 'ADVISORY');
  assert.equal(result.fairnessEvaluation?.roleAwareFlags?.immaterialLaneImpactAdvisoryApplied, true);
  assert.equal(result.fairnessEvaluation?.roleAwareFlags?.immaterialLaneImpactLane, 'consumer');
});

test('optimization summary is suppressed when the final fairness descriptor changed', () => {
  const context = loadAppContext();

  const fairnessEvaluation = context.applyOptimizationSummaryToFairnessEvaluation(
    {
      optimizationApplied: true,
      optimizationSummaryMessage: 'Optimization reached the primary consumer-lane variance target band (<= 15.0%).',
      optimizationTargetDescriptorKey: 'consumer_lane_count_variance'
    },
    {
      overallResult: 'ADVISORY',
      statusMetricDescriptor: {
        key: 'consumer_lane_dollar_variance',
        label: 'Consumer lane dollar variance',
        valuePercent: 20.8
      },
      summaryItems: ['Consumer dollar variance advisory band applied: 20.0%–25.0%']
    }
  );

  assert.deepEqual(
    fairnessEvaluation.summaryItems,
    ['Consumer dollar variance advisory band applied: 20.0%–25.0%']
  );
});
