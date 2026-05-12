const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function makeElement(tagName = 'div') {
  const element = {
    tagName,
    style: {},
    dataset: {},
    className: '',
    classList: { add() {}, remove() {}, toggle() { return false; }, contains() { return false; } },
    children: [],
    textContent: '',
    innerHTML: '',
    append() {},
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    replaceChildren(...children) {
      this.children = [...children];
    },
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

  return new Proxy(element, {
    get(target, property) {
      return property in target ? target[property] : undefined;
    },
    set(target, property, value) {
      target[property] = value;
      return true;
    }
  });
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
    body: makeElement('body')
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
    location: { hostname: 'localhost' },
    confirm: () => false,
    prompt: () => null,
    alert: () => {}
  };

  context.window = context;
  context.global = context;
  context.__elementsById = elementsById;

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
    'src/bootstrap/initApp.js'
  ].forEach((relativePath) => {
    const source = fs.readFileSync(path.join(root, relativePath), 'utf8');
    vm.runInContext(source, context, { filename: relativePath });
  });

  return context;
}

test('role-based mortgage dollar donut includes flex mortgage participation and keeps consumer donut HELOC-excluded', () => {
  const context = loadAppContext();
  context.setSelectedFairnessEngine('officer_lane');

  const chartConfigs = [];
  context.DistributionChartRenderer = {
    drawDonutChart(config) {
      chartConfigs.push(config);
      return { canvas: makeElement('canvas'), imageDataUrl: 'data:image/png;base64,stub' };
    }
  };

  const officers = [
    { name: 'Jade', eligibility: { consumer: true, mortgage: true } },
    { name: 'Kash', eligibility: { consumer: false, mortgage: true } },
    { name: 'Casey', eligibility: { consumer: true, mortgage: false } }
  ];
  const runningTotals = { officers: {} };
  const result = {
    officerAssignments: {
      Jade: [{ type: 'HELOC', amountRequested: 46000 }],
      Kash: [{ type: 'HELOC', amountRequested: 38500 }],
      Casey: [{ type: 'Personal', amountRequested: 9000 }]
    },
    fairnessEvaluation: {
      chartAnnotations: {
        mortgageTitleSuffix: ' (Role-Based)',
        mortgageNote: 'Primary mortgage allocation to MLOs is expected.'
      },
      statusMetricDescriptor: {
        key: 'flex_lane_dollar_variance',
        label: 'Flex lane dollar variance',
        valuePercent: 11.7,
        contextLabel: 'Flex lane thresholds'
      }
    }
  };

  context.renderDistributionCharts(result, officers, runningTotals);

  const mortgageAfterConfig = chartConfigs.find((config) => config.title.startsWith('Mortgage Goal Dollars After Run'));
  assert.ok(mortgageAfterConfig, 'Mortgage role-based chart config should be rendered.');
  assert.deepEqual(
    Array.from(mortgageAfterConfig.distribution, (entry) => entry.officer),
    ['Jade', 'Kash'],
    'Mortgage role-based chart should include flex and mortgage-only officers with mortgage dollars.'
  );
  assert.equal(mortgageAfterConfig.distribution.find((entry) => entry.officer === 'Jade')?.totalAmountRequested, 46000);
  assert.equal(mortgageAfterConfig.distribution.find((entry) => entry.officer === 'Kash')?.totalAmountRequested, 38500);

  const consumerAfterConfig = chartConfigs.find((config) => config.title.startsWith('Consumer Goal Dollars After Run'));
  assert.ok(consumerAfterConfig, 'Consumer role-based chart config should be rendered.');
  assert.deepEqual(
    Array.from(consumerAfterConfig.distribution, (entry) => entry.officer),
    ['Jade', 'Casey'],
    'Consumer role-based chart should keep holistic consumer-lane history visible even when low-volume participation protection is active.'
  );

  const distributionCharts = context.__elementsById.distributionCharts;
  const consumerAfterCard = distributionCharts.children[3];
  const consumerNoteTexts = consumerAfterCard.children
    .filter((child) => child.tagName === 'p')
    .map((child) => child.textContent);
  assert.ok(consumerNoteTexts.some((text) => text.includes('Variance/status view: Consumer lane dollar variance')));
  assert.ok(consumerNoteTexts.some((text) => text.includes('Overall status driver: Flex lane dollar variance 11.7%')));
  assert.ok(consumerNoteTexts.some((text) => text.includes('Flex officers can appear in consumer composition')));

  const mortgageAfterCard = distributionCharts.children[5];
  const noteTexts = mortgageAfterCard.children
    .filter((child) => child.tagName === 'p')
    .map((child) => child.textContent);
  assert.ok(noteTexts.some((text) => text.includes('Distribution share view: this donut is composition only; fairness status is based on lane variance metrics.')));
  assert.ok(noteTexts.some((text) => text.includes('Officers with zero mortgage-category dollars are excluded from this mortgage lane chart.')));
});

test('role-based mortgage donut excludes flex officers without current-run mortgage participation in one-mortgage-loan runs', () => {
  const context = loadAppContext();
  context.setSelectedFairnessEngine('officer_lane');

  const chartConfigs = [];
  context.DistributionChartRenderer = {
    drawDonutChart(config) {
      chartConfigs.push(config);
      return { canvas: makeElement('canvas'), imageDataUrl: 'data:image/png;base64,stub' };
    }
  };

  const officers = [
    { name: 'Jade', eligibility: { consumer: true, mortgage: true } },
    { name: 'Kash', eligibility: { consumer: false, mortgage: true } },
    { name: 'Casey', eligibility: { consumer: true, mortgage: false } }
  ];
  const runningTotals = {
    officers: {
      Jade: { loanCount: 2, totalAmountRequested: 380000, typeCounts: { 'First Mortgage': 2 } },
      Kash: { loanCount: 3, totalAmountRequested: 1110000, typeCounts: { 'First Mortgage': 3 } },
      Casey: { loanCount: 2, totalAmountRequested: 18000, typeCounts: { Personal: 2 } }
    }
  };
  const result = {
    officerAssignments: {
      Jade: [],
      Kash: [{ type: 'First Mortgage', amountRequested: 450000 }],
      Casey: [{ type: 'Personal', amountRequested: 9000 }]
    },
    fairnessEvaluation: {
      chartAnnotations: {
        mortgageTitleSuffix: ' (Role-Based)',
        mortgageNote: 'Primary mortgage allocation to MLOs is expected.'
      },
      statusMetricDescriptor: {
        key: 'mortgage_lane_dollar_variance',
        label: 'Mortgage lane dollar variance',
        valuePercent: 0,
        contextLabel: 'Mortgage lane thresholds'
      }
    }
  };

  context.renderDistributionCharts(result, officers, runningTotals);

  const mortgageAfterConfig = chartConfigs.find((config) => config.title.startsWith('Mortgage Goal Dollars After Run'));
  assert.ok(mortgageAfterConfig, 'Mortgage role-based chart config should be rendered.');
  assert.deepEqual(
    Array.from(mortgageAfterConfig.distribution, (entry) => entry.officer),
    ['Jade', 'Kash'],
    'Mortgage role-based chart should keep holistic mortgage history visible even when low-volume mortgage protection is active.'
  );
});

test('mortgage flex participation descriptor is treated as mortgage-lane driven in parity notes', () => {
  const context = loadAppContext();
  const evaluation = context.FairnessEngineService.evaluateFairness({
    engineType: 'officer_lane',
    officers: [
      { name: 'M1', eligibility: { consumer: false, mortgage: true } },
      { name: 'F1', eligibility: { consumer: true, mortgage: true } }
    ],
    officerStats: [
      { officer: 'M1', totalLoans: 6, totalAmount: 600000, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 6, mortgageAmount: 600000, typeBreakdown: { 'Home Refi': 6 } },
      { officer: 'F1', totalLoans: 1, totalAmount: 100000, consumerLoanCount: 0, consumerAmount: 0, mortgageLoanCount: 1, mortgageAmount: 100000, typeBreakdown: { 'Home Refi': 1 } }
    ]
  });
  assert.equal(evaluation.overallResult, 'REVIEW');
  assert.equal(evaluation.statusMetricDescriptor?.key, 'mortgage_flex_participation_policy');

  const parityNotes = context.getOfficerLaneChartParityNotes({
    chartLane: 'mortgage',
    fairnessEvaluation: evaluation
  });

  assert.match(parityNotes.statusMetricLine, /Mortgage lane dollar variance/);
  assert.match(parityNotes.statusMetricLine, /Mortgage lane thresholds/);
  assert.match(parityNotes.alignmentLine, /mortgage policy check/i);
  assert.match(parityNotes.statusDriverLine, /Flex mortgage participation policy/);
  assert.match(parityNotes.statusDriverLine, /Mortgage lane policy checks/);
  assert.doesNotMatch(parityNotes.alignmentLine, /driven by flex-lane variance/i);
});

test('linked mixed-loan runs still render distribution charts when the custom renderer fails', () => {
  const context = loadAppContext();
  context.setSelectedFairnessEngine('officer_lane');

  let attemptedTitles = 0;
  context.DistributionChartRenderer = {
    drawDonutChart(config) {
      attemptedTitles += 1;
      if (config.title.startsWith('Consumer Goal Dollars After Run')) {
        throw new Error('Simulated renderer failure');
      }
      return { canvas: makeElement('canvas'), imageDataUrl: 'data:image/png;base64,stub' };
    }
  };

  const officers = [
    { name: 'ashley', eligibility: { consumer: true, mortgage: false } },
    { name: 'jade', eligibility: { consumer: true, mortgage: true } },
    { name: 'peter', eligibility: { consumer: true, mortgage: true } },
    { name: 'kash', eligibility: { consumer: false, mortgage: true } }
  ];
  const runningTotals = {
    officers: {
      ashley: { loanCount: 3, totalAmountRequested: 41300, typeCounts: { Collateralized: 3 } },
      jade: { loanCount: 1, totalAmountRequested: 38500, typeCounts: { HELOC: 1 } },
      peter: { loanCount: 2, totalAmountRequested: 46700, typeCounts: { Auto: 1, HELOC: 1 } },
      kash: { loanCount: 1, totalAmountRequested: 38500, typeCounts: { HELOC: 1 } }
    }
  };
  const result = {
    loanAssignments: [
      {
        loan: { name: 'APP-2026-009', type: 'Collateralized', amountRequested: 25700, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' },
        officers: ['jade'],
        unitType: 'linked_group'
      },
      {
        loan: { name: 'APP-2026-011', type: 'Credit Card', amountRequested: 0, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' },
        officers: ['jade'],
        unitType: 'linked_group'
      },
      {
        loan: { name: 'APP-2026-012', type: 'HELOC', amountRequested: 45300, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' },
        officers: ['jade'],
        unitType: 'linked_group'
      }
    ],
    officerAssignments: {
      ashley: [],
      jade: [
        { name: 'APP-2026-009', type: 'Collateralized', amountRequested: 25700, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' },
        { name: 'APP-2026-011', type: 'Credit Card', amountRequested: 0, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' },
        { name: 'APP-2026-012', type: 'HELOC', amountRequested: 45300, linkedGroupId: 'MLG-001', linkedGroupLabel: 'member 1234' }
      ],
      peter: [],
      kash: []
    },
    fairnessEvaluation: {
      overallResult: 'ADVISORY',
      metrics: {
        consumerVariance: { maxAmountVariancePercent: 16.1 },
        mortgageVariance: { maxAmountVariancePercent: 18.4 },
        flexVariance: { maxAmountVariancePercent: 8.9, totalConsumerAmount: 25700, totalMortgageAmount: 45300 }
      },
      statusMetricDescriptor: {
        key: 'consumer_lane_dollar_variance',
        label: 'Consumer lane dollar variance',
        valuePercent: 16.1,
        contextLabel: 'Consumer lane thresholds'
      },
      chartAnnotations: { mortgageTitleSuffix: ' (Role-Based)', mortgageNote: 'Primary mortgage allocation to MLOs is expected.' }
    }
  };

  assert.doesNotThrow(() => {
    context.renderDistributionCharts(result, officers, runningTotals);
  });

  assert.ok(attemptedTitles > 0, 'Custom renderer should be attempted before fallback.');
  const distributionCharts = context.__elementsById.distributionCharts;
  assert.equal(distributionCharts.children.length, 6, 'Fallback renderer should still render all chart cards.');
  assert.notEqual(distributionCharts.textContent, 'No distribution charts yet.');
});
