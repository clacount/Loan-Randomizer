const test = require('node:test');
const assert = require('node:assert/strict');

function loadFairnessView(globalScope) {
  global.window = globalScope;
  global.document = globalScope.document;
  delete require.cache[require.resolve('../src/ui/views/fairnessView.js')];
  require('../src/ui/views/fairnessView.js');
  return globalScope.FairnessView;
}

test('live fairness summary groups metrics into sections instead of a flat text wall', () => {
  const appended = [];
  const globalScope = {
    FairnessDisplayService: {
      getFairnessModelLabel: () => 'Officer Lane Fairness',
      buildFairnessNotesForDisplay: () => [
        'Model note: Officer Lane variance is measured within lanes.',
        'Flex officers can appear in consumer composition while mortgage support is evaluated separately.'
      ]
    },
    document: {
      createElement() {
        return {
          className: '',
          innerHTML: ''
        };
      }
    }
  };
  const fairnessView = loadFairnessView(globalScope);
  const containerEl = {
    appendChild(node) {
      appended.push(node);
    }
  };

  fairnessView.renderLiveFairnessSummaryCard(containerEl, {
    overallResult: 'PASS',
    engineType: 'officer_lane',
    statusMetricDescriptor: {
      label: 'Flex lane dollar variance',
      valuePercent: 13.5,
      contextLabel: 'Flex lane thresholds'
    },
    summaryItems: [
      'Consumer loan variance (consumer-eligible lane): 4.0%',
      'Consumer dollar variance (consumer-eligible lane): 4.0%',
      'Mortgage dollar variance (M lane): 11.0%',
      'Flex dollar variance (F lane): 13.5%',
      'Mortgage lane routing to M officers: 100.0%',
      'Flex dollar variance advisory band applied: 20.0%–25.0%'
    ]
  });

  assert.equal(appended.length, 1);
  const html = appended[0].innerHTML;
  assert.match(html, /Overview/);
  assert.match(html, /Consumer Lane/);
  assert.match(html, /Mortgage Lane/);
  assert.match(html, /Flex Lane/);
  assert.match(html, /Policy And Thresholds/);
  assert.match(html, /Interpretation/);
  assert.match(html, /Primary status driver/);
  assert.match(html, /Mortgage lane routing to M officers/);
  assert.match(html, /Flex dollar variance advisory band applied/);
  assert.doesNotMatch(html, /audit-summary-line/);
});
