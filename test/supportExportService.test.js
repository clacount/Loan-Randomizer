const test = require('node:test');
const assert = require('node:assert/strict');

global.window = global;

const supportExport = require('../src/services/supportExportService.js');

test('support manifest includes app version, active tier, and fairness engine', () => {
  const manifest = supportExport.buildSupportManifest({
    appMetadata: {
      appName: 'LendingFair',
      appVersion: '0.1.0',
      releaseChannel: 'pilot',
      buildDate: '2026-05-02'
    },
    activeTier: 'pro',
    activeTierLabel: 'Pro',
    fairnessEngine: 'officer_lane',
    fairnessEngineLabel: 'Officer Lane Fairness',
    licenseMetadata: {
      licenseId: 'pilot-001',
      licenseType: 'pilot',
      expiresAt: '2026-06-30',
      licenseStatus: 'active',
      activeTier: 'pro'
    },
    exportTimestamp: '2026-05-02T12:00:00.000Z',
    monthFolderKey: '2026-05'
  });

  assert.equal(manifest.app.name, 'LendingFair');
  assert.equal(manifest.app.version, '0.1.0');
  assert.equal(manifest.context.activeTier, 'pro');
  assert.equal(manifest.context.fairnessEngine, 'officer_lane');
  assert.equal(manifest.context.monthFolderKey, '2026-05');
  assert.equal(manifest.license.licenseId, 'pilot-001');
  assert.equal(manifest.license.licenseStatus, 'active');
});

test('support manifest includes linked loan group metadata and approval details', () => {
  const manifest = supportExport.buildSupportManifest({
    linkedLoanGroupMetadata: {
      linkedGroupCount: 1,
      linkedGroups: [
        {
          linkedGroupId: 'MLG-001',
          linkedGroupLabel: 'Member 2048',
          assignedOfficer: 'Avery'
        }
      ],
      approval: {
        approvedBy: 'Jane Smith',
        approvedAt: '2026-05-05T12:00:00.000Z',
        approvalReason: 'Member relationship continuity'
      }
    }
  });

  assert.equal(manifest.linkedLoanGroups.linkedGroupCount, 1);
  assert.equal(manifest.linkedLoanGroups.linkedGroups[0].linkedGroupId, 'MLG-001');
  assert.equal(manifest.linkedLoanGroups.approval.approvedBy, 'Jane Smith');
});

test('missing support files are represented as missing and not fatal', () => {
  const files = [
    supportExport.createIncludedFileRecord({
      fileName: 'loan-randomizer-running-totals.csv',
      kind: supportExport.SUPPORT_FILE_KINDS.RUNNING_TOTALS,
      content: 'Officer,Loan Count\nAvery,1\n',
      size: 27
    }),
    supportExport.createMissingFileRecord({
      fileName: 'loan-randomizer-loan-history.csv',
      kind: supportExport.SUPPORT_FILE_KINDS.LOAN_HISTORY
    })
  ];

  const manifest = supportExport.buildSupportManifest({ files });

  assert.equal(manifest.summary.includedFileCount, 1);
  assert.equal(manifest.summary.missingFileCount, 1);
  assert.equal(manifest.files[1].status, 'missing');
});

test('simulation history is omitted when the active tier cannot use simulation', () => {
  const basicEntitlements = {
    FEATURES: { SIMULATION: 'simulation' },
    canUseFeature(feature) {
      assert.equal(feature, 'simulation');
      return false;
    }
  };

  assert.equal(supportExport.shouldIncludeSimulationHistory(basicEntitlements), false);
});

test('simulation history can be included when the active tier allows simulation', () => {
  const proEntitlements = {
    FEATURES: { SIMULATION: 'simulation' },
    canUseFeature(feature) {
      assert.equal(feature, 'simulation');
      return true;
    }
  };

  assert.equal(supportExport.shouldIncludeSimulationHistory(proEntitlements), true);
});
