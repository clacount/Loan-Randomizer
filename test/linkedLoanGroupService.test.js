const test = require('node:test');
const assert = require('node:assert/strict');

const linkedLoanGroupService = require('../src/services/linkedLoanGroupService.js');

test('buildAssignmentUnits groups linked loans into one assignment unit and keeps ungrouped loans separate', () => {
  const loans = [
    { name: 'L1', type: 'Auto', amountRequested: 10000, linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 2048' },
    { name: 'L2', type: 'Personal', amountRequested: 5000, linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 2048' },
    { name: 'L3', type: 'Credit Card', amountRequested: 2500 }
  ];

  const units = linkedLoanGroupService.buildAssignmentUnits(loans, {
    getLoanCategory: () => 'consumer',
    getLoanAmount: (loan) => loan.amountRequested
  });

  assert.equal(units.length, 2);
  const linkedUnit = units.find((unit) => unit.unitType === 'linked_group');
  assert.equal(linkedUnit.loans.length, 2);
  assert.equal(linkedUnit.totalAmount, 15000);
  assert.equal(linkedUnit.linkedGroupLabel, 'Member 2048');
});

test('validateLinkedLoanGroups rejects groups smaller than two loans', () => {
  const result = linkedLoanGroupService.validateLinkedLoanGroups({
    loans: [{ name: 'L1', type: 'Auto', linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 1' }],
    officers: [{ name: 'A' }]
  });

  assert.equal(result.valid, false);
  assert.equal(result.code, 'LINKED_LOAN_GROUP_TOO_SMALL');
});

test('validateLinkedLoanGroups rejects groups with no officer eligible for every loan', () => {
  const result = linkedLoanGroupService.validateLinkedLoanGroups({
    loans: [
      { name: 'L1', type: 'Auto', linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 1' },
      { name: 'L2', type: 'First Mortgage', linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 1' }
    ],
    officers: [
      { name: 'Consumer', eligibility: { consumer: true, mortgage: false } },
      { name: 'Mortgage', eligibility: { consumer: false, mortgage: true } }
    ],
    getLoanCategory: (loan) => loan.type === 'First Mortgage' ? 'mortgage' : 'consumer',
    isOfficerEligibleForLoan: (officer, loan) => (
      loan.type === 'First Mortgage'
        ? Boolean(officer.eligibility?.mortgage)
        : Boolean(officer.eligibility?.consumer)
    )
  });

  assert.equal(result.valid, false);
  assert.equal(result.code, 'LINKED_LOAN_GROUP_NO_ELIGIBLE_OFFICER');
});

test('validateLinkedLoanGroups accepts consumer-only and mixed groups when at least one officer can take all loans', () => {
  const result = linkedLoanGroupService.validateLinkedLoanGroups({
    loans: [
      { name: 'L1', type: 'Auto', linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 1' },
      { name: 'L2', type: 'First Mortgage', linkedGroupId: 'MLG-001', linkedGroupLabel: 'Member 1' }
    ],
    officers: [
      { name: 'Flex', eligibility: { consumer: true, mortgage: true } }
    ],
    getLoanCategory: (loan) => loan.type === 'First Mortgage' ? 'mortgage' : 'consumer',
    isOfficerEligibleForLoan: (officer, loan) => (
      loan.type === 'First Mortgage'
        ? Boolean(officer.eligibility?.mortgage)
        : Boolean(officer.eligibility?.consumer)
    )
  });

  assert.equal(result.valid, true);
  assert.equal(result.groups.length, 1);
});

test('validateManagerApprovalName requires first and last name', () => {
  assert.equal(linkedLoanGroupService.validateManagerApprovalName('Jane').valid, false);
  assert.equal(linkedLoanGroupService.validateManagerApprovalName('J S').valid, false);
  const valid = linkedLoanGroupService.validateManagerApprovalName('Jane Smith');
  assert.equal(valid.valid, true);
  assert.equal(valid.normalizedName, 'Jane Smith');
});
