(function initializeLinkedLoanGroupService(globalScope) {
  const DEFAULT_REASON = 'Member relationship continuity';

  function normalizeLinkedGroupReason(reason) {
    const normalized = String(reason || '').trim();
    return normalized || DEFAULT_REASON;
  }

  function normalizeLinkedGroupId(linkedGroupId) {
    return String(linkedGroupId || '').trim();
  }

  function normalizeLinkedGroupLabel(linkedGroupLabel, linkedGroupId) {
    const normalizedLabel = String(linkedGroupLabel || '').trim();
    if (normalizedLabel) {
      return normalizedLabel;
    }
    const normalizedId = normalizeLinkedGroupId(linkedGroupId);
    return normalizedId || '';
  }

  function buildLinkedLoanGroups(loans = []) {
    const groupsById = new Map();

    (Array.isArray(loans) ? loans : []).forEach((loan) => {
      const linkedGroupId = normalizeLinkedGroupId(loan?.linkedGroupId);
      if (!linkedGroupId) {
        return;
      }

      const existingGroup = groupsById.get(linkedGroupId) || {
        linkedGroupId,
        linkedGroupLabel: normalizeLinkedGroupLabel(loan?.linkedGroupLabel, linkedGroupId),
        linkedGroupReason: normalizeLinkedGroupReason(loan?.linkedGroupReason),
        loans: []
      };
      existingGroup.loans.push(loan);
      if (!existingGroup.linkedGroupLabel) {
        existingGroup.linkedGroupLabel = normalizeLinkedGroupLabel(loan?.linkedGroupLabel, linkedGroupId);
      }
      existingGroup.linkedGroupReason = normalizeLinkedGroupReason(existingGroup.linkedGroupReason || loan?.linkedGroupReason);
      groupsById.set(linkedGroupId, existingGroup);
    });

    return [...groupsById.values()].sort((groupA, groupB) => {
      const labelA = String(groupA.linkedGroupLabel || groupA.linkedGroupId || '');
      const labelB = String(groupB.linkedGroupLabel || groupB.linkedGroupId || '');
      return labelA.localeCompare(labelB);
    });
  }

  function buildAssignmentUnits(loans = [], options = {}) {
    const getLoanCategory = typeof options.getLoanCategory === 'function'
      ? options.getLoanCategory
      : ((loan) => String(loan?.category || loan?.loanCategory || 'consumer').trim().toLowerCase() || 'consumer');
    const getLoanAmount = typeof options.getLoanAmount === 'function'
      ? options.getLoanAmount
      : ((loan) => Number(loan?.amountRequested) || 0);

    const groupedLoans = buildLinkedLoanGroups(loans);
    const groupedLoanIds = new Set();
    groupedLoans.forEach((group) => {
      group.loans.forEach((loan) => groupedLoanIds.add(loan));
    });

    const units = [];
    groupedLoans.forEach((group) => {
      const loanTypes = group.loans.map((loan) => loan.type).filter(Boolean);
      const loanCategories = group.loans.map((loan) => getLoanCategory(loan)).filter(Boolean);
      units.push({
        unitId: group.linkedGroupId,
        unitType: 'linked_group',
        loans: [...group.loans],
        totalAmount: group.loans.reduce((sum, loan) => sum + getLoanAmount(loan), 0),
        loanTypes,
        loanCategories,
        linkedGroupId: group.linkedGroupId,
        linkedGroupLabel: group.linkedGroupLabel,
        linkedGroupReason: group.linkedGroupReason
      });
    });

    (Array.isArray(loans) ? loans : []).forEach((loan) => {
      if (groupedLoanIds.has(loan)) {
        return;
      }
      units.push({
        unitId: loan?.name || `single-${units.length + 1}`,
        unitType: 'single_loan',
        loans: [loan],
        totalAmount: getLoanAmount(loan),
        loanTypes: loan?.type ? [loan.type] : [],
        loanCategories: [getLoanCategory(loan)],
        linkedGroupId: '',
        linkedGroupLabel: '',
        linkedGroupReason: ''
      });
    });

    return units;
  }

  function validateManagerApprovalName(name) {
    const trimmed = String(name || '').trim().replace(/\s+/g, ' ');
    const parts = trimmed.split(' ').filter(Boolean);
    const valid = parts.length >= 2 && parts.every((part) => part.length >= 2);
    return {
      valid,
      normalizedName: trimmed,
      message: valid ? '' : 'Enter the approving manager first and last name.'
    };
  }

  function requiresLinkedGroupReviewApproval(result = {}) {
    return String(result?.fairnessReview?.selectedStatus || result?.fairnessEvaluation?.overallResult || '').toUpperCase() === 'REVIEW'
      && Array.isArray(result?.linkedLoanGroups)
      && result.linkedLoanGroups.length > 0;
  }

  function validateLinkedLoanGroups({
    loans = [],
    officers = [],
    canUseLinkedLoanGroups = true,
    getLockedFeatureMessage = () => 'Linked loan groups require Pro or Platinum.',
    getLoanCategory = (loan) => String(loan?.category || loan?.loanCategory || 'consumer').trim().toLowerCase() || 'consumer',
    isOfficerEligibleForLoan = () => true
  } = {}) {
    const groups = buildLinkedLoanGroups(loans);
    if (groups.length && !canUseLinkedLoanGroups) {
      return {
        valid: false,
        code: 'LINKED_LOAN_GROUPS_NOT_AVAILABLE',
        message: getLockedFeatureMessage()
      };
    }

    for (const group of groups) {
      if (group.loans.length < 2) {
        return {
          valid: false,
          code: 'LINKED_LOAN_GROUP_TOO_SMALL',
          message: `Linked group ${group.linkedGroupLabel || group.linkedGroupId} must include at least two loans.`
        };
      }

      const eligibleOfficers = (Array.isArray(officers) ? officers : []).filter((officer) => {
        if (officer?.isOnVacation) {
          return false;
        }
        return group.loans.every((loan) => isOfficerEligibleForLoan(officer, loan));
      });

      if (!eligibleOfficers.length) {
        return {
          valid: false,
          code: 'LINKED_LOAN_GROUP_NO_ELIGIBLE_OFFICER',
          message: `Linked group ${group.linkedGroupLabel || group.linkedGroupId} contains loans that cannot be assigned to one eligible officer. Adjust the group or officer eligibility before running assignment.`
        };
      }

      const categories = new Set(group.loans.map((loan) => getLoanCategory(loan)));
      group.loanCategories = [...categories];
    }

    return {
      valid: true,
      groups
    };
  }

  const api = {
    DEFAULT_REASON,
    normalizeLinkedGroupReason,
    normalizeLinkedGroupId,
    normalizeLinkedGroupLabel,
    buildLinkedLoanGroups,
    buildAssignmentUnits,
    validateLinkedLoanGroups,
    validateManagerApprovalName,
    requiresLinkedGroupReviewApproval
  };

  globalScope.LinkedLoanGroupService = api;
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
