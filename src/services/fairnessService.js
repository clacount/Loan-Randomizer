function createFairnessService({ assignmentRules, fairnessRules }) {
  return {
    chooseOfficer(loan, context = {}) {
      return assignmentRules.assignLoan(loan, fairnessRules, context);
    }
  };
}

window.createFairnessService = createFairnessService;
