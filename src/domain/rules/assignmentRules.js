const assignmentRules = {
  assignLoan(loan, fairnessRules, context = {}) {
    return fairnessRules.pickBestOfficer(loan, context);
  }
};

window.assignmentRules = assignmentRules;
