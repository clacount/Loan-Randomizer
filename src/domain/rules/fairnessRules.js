const fairnessRules = {
  pickBestOfficer(loan, context = {}) {
    if (typeof context.pickBestOfficer === 'function') {
      return context.pickBestOfficer(loan);
    }
    return { assignedOfficer: null, reason: 'No fairness engine configured.' };
  }
};

window.fairnessRules = fairnessRules;
