function createLoanController({ loanService, loanFormView, resultsView, alertsView }) {
  return {
    submitLoan(rawLoanInput) {
      const validation = loanService.validate(rawLoanInput);
      if (!validation.isValid) {
        alertsView.showError(validation.message);
        return { ok: false, error: validation.message };
      }

      const result = loanService.submitLoan(rawLoanInput);
      resultsView.renderAssignment(result);
      return { ok: true, result };
    },
    readLoanInput() {
      return loanFormView.getLoanInput();
    }
  };
}

window.createLoanController = createLoanController;
