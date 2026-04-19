const loanValidator = {
  validate(loan) {
    if (!String(loan?.name || '').trim()) {
      return { isValid: false, message: 'Loan name is required.' };
    }
    return { isValid: true };
  }
};

window.loanValidator = loanValidator;
