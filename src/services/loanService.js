function createLoanService({ loanValidator, duplicateLoanService, randomizerService, loanRepository }) {
  return {
    validate(rawLoanInput) {
      return loanValidator.validate(rawLoanInput);
    },
    submitLoan(rawLoanInput) {
      const duplicateCheck = duplicateLoanService.check(rawLoanInput);
      if (duplicateCheck.isDuplicate) {
        throw new Error(duplicateCheck.message);
      }
      const assignment = randomizerService.assign(rawLoanInput);
      loanRepository.save({ ...rawLoanInput, assignment });
      return assignment;
    }
  };
}

window.createLoanService = createLoanService;
