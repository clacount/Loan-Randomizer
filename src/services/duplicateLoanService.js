function createDuplicateLoanService({ loanRepository }) {
  return {
    check(loan) {
      return loanRepository.findDuplicate(loan);
    }
  };
}

window.createDuplicateLoanService = createDuplicateLoanService;
