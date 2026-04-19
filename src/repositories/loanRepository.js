function createLoanRepository({ historyStore }) {
  return {
    save(loanEntry) {
      historyStore.append(loanEntry);
    },
    findDuplicate(loan) {
      return historyStore.findDuplicateByIdOrName(loan);
    }
  };
}

window.createLoanRepository = createLoanRepository;
