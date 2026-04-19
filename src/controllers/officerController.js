function createOfficerController({ officerRepository, officerTotalsView }) {
  return {
    listOfficers() {
      return officerRepository.getAll();
    },
    updateTotals(officerTotals) {
      officerTotalsView.render(officerTotals);
    }
  };
}

window.createOfficerController = createOfficerController;
