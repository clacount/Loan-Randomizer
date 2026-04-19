function createRandomizerService({ fairnessService }) {
  return {
    assign(loan) {
      return fairnessService.chooseOfficer(loan);
    }
  };
}

window.createRandomizerService = createRandomizerService;
