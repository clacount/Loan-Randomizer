function createSimulationService({ fairnessService }) {
  return {
    runSimulation(seedData) {
      return fairnessService.chooseOfficer(seedData.loan, seedData.context);
    }
  };
}

window.createSimulationService = createSimulationService;
