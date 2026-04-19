function registerEventHandlers(controllerRegistry = {}) {
  return {
    loanController: controllerRegistry.loanController || null,
    reportController: controllerRegistry.reportController || null,
    officerController: controllerRegistry.officerController || null,
    settingsController: controllerRegistry.settingsController || null
  };
}

window.LoanRandomizerRegisterEventHandlers = registerEventHandlers;
