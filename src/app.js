(function bootstrapLoanRandomizerApp() {
  function init() {
    if (typeof window.initializeLoanRandomizerApp === 'function') {
      window.initializeLoanRandomizerApp();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
