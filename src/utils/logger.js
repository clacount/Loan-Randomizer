const logger = {
  info(...args) { console.info('[LoanRandomizer]', ...args); },
  warn(...args) { console.warn('[LoanRandomizer]', ...args); },
  error(...args) { console.error('[LoanRandomizer]', ...args); }
};

window.logger = logger;
