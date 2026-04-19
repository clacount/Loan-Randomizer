const officerValidator = {
  validate(officer) {
    return { isValid: Boolean(String(officer?.name || '').trim()) };
  }
};

window.officerValidator = officerValidator;
