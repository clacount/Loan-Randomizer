const settingsValidator = {
  validate(settings) {
    return { isValid: Boolean(settings) };
  }
};

window.settingsValidator = settingsValidator;
