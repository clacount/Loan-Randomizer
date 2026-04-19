function createSettingsController({ settingsRepository, settingsValidator }) {
  return {
    save(settings) {
      const validation = settingsValidator.validate(settings);
      if (!validation.isValid) {
        return validation;
      }
      settingsRepository.save(settings);
      return { isValid: true };
    }
  };
}

window.createSettingsController = createSettingsController;
