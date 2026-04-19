function createSettingsRepository({ jsonStore, filePath }) {
  return {
    load() {
      return jsonStore.read(filePath) || {};
    },
    save(settings) {
      jsonStore.write(filePath, settings);
    }
  };
}

window.createSettingsRepository = createSettingsRepository;
