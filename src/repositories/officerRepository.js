function createOfficerRepository({ jsonStore, filePath }) {
  return {
    getAll() {
      return jsonStore.read(filePath) || [];
    },
    saveAll(officers) {
      jsonStore.write(filePath, officers);
    }
  };
}

window.createOfficerRepository = createOfficerRepository;
