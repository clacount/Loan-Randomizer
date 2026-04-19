function createHistoryRepository({ csvStore, filePath }) {
  return {
    readAll() {
      return csvStore.read(filePath);
    },
    append(entry) {
      return csvStore.append(filePath, entry);
    },
    findDuplicateByIdOrName(loan) {
      return csvStore.findDuplicateByIdOrName(filePath, loan);
    }
  };
}

window.createHistoryRepository = createHistoryRepository;
