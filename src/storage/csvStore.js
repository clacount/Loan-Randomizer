function createCsvStore() {
  return {
    read() { return []; },
    append() { return true; },
    findDuplicateByIdOrName() { return { isDuplicate: false }; }
  };
}

window.createCsvStore = createCsvStore;
