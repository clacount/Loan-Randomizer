function createJsonStore() {
  return {
    read() { return null; },
    write() { return true; }
  };
}

window.createJsonStore = createJsonStore;
