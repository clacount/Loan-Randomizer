const idUtils = {
  normalizeId(value) {
    return String(value || '').trim().toLowerCase();
  }
};

window.idUtils = idUtils;
