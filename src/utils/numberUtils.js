const numberUtils = {
  toNumber(value, fallback = 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }
};

window.numberUtils = numberUtils;
