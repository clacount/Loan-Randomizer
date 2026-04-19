const dateUtils = {
  toMonthKey(date = new Date()) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }
};

window.dateUtils = dateUtils;
