function createLoan({ id, name, type, amountRequested }) {
  return {
    id: String(id || '').trim(),
    name: String(name || '').trim(),
    type: String(type || '').trim(),
    amountRequested: Number(amountRequested) || 0
  };
}

window.createLoan = createLoan;
