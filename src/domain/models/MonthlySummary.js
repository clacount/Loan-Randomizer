function createMonthlySummary({ monthKey, loanCount, totalAmountRequested, officers }) {
  return {
    monthKey,
    loanCount: Number(loanCount) || 0,
    totalAmountRequested: Number(totalAmountRequested) || 0,
    officers: officers || []
  };
}

window.createMonthlySummary = createMonthlySummary;
