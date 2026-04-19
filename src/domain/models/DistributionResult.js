function createDistributionResult({ assignments = [], audit = [] } = {}) {
  return { assignments, audit };
}

window.createDistributionResult = createDistributionResult;
