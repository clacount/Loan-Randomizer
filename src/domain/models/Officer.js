function createOfficer({ name, activeSessionCount = 0, totals = {} }) {
  return {
    name: String(name || '').trim(),
    activeSessionCount: Number(activeSessionCount) || 0,
    totals
  };
}

window.createOfficer = createOfficer;
