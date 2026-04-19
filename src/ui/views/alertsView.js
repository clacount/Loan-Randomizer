const alertsView = {
  showError(message) {
    const el = document.getElementById('message');
    if (el) {
      el.textContent = message;
      el.dataset.tone = 'error';
    }
  }
};

window.alertsView = alertsView;
