const loanFormView = {
  getLoanInput() {
    return {
      id: document.getElementById('loanName')?.value || '',
      name: document.getElementById('loanName')?.value || '',
      type: document.getElementById('loanType')?.value || '',
      amountRequested: document.getElementById('loanAmount')?.value || ''
    };
  }
};

window.loanFormView = loanFormView;
