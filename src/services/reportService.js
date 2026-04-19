function createReportService({ loanRunPdfGenerator, endOfMonthPdfGenerator, reportFormatter }) {
  return {
    generateRunReport(context) {
      return loanRunPdfGenerator.generate(reportFormatter.formatRun(context));
    },
    generateEndOfMonthReport(context) {
      return endOfMonthPdfGenerator.generate(reportFormatter.formatEom(context));
    }
  };
}

window.createReportService = createReportService;
