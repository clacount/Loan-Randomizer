function createReportController({ reportService }) {
  return {
    generateRunReport(context) {
      return reportService.generateRunReport(context);
    },
    generateEndOfMonthReport(context) {
      return reportService.generateEndOfMonthReport(context);
    }
  };
}

window.createReportController = createReportController;
