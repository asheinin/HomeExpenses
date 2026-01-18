function main() {
}

// global variables
MAILER = 'HomePayments';
FILENAME = 'Home payments';


function staticNumbers() { // constructor function
  this.thresholdLimitForClosingMonth = 1;

  this.expenseCarryOverRow = 2;
  this.expenseSpToSpRow = 2;
  this.expenseFirstRow = 3;
  this.expenseLastRow = 50;
  this.expenseTotalAmountRow = 53;
  this.expenseTotalPaidSp2Row = 56;
  this.expenseTotalPaidSp1Row = 60;
  this.expenseSp2MonthlyBalanceRow = 63;
  this.expenseSp1MonthlyBalanceRow = 64;
  this.expenseSp2InitBalancePaidRow = 66;
  this.expenseSp2InitBalanceLeftRow = 67;
  this.expenseSp1InitBalancePaidRow = 69;
  this.expenseSp1InitBalanceLeftRow = 70;

  this.expenseTypeColumn = 1;
  this.expenseDescrColumn = 2;
  this.expenseDateColumn = 3;
  this.expenseAmountColumn = 4;
  this.expenseCarryOverSp2OwesColumn = 5;
  this.expenseCarryOverSp1OwesColumn = 6;
  this.expenceSplit2Column = 5;
  this.expenceSplit1Column = 6;
  this.expenceSplitColumn = 7;
  this.expenseFirstPayColumn = 8;
  this.expenseSecondPayColumn = 9;
  this.expenseInitialBalanceCol = 2;
  this.expencePeriodColumn = 10;
  this.expensePaidColumn = 11;
  this.expensePAPColumn = 12;
  this.expenseSp2ToSp1Column = 13;
  this.expenseSp1ToSp2Column = 14;

  this.dashAddressRow = 1;
  this.dashNamesRow = 2;
  this.dashEmailsRow = 3;
  this.dashSplitRow = 4
  this.dashBalancesRow = 5;
  this.dashTitleRow = 6;
  this.dashFirstMonthRow = 7;

  this.dashSpouse1NameColumn = 2;
  this.dashSpouse2NameColumn = 3;
  this.dashAddressColumn = 2;
  this.dashSp1SplitColumn = 2;
  this.dashSp2SplitColumn = 3;



  this.dashMonthNameColumn = 1;
  this.dashSp1BalanceUsedColumn = 2;
  this.dashSp2BalanceUsedColumn = 3;




  this.dashAmountTotalBeforeSplitColumn = 4;
  this.dashAmountTotalColumn = 5;
  this.dashSp2PartColumn = 6;
  this.dashSp1PartColumn = 7;
  this.dashSp2PaidColumn = 8;
  this.dashSp1PaidColumn = 9;
  this.dashSp2ToSp1Column = 10;
  this.dashSp1ToSp2Column = 11;
  this.dashSp2BalanceColumn = 12;
  this.dashSp1BalanceColumn = 13;
  this.dashColumns = 13;

  this.summaryHeaderRow = 1
  this.summarySumRow = 2
  this.summaryMinStartAnalyticsRow = 28;

  this.summaryAmountColumn = 3;
  this.summaryAnalyticsYearColumn = 1;
  this.summaryAnalyticsDataStartColumn = 4;
  this.summaryChartsStartColumn = 17;
  this.summaryAnalyticsMonthColumn = 1;

  this.dashBalanceNegativeBgColor = "red";
  this.dashBalancePositiveBgColor = "green";
  this.dashBalanceNeutralBgColor = "green";

  this.privilegedMethod = function () {
    alert();
  };
}