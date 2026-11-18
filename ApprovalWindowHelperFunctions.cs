using System.IO;
using System.Windows.Documents;

namespace FileDownload
{
    public class ApprovalWindowHelperFunctions
    {
        public static void ClearPdfs()
        {
            string pdfPath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Pdfs\";
            foreach (string file in Directory.GetFiles(pdfPath))
            {
                try { File.Delete(file); }
                catch { /* ignore */ }
            }
        }

        public static bool? CheckForRequestNumber(List<int> _checkList, int _requestNumber)
        {
            List<int> checkList = _checkList;
            int requestNumber = _requestNumber;

            //Check to see if the request number exists in the excel sheet
            //if the id doesn't exist on the rap calc it will return 0, in this case we want to show null so a X doesn't apear on the UI
            bool? result = requestNumber > 0 ? checkList.Contains(requestNumber) : null;

            return result;
        }

        public static bool? CheckForExGratia(List<string> checkList, RapData rap, int quarter)
        {
            string valueToCheck = $"{rap.requesterNumber}Q{quarter}-{rap.year % 100:D2}";

            double exGratiaValue = quarter switch
            {
                1 => rap.exGratia1,
                2 => rap.exGratia2,
                3 => rap.exGratia3,
                4 => rap.exGratia4,
                _ => throw new ArgumentOutOfRangeException(nameof(quarter), "Quarter must be 1–4")
            };

            // If the ex gratia value is 0, it is not filled in on the RAP so we don't want to show on the Grid
            if (exGratiaValue == 0)
                return null;

            // If the ex gratia value exists we check for it in the cooresponding sheet
            return checkList.Contains(valueToCheck);
        }

        public static List<GridRow> CheckHistory(RapData rapData,List<int> WeeklyRAPApprovalIds, List<int> PaymentSpreadsheetIds, List<int> RapTrackerIds, List<string> ExGratiaIds, int runCount = 0)
        {

            RapHistoryCheck quarterlyCheck = new();
            //Do all 4 checks on the wrap data we have
            quarterlyCheck.RapApproval1 = CheckForRequestNumber(WeeklyRAPApprovalIds, rapData.requestId1);
            quarterlyCheck.RapApproval2 = CheckForRequestNumber(WeeklyRAPApprovalIds, rapData.requestId2);
            quarterlyCheck.RapApproval3 = CheckForRequestNumber(WeeklyRAPApprovalIds, rapData.requestId3);
            quarterlyCheck.RapApproval4 = CheckForRequestNumber(WeeklyRAPApprovalIds, rapData.requestId4);

            quarterlyCheck.PaymentSheet1 = CheckForRequestNumber(PaymentSpreadsheetIds, rapData.requestId1);
            quarterlyCheck.PaymentSheet2 = CheckForRequestNumber(PaymentSpreadsheetIds, rapData.requestId2);
            quarterlyCheck.PaymentSheet3 = CheckForRequestNumber(PaymentSpreadsheetIds, rapData.requestId3);
            quarterlyCheck.PaymentSheet4 = CheckForRequestNumber(PaymentSpreadsheetIds, rapData.requestId4);

            quarterlyCheck.RapTracker1 = CheckForRequestNumber(RapTrackerIds, rapData.requestId1);
            quarterlyCheck.RapTracker2 = CheckForRequestNumber(RapTrackerIds, rapData.requestId2);
            quarterlyCheck.RapTracker3 = CheckForRequestNumber(RapTrackerIds, rapData.requestId3);
            quarterlyCheck.RapTracker4 = CheckForRequestNumber(RapTrackerIds, rapData.requestId4);

            quarterlyCheck.ExGratia1 = CheckForExGratia(ExGratiaIds, rapData, 1);
            quarterlyCheck.ExGratia2 = CheckForExGratia(ExGratiaIds, rapData, 2);
            quarterlyCheck.ExGratia3 = CheckForExGratia(ExGratiaIds, rapData, 3);
            quarterlyCheck.ExGratia4 = CheckForExGratia(ExGratiaIds, rapData, 4);

            quarterlyCheck.Historical1 = rapData.historicalQuarter1;
            quarterlyCheck.Historical2 = rapData.historicalQuarter2;
            quarterlyCheck.Historical3 = rapData.historicalQuarter3;
            quarterlyCheck.Historical4 = rapData.historicalQuarter4;

            var retDataContext = new List<GridRow>
            {
                new GridRow { Header = "Q1", C1=quarterlyCheck.RapTracker1,C2=quarterlyCheck.RapApproval1,C3=quarterlyCheck.PaymentSheet1,C4=quarterlyCheck.ExGratia1,C5=quarterlyCheck.Historical1},
                new GridRow { Header = "Q2", C1=quarterlyCheck.RapTracker2,C2=quarterlyCheck.RapApproval2,C3=quarterlyCheck.PaymentSheet2,C4=quarterlyCheck.ExGratia2,C5=quarterlyCheck.Historical2},
                new GridRow { Header = "Q3", C1=quarterlyCheck.RapTracker3,C2=quarterlyCheck.RapApproval3,C3=quarterlyCheck.PaymentSheet3,C4=quarterlyCheck.ExGratia3,C5=quarterlyCheck.Historical3},
                new GridRow { Header = "Q4", C1=quarterlyCheck.RapTracker4,C2=quarterlyCheck.RapApproval4,C3=quarterlyCheck.PaymentSheet4,C4=quarterlyCheck.ExGratia4,C5=quarterlyCheck.Historical4}
            };

            return retDataContext;
        }


        public static async Task WaitUntilAsync(Func<bool> condition, int timeoutMs = 30000, int pollInterval = 200)
        {
            var start = DateTime.UtcNow;

            while (!condition())
            {
                if ((DateTime.UtcNow - start).TotalMilliseconds > timeoutMs)
                    throw new TimeoutException("Condition was not met in time.");

                await Task.Delay(pollInterval);
            }
        }


    }
    public class RapData
    {
        internal string originalPrincipal;
        internal string RAPUntaxedIncome;

        public string LRAPStudentId { get; set; }
        public int requestId1 { get; set; }
        public int requestId2 { get; set; }
        public int requestId3 { get; set; }
        public int requestId4 { get; set; }
        public int requestNumber1 { get; set; }
        public int requestNumber2 { get; set; }
        public int requestNumber3 { get; set; }
        public int requestNumber4 { get; set; }
        public int year { get; set; }
        public int requesterNumber { get; set; }
        public double exGratia1 { get; set; }
        public double exGratia2 { get; set; }
        public double exGratia3 { get; set; }
        public double exGratia4 { get; set; }
        public bool? historicalQuarter1 { get; set; }
        public bool? historicalQuarter2 { get; set; }
        public bool? historicalQuarter3 { get; set; }
        public bool? historicalQuarter4 { get; set; }
        public string cityState { get; internal set; }
        public string married { get; internal set; }
        public string cohortAdmitType { get; internal set; }
        public string initialEligibilityMet { get; internal set; }
        public string ach { get; internal set; }
        public string currentAddress { get; internal set; }
        public string firstName { get; internal set; }
        public string lastName { get; internal set; }
        public string w9 { get; internal set; }
        public string maiden { get; internal set; }
        public string originalPercentCovered { get; internal set; }
        public string originalLRAPPrincipal { get; internal set; }
        public string stateResidence { get; internal set; }
        public string familySize { get; internal set; }
        public string tcHouseholdIncome { get; internal set; }
        public string tcUntaxedIncome { get; internal set; }
        public string currentPrincipal { get; internal set; }
        public string currentPercentCovered { get; internal set; }
        public string currentLRAPPrincipal { get; internal set; }
        public string uit { get; internal set; }
        public string lit { get; internal set; }
        public string minWageTC { get; internal set; }
        public string percentRepaid { get; internal set; }
        public string totalReimbursementPaid { get; internal set; }
        public string lrapAwardDate { get; internal set; }
        public string annualLimit { get; internal set; }
        public string aggregateLimit { get; internal set; }
        public string priorYearOnFile { get; internal set; }
        public string initEligibilityTerms { get; internal set; }
        public string agiVsAnnualized { get; internal set; }
        public string idrImplemented { get; internal set; }
        public string monthlyPaymentBeforeProration { get; internal set; }
        public string monthlyPaymentAfterBorrowingProration { get; internal set; }
        public string graduationDate { get; internal set; }
        public string gradVerified { get; internal set; }
        public string major { get; internal set; }
        public string lrapRetro { get; internal set; }
        public string effectiveTerm { get; internal set; }
        public string requestorInPlan { get; internal set; }
        public string requestorEligibleFor { get; internal set; }
        public string monthlyPaymentOnPlan { get; internal set; }
        public string quarterIdrImplemented { get; internal set; }
    }

    public class RapHistoryCheck
    {
        public bool? RapApproval1 { get; set; }
        public bool? RapApproval2 { get; set; }
        public bool? RapApproval3 { get; set; }
        public bool? RapApproval4 { get; set; }
        public bool? PaymentSheet1 { get; set; }
        public bool? PaymentSheet2 { get; set; }
        public bool? PaymentSheet3 { get; set; }
        public bool? PaymentSheet4 { get; set; }
        public bool? ExGratia1 { get; set; }
        public bool? ExGratia2 { get; set; }
        public bool? ExGratia3 { get; set; }
        public bool? ExGratia4 { get; set; }
        public bool? RapTracker1 { get; set; }
        public bool? RapTracker2 { get; set; }
        public bool? RapTracker3 { get; set; }
        public bool? RapTracker4 { get; set; }
        public bool? Historical1 { get; set; }
        public bool? Historical2 { get; set; }
        public bool? Historical3 { get; set; }
        public bool? Historical4 { get; set; }
    }

    public class GridRow
    {
        public string Header { get; set; }
        public bool? C1 { get; set; }
        public bool? C2 { get; set; }
        public bool? C3 { get; set; }
        public bool? C4 { get; set; }
        public bool? C5 { get; set; }
    }

    public class RapFile
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
    }

    public class RapCalcTableData
    {
        public string Header { get; set; }
        public string RequestNumber { get; set; }
        public string MonthsRequestedInQuarter { get; set; }
        public string MonthsInQuarterQualifying { get; set; }
        public string ActualIncome { get; set; }
        public string AnnualizedIncome { get; set; }
        public string BorrowingProration { get; set; }
        public string IncomeProration { get; set; }
        public string TimeWorkedProration { get; set; }
        public string EligibleLoanPayments { get; set; }
        public string IdrEligibleLoanPayments { get; set; }
        public string EligibleLoanPaymentsYtd { get; set; }
        public string AssistanceDueYtd { get; set; }
        public string AssistancePaidYtd { get; set; }
        public string AssistanceDueNow { get; set; }
        public string AssistancePaidNow { get; set; }
        public string PriorPeriodAdjustment { get; set; }
        public string CheckAmount { get; set; }
        public string Date { get; set; }
        public string Notes { get; set; }
        public string TimeWorkedExGratia { get; set; }
    }
}
