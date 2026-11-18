using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileDownload
{
    public class GetRAPExcelData
    {

        public static RapData GetRAPCalcData(string rapFileName)
        {
            //These are used to find request ids in the wrap calc on the RAP Calculator tab and see what quarter it is
            string rowFinder, requestIdFinder, rowFinderDemographic, requestNumberFinder;
            int quarterCounter = 1;

            int colOverallRequestNumber = 4;
            int startRowOverallRequestNumber = 2;
            int colSearcher = 1;
            int startRowRequestorDemographicInformation = 20;
            int startRowExGratia = 50;
            int colExGratia = 4;
            int rowsToScan = 30;

            RapData rap = new RapData();

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("RAP Calculator");

                //Find and save all of the overall request #s and assign them their proper quarter
                for (int row = startRowOverallRequestNumber; row <= startRowOverallRequestNumber + rowsToScan; row++)
                {
                    rowFinder = worksheet.Cell(row, colOverallRequestNumber).GetString();
                    if (rowFinder == "Overall Request #")
                    {
                        requestIdFinder = worksheet.Cell(row, colOverallRequestNumber + 1).GetString();
                        requestNumberFinder = worksheet.Cell(row - 1, colOverallRequestNumber + 1).GetString();

                        if (int.TryParse(requestIdFinder, out int result))
                        {

                            switch (quarterCounter)
                            {
                                case 1:
                                    rap.requestId1 = result;
                                    rap.requestNumber1 = Int32.Parse(requestNumberFinder);
                                    break;

                                case 2:
                                    rap.requestId2 = result;
                                    rap.requestNumber2 = Int32.Parse(requestNumberFinder);
                                    break;

                                case 3:
                                    rap.requestId3 = result;
                                    rap.requestNumber3 = Int32.Parse(requestNumberFinder);
                                    break;

                                case 4:
                                    rap.requestId4 = result;
                                    rap.requestNumber4 = Int32.Parse(requestNumberFinder);
                                    break;

                                default:
                                    Console.WriteLine("Invalid option");
                                    break;
                            }
                        }
                        quarterCounter++;
                    }
                }

                //Find and save the information in the middle of the RAP Calculator Sheet
                for (int row = startRowRequestorDemographicInformation; row <= startRowRequestorDemographicInformation + rowsToScan; row++)
                {
                    rowFinderDemographic = worksheet.Cell(row, colSearcher).GetString();

                    if (rowFinderDemographic == "Requestor #")
                    {
                        requestIdFinder = worksheet.Cell(row, colSearcher + 1).GetString();

                        if (int.TryParse(requestIdFinder, out int result))
                        {
                            rap.requesterNumber = result;
                        }
                        else
                        {
                            //TODO: error processing here
                        }
                    }
                    else if (rowFinderDemographic == "LRAP Student ID")
                    {
                        rap.LRAPStudentId = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic == "Married?")
                    {
                        rap.married = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic == "Cohort/Admit Type")
                    {
                        rap.cohortAdmitType = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic == "City, State of Residency")
                    {
                        rap.cityState = worksheet.Cell(row, colSearcher + 1).GetString();
                    }

                    else if (rowFinderDemographic == "Year Requesting")
                    {
                        requestIdFinder = worksheet.Cell(row, colSearcher + 1).GetString();

                        if (int.TryParse(requestIdFinder, out int result))
                        {
                            rap.year = result;
                        }
                        else
                        {
                            //TODO: error processing here
                        }
                    }
                    else if (rowFinderDemographic == "Initial Eligibility Met")
                    {
                        rap.initialEligibilityMet = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic == "ACH?")
                    {
                        rap.ach = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic == "Current Address")
                    {
                        rap.currentAddress = worksheet.Cell(row, colSearcher + 1).GetString();

                        //Move to the next section of data
                        colSearcher += 3;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("Name"))
                    {
                        rap.firstName = worksheet.Cell(row, colSearcher + 1).GetString();
                        rap.lastName = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("W-9 on file?"))
                    {
                        rap.w9 = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Maiden"))
                    {
                        rap.maiden = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 2;
                        row = startRowRequestorDemographicInformation;
                         
                    }
                    else if (rowFinderDemographic.Equals("Original Principal"))
                    {
                        rap.originalPrincipal = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("% Covered"))
                    {
                        rap.originalPercentCovered = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Original LRAP Principal"))
                    {
                        rap.originalLRAPPrincipal = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 2;
                        row = startRowRequestorDemographicInformation;

                    }
                    else if (rowFinderDemographic.Equals("State of Residence"))
                    {
                        rap.stateResidence = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Family size?"))
                    {
                        rap.familySize = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("T&C Includes Household Income?"))
                    {
                        rap.tcHouseholdIncome = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("T&C Includes Untaxed income?"))
                    {
                        rap.tcUntaxedIncome = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 1;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if(rowFinderDemographic.Equals("Current Principal"))
                    {
                        rap.currentPrincipal = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("% Covered"))
                    {
                        rap.currentPercentCovered = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Current LRAP Principal"))
                    {
                        rap.currentLRAPPrincipal = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 2;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("UIT"))
                    {
                        rap.uit = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("LIT"))
                    {
                        rap.lit = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Min Wage in T&C"))
                    {
                        rap.minWageTC = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("RAPper has Untaxed Income?"))
                    {
                        rap.RAPUntaxedIncome = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 1;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("% Repaid"))
                    {
                        rap.percentRepaid = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("TOTAL Reimbursement Paid"))
                    {
                        rap.totalReimbursementPaid = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 2;
                        row = startRowRequestorDemographicInformation;
                    }

                    
                    else if (rowFinderDemographic.Equals("LRAP Award Date"))
                    {
                        rap.lrapAwardDate = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Annual Limit"))
                    {
                        rap.annualLimit = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Aggregate Limit"))
                    {
                        rap.aggregateLimit = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Prior Yr 1040 on File"))
                    {
                        rap.priorYearOnFile = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Init. Eligibilty Terms"))
                    {
                        rap.initEligibilityTerms = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("AGI vs Annualized"))
                    {
                        rap.agiVsAnnualized = worksheet.Cell(row, colSearcher + 2).GetString();

                        colSearcher += 1;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("IDR Implemented?"))
                    {
                        rap.idrImplemented = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 1;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("Monthly payment before proration"))
                    {
                        rap.monthlyPaymentBeforeProration = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Monthly payment after borrowing proration"))
                    {
                        rap.monthlyPaymentAfterBorrowingProration = worksheet.Cell(row, colSearcher + 1).GetString();

                        colSearcher += 2;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("Graduation Date"))
                    {
                        rap.graduationDate = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Grad Verified?"))
                    {
                        rap.gradVerified = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Major"))
                    {
                        rap.major = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("LRAP Retro"))
                    {
                        rap.lrapRetro = worksheet.Cell(row, colSearcher + 2).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Effective Term"))
                    {
                        rap.effectiveTerm = worksheet.Cell(row, colSearcher + 2).GetString();

                        colSearcher += 1;
                        row = startRowRequestorDemographicInformation;
                    }
                    else if (rowFinderDemographic.Equals("Requestor in IDR/PAYE Plan?"))
                    {
                        rap.requestorInPlan = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Requestor Eligible for?"))
                    {
                        rap.requestorEligibleFor = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("MonthlyPaymentOnPlan"))
                    {
                        rap.monthlyPaymentOnPlan = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                    else if (rowFinderDemographic.Equals("Quarter IDR Implemented?"))
                    {
                        rap.quarterIdrImplemented = worksheet.Cell(row, colSearcher + 1).GetString();
                    }
                }
                //Find and save the Ex Gratia amounts
                for (int row = startRowExGratia; row <= startRowExGratia + rowsToScan; row++)
                {
                    rowFinder = worksheet.Cell(row, colExGratia).GetString();

                    if (rowFinder.Equals("Time-Worked Ex Gratia:"))
                    {
                        //Get the 4 ExGratia amounts in the RAP theres a space between the header and values
                        var requestIdFinder1 = worksheet.Cell(row, colExGratia + 2).GetString();
                        var requestIdFinder2 = worksheet.Cell(row, colExGratia + 3).GetString();
                        var requestIdFinder3 = worksheet.Cell(row, colExGratia + 4).GetString();
                        var requestIdFinder4 = worksheet.Cell(row, colExGratia + 5).GetString();

                        rap.exGratia1 = double.TryParse(requestIdFinder1, out double result1) ? result1 : 0;
                        rap.exGratia2 = double.TryParse(requestIdFinder2, out double result2) ? result2 : 0;
                        rap.exGratia3 = double.TryParse(requestIdFinder3, out double result3) ? result3 : 0;
                        rap.exGratia4 = double.TryParse(requestIdFinder4, out double result4) ? result4 : 0;
                    }
                }

                var historicalWorksheet = workbook.Worksheet("Historical and Analytics");
                string valueToFind = $"Q1-{rap.year % 100:D2}";
                bool quarterFound = false;
                string finder;
                int quarterRow = 2;
                int colCounter = 1;
                while (!quarterFound)
                {
                   finder = historicalWorksheet.Cell(quarterRow, colCounter).GetString();

                    if (finder.Equals(valueToFind))
                    {
                        quarterFound = true;
                        int requestRow = quarterRow + 2;
                        rap.historicalQuarter1 = rap.requestId1 > 0 ? int.TryParse(historicalWorksheet.Cell(requestRow, colCounter).GetString(), out int result1) ? true : false : null;
                        rap.historicalQuarter2 = rap.requestId2 > 0 ? int.TryParse(historicalWorksheet.Cell(requestRow, colCounter + 1).GetString(), out int result2) ? true : false : null;
                        rap.historicalQuarter3 = rap.requestId3 > 0 ? int.TryParse(historicalWorksheet.Cell(requestRow, colCounter + 2).GetString(), out int result3) ? true : false : null;
                        rap.historicalQuarter4 = rap.requestId4 > 0 ? int.TryParse(historicalWorksheet.Cell(requestRow, colCounter + 3).GetString(), out int result4) ? true : false : null;

                    }
                    colCounter++;
                }
            }

            return rap;
        }

        public static List<RapCalcTableData> GetRAPCalcTableData(string rapFileName)
        {
            //These are used to find request ids in the wrap calc on the RAP Calculator tab and see what quarter it is
            string dataFinder, requestIdFinder, rowFinderRequestor, rowFinderYear;
            int quarterCounter = 1;

            int colStartPoint = 5;
            int rowStartPoint = 49;
            int rowExGratia = 72;
            int colRowHeaders = 3;
            int rowsToScan = 23;
            int colsToScan = 6;

            var headerMap = BuildHeaderMap();
            var tableDataList = new List<RapCalcTableData>();

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("RAP Calculator");

                for(int col = colStartPoint; col <= colStartPoint + colsToScan; col++)
                {
                    var tableData = new RapCalcTableData();

                    tableData.Header = worksheet.Cell(rowStartPoint, col).GetString();
                    for (int row = rowStartPoint; row <= rowStartPoint + rowsToScan; row++)
                    {
                        // Header is in the header column
                        string headerRaw = worksheet.Cell(row, colRowHeaders).GetString();
                        if (string.IsNullOrWhiteSpace(headerRaw)) continue;

                        string headerKey = Normalize(headerRaw);

                        // Value for this record at current col
                        string value = worksheet.Cell(row, col).GetString();

                        if (headerMap.TryGetValue(headerKey, out var setter))
                        {
                            setter(tableData, value);
                        }
                        else
                        {
                            // Optional: handle unknown headers (log, collect, ignore)
                            // Debug.WriteLine($"Unmapped header: '{headerRaw}' at row {row}");
                        }
                    }
                    //Get Ex Gratia cause the header is displaced by one column
                    string headerKeyGratia = Normalize(worksheet.Cell(rowExGratia, colRowHeaders+1).GetString());
                    if (headerMap.TryGetValue(headerKeyGratia, out var setterGratia))
                    {
                        setterGratia(tableData, worksheet.Cell(rowExGratia, col).GetString());
                    }

                    tableDataList.Add(tableData);
                }

            }

            return tableDataList;
        }

        public static PaymentSheetData GetPaymentSheetData(string rapFileName, int quarter)
        {
            int rowPayments = 6;
            int col = 1;

            var paymentData = new PaymentSheetData();

            string quarterStr = $"Quarter {quarter}";

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("Closing");

                //Get Payments Spreadsheet Data from Closing Sheet
                paymentData.requestorName = worksheet.Cell(rowPayments, col).GetString();
                paymentData.requestorAddress = worksheet.Cell(rowPayments, col + 1).GetString();
                paymentData.institution = worksheet.Cell(rowPayments, col + 2).GetString();
                paymentData.amount = worksheet.Cell(rowPayments, col + 3).GetString();
                paymentData.invoiceNumber = worksheet.Cell(rowPayments, col + 4).GetString();
            }

            return paymentData;
        }
        public static RAPTrackerData GetRAPData(string rapFileName, int quarter)
        {
            int rowRAPTracker = 11;
            int col = 1;

            var headerMap = BuildHeaderMap();
            var rapData = new RAPTrackerData();

            string quarterStr = $"Quarter {quarter}";

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("Closing");

                //Get RapTracker Data from Closing sheet
                rapData.fdTicket = worksheet.Cell(rowRAPTracker, col + 1).GetString();
                rapData.agentProcessing = worksheet.Cell(rowRAPTracker, col + 2).GetString();
                rapData.rapDifficultyLevel = worksheet.Cell(rowRAPTracker, col + 3).GetString();
                rapData.contactAttemptsToRapper = worksheet.Cell(rowRAPTracker, col + 4).GetString();
                rapData.reimbursementDocumentsAnalyzed = worksheet.Cell(rowRAPTracker, col + 5).GetString();
                rapData.missingInfoLetterSent = worksheet.Cell(rowRAPTracker, col + 6).GetString();
                rapData.dateOfLastCorrespondence = worksheet.Cell(rowRAPTracker, col + 7).GetString();
                rapData.missingInfoCollected = worksheet.Cell(rowRAPTracker, col + 8).GetString();
                rapData.rapCalcCompleted = worksheet.Cell(rowRAPTracker, col + 9).GetString();
                rapData.reimbursementSubmittedToPAS = worksheet.Cell(rowRAPTracker, col + 10).GetString();
                rapData.checkSentToStudent = worksheet.Cell(rowRAPTracker, col + 11).GetString();
                rapData.amountPaid = worksheet.Cell(rowRAPTracker, col + 12).GetString();
                rapData.ytdPaidAdjustments = worksheet.Cell(rowRAPTracker, col + 13).GetString();
                rapData.monthlyCoveredPayment = worksheet.Cell(rowRAPTracker, col + 14).GetString();
                rapData.monthlyEligibleAmount = worksheet.Cell(rowRAPTracker, col + 15).GetString();
                rapData.monthlyLrapAssistanceAmount = worksheet.Cell(rowRAPTracker, col + 16).GetString();
                rapData.avgMonthlyAssistancePaid = worksheet.Cell(rowRAPTracker, col + 17).GetString();
                rapData.paymentDate = worksheet.Cell(rowRAPTracker, col + 18).GetString();
                rapData.quarterlyPaymentAdjusted = worksheet.Cell(rowRAPTracker, col + 19).GetString();
                rapData.monthsCovered = worksheet.Cell(rowRAPTracker, col + 20).GetString();
                rapData.monthsRequested = worksheet.Cell(rowRAPTracker, col + 21).GetString();
                rapData.annualizedQuarterlyIncome = worksheet.Cell(rowRAPTracker, col + 22).GetString();
                rapData.recordedInTracker = worksheet.Cell(rowRAPTracker, col + 23).GetString();
                rapData.amendedReimbursementProcessed = worksheet.Cell(rowRAPTracker, col + 24).GetString();
                rapData.borrowingProration = worksheet.Cell(rowRAPTracker, col + 25).GetString();
                rapData.incomeProration = worksheet.Cell(rowRAPTracker, col + 26).GetString();
                rapData.timeWorkedProration = worksheet.Cell(rowRAPTracker, col + 27).GetString();
                rapData.monthlyEligibleLoanPaymentsCovered = worksheet.Cell(rowRAPTracker, col + 28).GetString();
                rapData.monthlyIbrEligibleStandard = worksheet.Cell(rowRAPTracker, col + 29).GetString();
                rapData.nonIbrLrapCovered = worksheet.Cell(rowRAPTracker, col + 30).GetString();
                rapData.calculatedIbrPayment = worksheet.Cell(rowRAPTracker, col + 31).GetString();
                rapData.payeOrIbr = worksheet.Cell(rowRAPTracker, col + 32).GetString();
                rapData.householdSize = worksheet.Cell(rowRAPTracker, col + 33).GetString();
                rapData.minimumPaymentIbrPaye = worksheet.Cell(rowRAPTracker, col + 34).GetString();
                rapData.actualPayment = worksheet.Cell(rowRAPTracker, col + 35).GetString();

            }

            return rapData;
        }
        public static ExGratiaClosingData GetExGratiaData(string rapFileName, int quarter)
        {
            int rowExgratia = 16;

            int col = 1;

            var headerMap = BuildHeaderMap();
            var exGratiaData = new ExGratiaClosingData();



            string quarterStr = $"Quarter {quarter}";

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("Closing");

                //Ex Gratia Tracker Data
                exGratiaData.exGratiaPaymentType = worksheet.Cell(rowExgratia, col + 0).GetString();
                exGratiaData.requestorName = worksheet.Cell(rowExgratia, col + 1).GetString();
                exGratiaData.requestorNumber = worksheet.Cell(rowExgratia, col + 2).GetString();
                exGratiaData.institution = worksheet.Cell(rowExgratia, col + 3).GetString();
                exGratiaData.cohort = worksheet.Cell(rowExgratia, col + 4).GetString();
                exGratiaData.quarterCovered = worksheet.Cell(rowExgratia, col + 5).GetString();
                exGratiaData.date = worksheet.Cell(rowExgratia, col + 6).GetString();
                exGratiaData.academicYearsOfLoans = worksheet.Cell(rowExgratia, col + 7).GetString();
                exGratiaData.principalAmountWithIssue = worksheet.Cell(rowExgratia, col + 8).GetString();
                exGratiaData.totalEligiblePrincipal = worksheet.Cell(rowExgratia, col + 9).GetString();
                exGratiaData.totalEligibleMonthlyPayment = worksheet.Cell(rowExgratia, col + 10).GetString();
                exGratiaData.quarterlyEligibleNotBasedOnIdr = worksheet.Cell(rowExgratia, col + 11).GetString();
                exGratiaData.quarterlyEligibleBasedOnAgi = worksheet.Cell(rowExgratia, col + 12).GetString();
                exGratiaData.quarterlyEligibleBasedOnIdrCalc = worksheet.Cell(rowExgratia, col + 13).GetString();
                exGratiaData.borrowingProration = worksheet.Cell(rowExgratia, col + 14).GetString();
                exGratiaData.hoursIncomeProration = worksheet.Cell(rowExgratia, col + 15).GetString();
                exGratiaData.reimbursementForQuarter = worksheet.Cell(rowExgratia, col + 16).GetString();
                exGratiaData.idrPaymentBeforeAdjustment = worksheet.Cell(rowExgratia, col + 17).GetString();
                exGratiaData.timeWorkedProration = worksheet.Cell(rowExgratia, col + 18).GetString();

            }

            return exGratiaData;
        }

        // Utility: normalize headers so matching is resilient to spaces/case/symbols
        private static string Normalize(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            // Trim, collapse spaces, lower-case
            s = s.Trim();

            // Replace multiple spaces with single
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ");

            // Normalize some common symbols/words to help matching
            s = s.Replace("№", "#", StringComparison.Ordinal);
            s = s.Replace("＃", "#", StringComparison.Ordinal);
            s = s.Replace("$", "dollar", StringComparison.OrdinalIgnoreCase);

            // Lower for case-insensitive compare
            s = s.ToLowerInvariant();

            return s;
        }

        // Build the header → setter map once.
        // We match on normalized header strings (lower, trimmed, simplified).
        private static Dictionary<string, Action<RapCalcTableData, string>> BuildHeaderMap()
        {
            var map = new Dictionary<string, Action<RapCalcTableData, string>>(StringComparer.OrdinalIgnoreCase)
            {
                // Use multiple keys for robustness (aliases)
                [Normalize("Request #")] = (t, v) => t.RequestNumber = v,

                [Normalize("# of months requested in quarter")] = (t, v) => t.MonthsRequestedInQuarter = v,
                [Normalize("number of months requested in quarter")] = (t, v) => t.MonthsRequestedInQuarter = v,

                [Normalize("# of months in quarter qualifying")] = (t, v) => t.MonthsInQuarterQualifying = v,
                [Normalize("number of months in quarter qualifying")] = (t, v) => t.MonthsInQuarterQualifying = v,

                [Normalize("Actual Income")] = (t, v) => t.ActualIncome = v,
                [Normalize("Annualized Income")] = (t, v) => t.AnnualizedIncome = v,

                [Normalize("Borrowing Proration")] = (t, v) => t.BorrowingProration = v,
                [Normalize("Income Proration")] = (t, v) => t.IncomeProration = v,
                [Normalize("Time Worked Proration")] = (t, v) => t.TimeWorkedProration = v,

                [Normalize("Eligible Loan Payments")] = (t, v) => t.EligibleLoanPayments = v,
                [Normalize("IDR Eligible Loan Payments")] = (t, v) => t.IdrEligibleLoanPayments = v,

                [Normalize("Eligible Loan Payments YTD")] = (t, v) => t.EligibleLoanPaymentsYtd = v,
                [Normalize("Assistance Due YTD")] = (t, v) => t.AssistanceDueYtd = v,
                [Normalize("Assistance Paid YTD")] = (t, v) => t.AssistancePaidYtd = v,

                [Normalize("Assistance Due Now")] = (t, v) => t.AssistanceDueNow = v,
                [Normalize("Assistance Paid Now")] = (t, v) => t.AssistancePaidNow = v,

                [Normalize("Prior Period Adjustment")] = (t, v) => t.PriorPeriodAdjustment = v,

                [Normalize("Check $")] = (t, v) => t.CheckAmount = v,
                [Normalize("Check $ amount")] = (t, v) => t.CheckAmount = v,
                [Normalize("Check Amount")] = (t, v) => t.CheckAmount = v,

                [Normalize("Date")] = (t, v) => t.Date = v,
                [Normalize("Notes")] = (t, v) => t.Notes = v,
                [Normalize("Time-Worked Ex Gratia:")] = (t, v) => t.Notes = v,
            };

            return map;
        }
    }

    public class RAPTrackerData
    {
        public string fdTicket { get; set; }
        public string agentProcessing { get; set; }
        public string rapDifficultyLevel { get; set; }
        public string contactAttemptsToRapper { get; set; }
        public string reimbursementDocumentsAnalyzed { get; set; }
        public string missingInfoLetterSent { get; set; }
        public string dateOfLastCorrespondence { get; set; }
        public string missingInfoCollected { get; set; }
        public string rapCalcCompleted { get; set; }
        public string reimbursementSubmittedToPAS { get; set; }
        public string checkSentToStudent { get; set; }
        public string amountPaid { get; set; }
        public string ytdPaidAdjustments { get; set; }
        public string monthlyCoveredPayment { get; set; }
        public string monthlyEligibleAmount { get; set; }
        public string monthlyLrapAssistanceAmount { get; set; }
        public string avgMonthlyAssistancePaid { get; set; }
        public string paymentDate { get; set; }
        public string quarterlyPaymentAdjusted { get; set; }
        public string monthsCovered { get; set; }
        public string monthsRequested { get; set; }
        public string annualizedQuarterlyIncome { get; set; }
        public string recordedInTracker { get; set; }
        public string amendedReimbursementProcessed { get; set; }
        public string borrowingProration { get; set; }
        public string incomeProration { get; set; }
        public string timeWorkedProration { get; set; }
        public string monthlyEligibleLoanPaymentsCovered { get; set; }
        public string monthlyIbrEligibleStandard { get; set; }
        public string nonIbrLrapCovered { get; set; }
        public string calculatedIbrPayment { get; set; }
        public string payeOrIbr { get; set; }
        public string householdSize { get; set; }
        public string minimumPaymentIbrPaye { get; set; }
        public string actualPayment { get; set; }
    }
    public class PaymentSheetData
    {
        public string requestorName { get; set; }
        public string requestorAddress { get; set; }
        public string institution { get; set; }
        public string amount { get; set; }
        public string invoiceNumber { get; set; }
    }

    public class ExGratiaClosingData
    {
        public string exGratiaPaymentType { get; set; }
        public string requestorName { get; set; }
        public string requestorNumber { get; set; }
        public string institution { get; set; }
        public string cohort { get; set; }
        public string quarterCovered { get; set; }
        public string date { get; set; }
        public string academicYearsOfLoans { get; set; }
        public string principalAmountWithIssue { get; set; }
        public string totalEligiblePrincipal { get; set; }
        public string totalEligibleMonthlyPayment { get; set; }
        public string quarterlyEligibleNotBasedOnIdr { get; set; }
        public string quarterlyEligibleBasedOnAgi { get; set; }
        public string quarterlyEligibleBasedOnIdrCalc { get; set; }
        public string borrowingProration { get; set; }
        public string hoursIncomeProration { get; set; }
        public string reimbursementForQuarter { get; set; }
        public string idrPaymentBeforeAdjustment { get; set; }
        public string timeWorkedProration { get; set; }
    }
}
