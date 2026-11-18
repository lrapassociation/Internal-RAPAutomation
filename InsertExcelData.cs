using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileDownload
{
    public class InsertExcelData
    {

        public static void SetQuarter(string rapFileName, int quarter)
        {
            using (var workbook = new XLWorkbook(rapFileName))
            {

                var worksheet = workbook.Worksheet("Closing");

                worksheet.Cell(1, 1).Value = $"Quarter {quarter}";

                workbook.Save();
            }
        }

        public static void InsertHistoricalData(string rapFileName, RapData rap, List<RapCalcTableData> tableData, List<int> quarters)
        {
            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet("Historical and Analytics");


                foreach(var q in quarters)
                {
                    string valueToFind = $"Q{q}-{rap.year % 100:D2}";
                    bool quarterFound = false;
                    string finder;
                    int quarterRow = 2;
                    int colCounter = 1;
                    while (!quarterFound)
                    {
                        finder = worksheet.Cell(quarterRow, colCounter).GetString();

                        if (finder.Equals(valueToFind))
                        {
                            quarterFound = true;
                            int requestRow = quarterRow + 2;

                            foreach (var data in tableData)
                            {
                                if (data.Header.Equals(valueToFind))
                                {
                                    worksheet.Cell(requestRow, colCounter).Value = Int32.Parse(data.RequestNumber);
                                    worksheet.Cell(requestRow +2, colCounter).Value = decimal.Parse(data.ActualIncome).ToString("C2");
                                    worksheet.Cell(requestRow +3, colCounter).Value = decimal.Parse(data.AnnualizedIncome).ToString("C2");
                                    worksheet.Cell(requestRow +4, colCounter).Value = decimal.Parse(data.BorrowingProration).ToString("P");
                                    worksheet.Cell(requestRow +5, colCounter).Value = decimal.Parse(data.IncomeProration).ToString("P");
                                    worksheet.Cell(requestRow +6, colCounter).Value = decimal.Parse(data.TimeWorkedProration).ToString("P");
                                    worksheet.Cell(requestRow +7, colCounter).Value = decimal.Parse(data.EligibleLoanPayments).ToString("C2");
                                    worksheet.Cell(requestRow +8, colCounter).Value = decimal.Parse(data.IdrEligibleLoanPayments).ToString("C2");
                                    worksheet.Cell(requestRow +9, colCounter).Value = decimal.Parse(data.EligibleLoanPaymentsYtd).ToString("C2");
                                    worksheet.Cell(requestRow +10, colCounter).Value = decimal.Parse(data.AssistanceDueYtd).ToString("C2");
                                    worksheet.Cell(requestRow +11, colCounter).Value = decimal.Parse(data.AssistancePaidYtd).ToString("C2");
                                    worksheet.Cell(requestRow +12, colCounter).Value = decimal.Parse(data.AssistanceDueNow).ToString("C2");
                                    worksheet.Cell(requestRow +13, colCounter).Value = decimal.Parse(data.AssistancePaidNow).ToString("C2");
                                    worksheet.Cell(requestRow +14, colCounter).Value = decimal.Parse(data.PriorPeriodAdjustment).ToString("C2");

                                }
                            }
                            
                        }
                        colCounter++;
                    }
                }
                workbook.Save();
            }
        }

        //TODO
        public static void InsertPaymetSpreadsheet(string paymentFilePath, RapData rap, PaymentSheetData data, int quarter)
        {
            paymentFilePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Payment Information_jtesting.xlsx";

            using (var workbook = new XLWorkbook(paymentFilePath))
            {
                string sheetName = DateTime.Today.ToString("Mddyyyy");

                IXLWorksheet worksheet;

                if (workbook.Worksheets.Contains(sheetName))
                {
                    worksheet = workbook.Worksheet(sheetName);
                }
                else
                {
                    var template = workbook.Worksheet("1162026");

                    // Copy it
                    worksheet = template.CopyTo(sheetName);

                    worksheet.Range(2, 1, 15, 5).Clear(XLClearOptions.Contents);
                }

                int startRow = 2;
                int maxDataRow = 15;

                worksheet.Rows(startRow, 16).Unhide();

                int insertRow = -1;

                for (int row = startRow; row <= maxDataRow; row++)
                {
                    bool isEmpty = true;

                    for (int col = 1; col <= 5; col++) 
                    {
                        if (!worksheet.Cell(row, col).IsEmpty())
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        insertRow = row;
                        break;
                    }
                }

                if (insertRow != -1)
                {
                    worksheet.Cell(insertRow, 1).Value = data.requestorName;
                    worksheet.Cell(insertRow, 2).Value = data.requestorAddress;
                    worksheet.Cell(insertRow, 3).Value = data.institution;
                    worksheet.Cell(insertRow, 4).Value = decimal.Parse(data.amount);
                    worksheet.Cell(insertRow, 4).Style.NumberFormat.Format = "$#,##0.00";
                    worksheet.Cell(insertRow, 5).Value = data.invoiceNumber;

                    if (insertRow + 1 <= maxDataRow)
                    {
                        worksheet.Rows(insertRow + 1, maxDataRow).Hide();
                    }
                }
                else
                {
                    // Optional: handle case where no empty row is found
                    throw new Exception("No available rows between 2 and 15.");
                }

                workbook.Save();
            }
        }

        public static void InsertRAPTracker(string paymentFilePath, RapData rap, List<RapCalcTableData> tableData, List<int> quarters)
        {
            using (var workbook = new XLWorkbook(paymentFilePath))
            {
                var worksheet = workbook.Worksheet("Historical and Analytics");


                foreach (var q in quarters)
                {
                    string valueToFind = $"Q{q}-{rap.year % 100:D2}";
                    bool quarterFound = false;
                    string finder;
                    int quarterRow = 2;
                    int colCounter = 1;
                    while (!quarterFound)
                    {
                        finder = worksheet.Cell(quarterRow, colCounter).GetString();

                        if (finder.Equals(valueToFind))
                        {
                            quarterFound = true;
                            int requestRow = quarterRow + 2;

                            foreach (var data in tableData)
                            {
                                if (data.Header.Equals(valueToFind))
                                {

                                }
                            }

                        }
                        colCounter++;
                    }
                }
                workbook.Save();
            }
        }

        public static void InsertExGratia(string filePath, RapData rap, ExGratiaClosingData data, int quarter)
        {
            filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Tracker for Ex Gratia Payments_jtesting.xlsx";

            using (var workbook = new XLWorkbook(filePath))
            {

                IXLWorksheet worksheet = workbook.Worksheet("Payments Made");

                int startRow = 2;
                int maxDataRow = 15;

                worksheet.Rows(startRow, 16).Unhide();

                int insertRow = -1;

                for (int row = startRow; row <= maxDataRow; row++)
                {
                    bool isEmpty = true;

                    for (int col = 1; col <= 5; col++)
                    {
                        if (!worksheet.Cell(row, col).IsEmpty())
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        insertRow = row;
                        break;
                    }
                }

                if (insertRow != -1)
                {
                    worksheet.Cell(insertRow, 1).Value = data.requestorName;


                    if (insertRow + 1 <= maxDataRow)
                    {
                        worksheet.Rows(insertRow + 1, maxDataRow).Hide();
                    }
                }
                else
                {
                    // Optional: handle case where no empty row is found
                    throw new Exception("No available rows between 2 and 15.");
                }

                workbook.Save();
            }
        }
    }
}
