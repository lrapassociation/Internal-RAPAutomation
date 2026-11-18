using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace FileDownload
{
    public class GetExcelLists
    {
        public static List<int> GetRAPTracker()
        {
            string filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\2B_Rap tracker_jtest2.xlsm";
            int idColNumber = 7; // col where IDs live 

            var ids = new List<int>();

            var stream = ExtractWorksheet(filePath, "Request Tracker");
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheet(1);

                //check for the id in the row we most likely think it is
                var col = worksheet.Column(idColNumber);

                foreach (var cell in col.CellsUsed())
                {
                    if (int.TryParse(cell.GetString(), out int id))
                    {
                        ids.Add(id);
                    }
                }
            }

            return ids;
        }

        //private List<int> GetRAPTracker()
        //    {
        //        string filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\2B_Rap tracker_jtest.xlsm";
        //        int idColNumber = 7; // col where IDs live 

        //        var ids = new List<int>();

        //        var app = new Excel.Application(); 
        //        var wb = app.Workbooks.Open(filePath); 
        //        var ws = wb.Sheets["Request Tracker"];

        //        int lastrow = ws.Cells[ws.Rows.Count, "C"].End(Excel.XlDirection.xlUp).Row;

        //        for (int row = 1; row <= lastrow; row++)
        //        {
        //            var value = ws.Cells[row, 1].Value2;

        //            if (int.TryParse((string?)value, out int id))
        //            {
        //                ids.Add(id);
        //            }
        //        }

        //        wb.Close(true);
        //        app.Quit();

        //        return ids;
        //    }

        public static List<string> GetExGratia()
        {
            string filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Tracker for Ex Gratia Payments_jtesting.xlsx";
            int idColNumber = 3, quarterColNumber = 6; // col where IDs live 

            var ids = new List<string>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("Payments Made");

                int firstRow = 15;
                int lastRow = worksheet.Column("C").LastCellUsed().Address.RowNumber;

                //check for the id in the row we most likely think it is
                var idCol = worksheet.Column(idColNumber);
                var quarterCol = worksheet.Column(quarterColNumber);


                for (int row = firstRow; row <= lastRow; row++)
                {
                    var requestNumber = worksheet.Cell(row, idColNumber).GetString();
                    var quarter = worksheet.Cell(row, quarterColNumber).GetString();

                    if (int.TryParse(requestNumber, out int id))
                    {
                        string outputId = string.Concat(requestNumber, quarter);
                        ids.Add(outputId);
                    }
                }

            }

            return ids;
        }

        public static List<int> GetPaymentSpreadsheet()
        {
            string filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Payment Information_jtesting.xlsx";
            int idColNumber = 5; // col where IDs live most of the time

            var ids = new List<int>();

            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    //check for the id in the row we most likely think it is
                    var col = sheet.Column(idColNumber);
                    string firstCell = col.FirstCell().GetString();
                    if (string.Equals(firstCell, "Invoice #", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var cell in col.CellsUsed())
                        {
                            string rid = cell.GetString().Split(' ').Last();
                            if (int.TryParse(rid, out int id))
                            {
                                ids.Add(id);
                            }
                        }
                    }
                    else
                    {
                        //grab first column to find where request number lives
                        var row = sheet.Row(1);
                        int colCount = 1;
                        foreach (var rowCell in row.CellsUsed())
                        {
                            if (string.Equals(rowCell.GetString(), "Invoice #", StringComparison.OrdinalIgnoreCase))
                            {
                                col = sheet.Column(colCount);
                                foreach (var colCell in col.CellsUsed())
                                {
                                    string rid = colCell.GetString().Split(' ').Last();
                                    if (int.TryParse(rid, out int id))
                                    {
                                        ids.Add(id);
                                    }
                                }
                            }
                            colCount++;
                        }
                    }
                }
            }

            return ids;
        }

        public static List<int> GetWeeklyRAPApprovals()
        {
            string filePath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Weekly RAP Approvals_jtesting.xlsx";
            int idRowNumber = 13; // row where IDs live most of the time

            var ids = new List<int>();

            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    //check for the id in the row we most likely think it is
                    var row = sheet.Row(idRowNumber);
                    string firstCell = row.FirstCell().GetString();
                    if (string.Equals(firstCell, "Overall Request Number", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var cell in row.CellsUsed())
                        {
                            if (int.TryParse(cell.GetString(), out int id))
                            {
                                ids.Add(id);
                            }
                        }
                    }
                    else
                    {
                        //grab first column to find where request number lives
                        var col = sheet.Column(1);
                        int rowCount = 1;
                        foreach (var colCell in col.CellsUsed())
                        {
                            if (string.Equals(colCell.GetString(), "Overall Request Number", StringComparison.OrdinalIgnoreCase))
                            {
                                row = sheet.Row(rowCount);
                                foreach (var rowCell in row.CellsUsed())
                                {
                                    if (int.TryParse(rowCell.GetString(), out int id))
                                    {
                                        ids.Add(id);
                                    }
                                }
                            }
                            rowCount++;
                        }
                    }
                }
            }

            return ids;
        }

        public static MemoryStream ExtractWorksheet(string path, string sheetName)
        {
            var ms = new MemoryStream();

            using (var source = SpreadsheetDocument.Open(path, false))
            using (var dest = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var sourceWb = source.WorkbookPart;
                var destWb = dest.AddWorkbookPart();
                destWb.Workbook = new Workbook();
                var sheets = destWb.Workbook.AppendChild(new Sheets());

                // Find sheet by name
                var sheet = sourceWb.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found");

                // Copy worksheet part
                var sourceSheetPart = (WorksheetPart)sourceWb.GetPartById(sheet.Id);
                var destSheetPart = destWb.AddPart(sourceSheetPart);

                sheets.Append(new Sheet
                {
                    Id = destWb.GetIdOfPart(destSheetPart),
                    SheetId = 1,
                    Name = sheetName
                });

                // Copy shared strings
                if (sourceWb.SharedStringTablePart != null)
                    destWb.AddPart(sourceWb.SharedStringTablePart);

                // Copy styles
                if (sourceWb.WorkbookStylesPart != null)
                    destWb.AddPart(sourceWb.WorkbookStylesPart);

                // Copy theme
                if (sourceWb.ThemePart != null)
                    destWb.AddPart(sourceWb.ThemePart);
            }

            ms.Position = 0;
            return ms;
        }
    }
}
