using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileDownload
{
    public class RapExcelHelper
    {

        public static List<List<string>> GetRapExcelTableForDataGrid(string rapFileName,string sheetName, int colStartPoint, int colEndPoint, int rowStartPoint,int rowEndPoint)
        {
            var tableDataList = new List<List<string>>();

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet(sheetName); 

                for (int row = rowStartPoint; row <= rowEndPoint; row++)
                {
                    var rowData = new List<string>();
                    for (int col = colStartPoint; col <= colEndPoint; col++)
                    {
                        var cellContents = worksheet.Cell(row, col);

                        if (cellContents.DataType == XLDataType.DateTime)
                        {
                            var dt = cellContents.GetDateTime();
                            rowData.Add(dt.ToString("d"));
                        }
                        else if (cellContents.DataType == XLDataType.Number)
                        {
                            var dec = cellContents.GetDouble();
                            rowData.Add(dec.ToString("0.00"));
                        }
                        else rowData.Add(worksheet.Cell(row, col).GetString());

                    }
                    tableDataList.Add(rowData);
                }

            }

            return tableDataList;
        }

        public static List<List<string>> GetRapExcelTableForDataGridSwitchRowsAndColumns(string rapFileName, string sheetName, int colStartPoint, int colEndPoint, int rowStartPoint, int rowEndPoint)
        {
            var tableDataList = new List<List<string>>();

            using (var workbook = new XLWorkbook(rapFileName))
            {
                var worksheet = workbook.Worksheet(sheetName);

                for (int col = colStartPoint; col <= colEndPoint; col++)
                {
                    var rowData = new List<string>();
                    for (int row = rowStartPoint; row <= rowEndPoint; row++)
                    {
                        var cellContents = worksheet.Cell(row, col);

                        if (cellContents.DataType == XLDataType.DateTime)
                        { 
                            var dt = cellContents.GetDateTime();
                            rowData.Add(dt.ToString("d"));
                        }
                        else if (cellContents.DataType == XLDataType.Number)
                        {
                            var dec = cellContents.GetDouble();
                            rowData.Add(dec.ToString("0.00"));
                        }
                        else rowData.Add(worksheet.Cell(row, col).GetString());
                    }
                    tableDataList.Add(rowData);
                }
            }

            return tableDataList;
        }

        public static void ExportRangeAsPdf(
                string inputFilePath,
                string sheetName,
                string rangeAddress,
                string outputPdfPath,
                bool fitToOnePage = true)
        {
            Excel.Application app = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Sheets wss = null;
            Excel.Worksheet ws = null;

            try
            {
                app = new Excel.Application();

                app.Visible = false;
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Open(inputFilePath);

                ws = string.IsNullOrWhiteSpace(sheetName)
                    ? (Excel.Worksheet)wb.Worksheets[1]
                    : (Excel.Worksheet)wb.Worksheets[sheetName];

                // Ensure only the target sheet is active/selected
                ws.Activate();
                ws.Select(false);

                // --- Configure page setup to limit export to your selection
                ws.PageSetup.PrintArea = rangeAddress;     // <-- critical: only this range gets printed
                ws.PageSetup.CenterHorizontally = true;
                ws.PageSetup.CenterVertically = false;

                if (fitToOnePage)
                {
                    ws.PageSetup.Zoom = false;             // must disable Zoom before FitToPages
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 1;
                }
                else
                {
                    ws.PageSetup.Zoom = 100;
                }

                // Optional margins/orientation
                ws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                ws.PageSetup.LeftMargin = app.InchesToPoints(0.25f);
                ws.PageSetup.RightMargin = app.InchesToPoints(0.25f);
                ws.PageSetup.TopMargin = app.InchesToPoints(0.5f);
                ws.PageSetup.BottomMargin = app.InchesToPoints(0.5f);

                // Export the workbook—Excel will honor the sheet's PrintArea if IgnorePrintAreas=false
                ws.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: outputPdfPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,   // <-- DO NOT ignore print areas
                    From: Type.Missing,
                    To: Type.Missing,
                    OpenAfterPublish: false);

            }
            finally
            {
                // Clean up COM objects in reverse order
                if (ws != null) Marshal.FinalReleaseComObject(ws);
                if (wb != null)
                {
                    wb.Close(SaveChanges: false);
                    Marshal.FinalReleaseComObject(wb);
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                }
                ws = null; wb = null; app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

    }

}
