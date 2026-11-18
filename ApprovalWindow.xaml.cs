using DocumentFormat.OpenXml;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using TechArch;

namespace FileDownload
{
    /// <summary>
    /// Interaction logic for ApprovalWindow.xaml
    /// </summary>
    public partial class ApprovalWindow : Window
    {
        List<int> WeeklyRAPApprovalIds, PaymentSpreadsheetIds, RapTrackerIds;
        List<string> ExGratiaIds;

        RapData currentRAP;
        List<RapCalcTableData> currentCalcTableData;

        bool ListsReady = false;
        bool ListError = false;
        bool StudentDataReady = false;
        bool StudentDataError = false;

        string _configFile;
        string _connString;

        string CurrentRapFile = "";

        private Logging objLogging = null;

        public ApprovalWindow(string configFile, string conn_string)
        {
            _configFile = configFile;
            _connString = conn_string;

            DateTime date = DateTime.Now;

            string s = @"C:\TEAM LRAP Dropbox\SECURITY_APPLICATION\YourSLI\Logs\" + date.ToString("yyyyMMdd-HHmmss") + ".txt";

            objLogging = new Logging(s);
            objLogging.start();
            objLogging.logMessage("Approval Window Opened");

            InitializeComponent();

            CreateExcelDictionaries();
            

            string path = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\test wraps";

            string[] filepaths = Directory.GetFiles(path,"*.xlsm");
            List<RapFile> files = new();

            foreach (string file in filepaths)
            {
                files.Add(new RapFile { FilePath = file, FileName = Path.GetFileName(file) });
            }

            ExcelDataGrid.ItemsSource = files;

            //initial tracker
            //copy paste historical data
            //copy paste rap tracker
            //copy payments spreadsheet
            //copy ex gratia
            //copy weekly rap approvals


        }

        private async void CreateExcelDictionaries()
        {
            DataLoadingText.Visibility = Visibility.Visible;
            

            //TODO: add error checking to each call
            await Task.Run(() =>
            {
                SetLoadingValues(DataLoadingText, LoadingBar, "Loading Weekly RAP Approvals", 10);
                WeeklyRAPApprovalIds = GetExcelLists.GetWeeklyRAPApprovals();
                SetLoadingValues(DataLoadingText, LoadingBar, "Loading Payment Spreadsheet", 30);
                PaymentSpreadsheetIds = GetExcelLists.GetPaymentSpreadsheet();
                SetLoadingValues(DataLoadingText, LoadingBar, "Loading Ex Gratia", 50);
                ExGratiaIds = GetExcelLists.GetExGratia();
                SetLoadingValues(DataLoadingText, LoadingBar, "Loading RAP Tracker", 80);
                RapTrackerIds = GetExcelLists.GetRAPTracker();
                SetLoadingValues(DataLoadingText, LoadingBar, "Done", 100);

            });

            bool hasWeekly = WeeklyRAPApprovalIds.Any();
            bool hasPayment = PaymentSpreadsheetIds.Any();
            bool hasExGratia = ExGratiaIds.Any();
            bool hasTracker = RapTrackerIds.Any();
            bool allHaveData =
                WeeklyRAPApprovalIds.Any() &&
                PaymentSpreadsheetIds.Any() &&
                ExGratiaIds.Any() &&
                RapTrackerIds.Any();

            if (allHaveData)
            {
                DataLoadingText.Visibility = Visibility.Collapsed;
                DataReadyText.Visibility = Visibility.Visible;
                ListsReady = true;
            }
            else 
            { 
                DataLoadingText.Visibility = Visibility.Collapsed;
                DataErrorText.Visibility = Visibility.Visible;
                ListError = true;
            }
            
        }

        public void SetLoadingValues(TextBlock textToUpdate, ProgressBar barToUpdate, string message, int progress)
        {
            //Need to see what thread we are currently on: If UI thread then all good if not need to force it since we are running from within await Task.Run
            if (Dispatcher.CheckAccess())
            {
                textToUpdate.Text = message;
                barToUpdate.Value = progress;
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    textToUpdate.Text = message;
                    barToUpdate.Value = progress;                        
                });

            }
        }

        

        private async void BeginApprovalProcess(string rapFileName)
        {
            CurrentRapFile = rapFileName;
            RapData rap = new();
            bool hasCalc = false,hasSummary = false,hasTracker = false,hasPayments = false,hasExGratia = false, pdfGenerated = false;
            var tableDataList = new List<RapCalcTableData>();

            await Task.Run(() =>
            {
                ApprovalWindowHelperFunctions.ClearPdfs();

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Getting RAP Calculator", 0);

                rap = GetRAPExcelData.GetRAPCalcData(rapFileName);

                tableDataList = GetRAPExcelData.GetRAPCalcTableData(rapFileName);

                //RapGrid.ItemsSource = tableDataList;

                var displayRapCalculatorGrid = RapExcelHelper.GetRapExcelTableForDataGrid(rapFileName, "RAP Calculator", 3, 10, 49, 72);

                Dispatcher.Invoke(() => BindListOfListsToDataGrid(RapCalculatorGrid, displayRapCalculatorGrid));

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Getting RAP Summary", 20);
                var displayRapSummaryGrid = RapExcelHelper.GetRapExcelTableForDataGrid(rapFileName, "RAP Summary", 1, 5, 1, 46);

                Dispatcher.Invoke(() => BindListOfListsToDataGrid(RapSummaryGrid, displayRapSummaryGrid, 110));

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Getting Closing Payments", 40);
                var displayPaymentsGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(rapFileName, "Closing", 1, 5, 5, 6);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(PaymentsGrid, displayPaymentsGrid, 90));

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Getting Closing RAP Tracker", 60);
                var displayRAPTrackerGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(rapFileName, "Closing", 1, 36, 10, 11);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(RapTrackerGrid, displayRAPTrackerGrid));

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Getting Closing RAP Tracker", 80);
                var displayExGratiaGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(rapFileName, "Closing", 1, 19, 15, 16);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(ExGratiaGrid, displayExGratiaGrid, 90));

                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Generating RAP Approval PDF", 90);
                string outputpath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Pdfs\" + Path.GetFileNameWithoutExtension(CurrentRapFile);
                GenerateExcelPDF(CurrentRapFile, "RAP Approval", "A1:I55", outputpath);
                
                hasCalc = displayRapCalculatorGrid.Any();
                hasSummary = displayRapSummaryGrid.Any();
                hasPayments = displayPaymentsGrid.Any();
                hasTracker = displayRAPTrackerGrid.Any();
                hasExGratia = displayExGratiaGrid.Any();
                pdfGenerated = File.Exists(outputpath + ".pdf");
            });

            bool allHaveData = hasCalc && hasSummary && hasTracker && hasPayments && hasExGratia;

            DoesToolHaveAllData(allHaveData, pdfGenerated);

            //Wait until the lists have all the data they need to check
            await ApprovalWindowHelperFunctions.WaitUntilAsync(() => ListsReady);

            //Checking data from the RAP calc against each individual excel sheet to see if they exist within them
            DataContext = ApprovalWindowHelperFunctions.CheckHistory(rap, WeeklyRAPApprovalIds, PaymentSpreadsheetIds, RapTrackerIds, ExGratiaIds);

            //Set Global Variables to be accessible to other functions
            currentRAP = rap;
            currentCalcTableData = tableDataList;



        }



        private void DoesToolHaveAllData(bool allHaveData, bool pdfGenerated)
        {

            if (allHaveData)
            {
                StudentBtnStack.Visibility = Visibility.Visible;
                StudentDataReady = true;

                //Check if the rap approval pdf was successfully created, if not need to disable the button
                //Done this way because the pdf doc needs the excel to be closed while the other tabs dont
                if (pdfGenerated)
                {
                    SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "", 100);
                    StudentDataLoadingText.Visibility = Visibility.Collapsed;
                    StudentDataReadyText.Visibility = Visibility.Visible;
                    btnRAPApproval.IsEnabled = true;
                }
                else
                {
                    btnRAPApproval.IsEnabled = false;
                }

            }
            else
            {
                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "", 0);
                StudentDataLoadingText.Visibility = Visibility.Collapsed;
                StudentDataErrorText.Visibility = Visibility.Visible;
                StudentDataError = true;
            }
        }

        private async void GenerateExcelPDF(string currentRapFile, string v1, string v2, string outputpath, bool FileWasOpen = false)
        {
            try
            {
                RapExcelHelper.ExportRangeAsPdf(CurrentRapFile, "RAP Approval", "A1:I55", outputpath);

                bool pdfGenerated = File.Exists(outputpath + ".pdf");
                if (FileWasOpen && pdfGenerated) Dispatcher.Invoke(() => DoesToolHaveAllData(true, pdfGenerated));
                else if(FileWasOpen) throw new Exception("PDF Not Generated Error");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                SetLoadingValues(StudentDataLoadingText, StudentLoadingBar, "Please Close Excel to Generate Approval PDF", 90);
                await Task.Run(() =>
                {
                    //Wait 10 seconds and try again
                    Task.Delay(10000);
                    GenerateExcelPDF(currentRapFile, v1, v2, outputpath, true);
                });
            }

            
        }

        public void BindListOfListsToDataGrid(DataGrid grid, List<List<string>> myRows, int columnWidth = 70)
        {
            grid.AutoGenerateColumns = false;
            grid.Columns.Clear();

            if (myRows == null || myRows.Count == 0)
            {
                grid.ItemsSource = null;
                return;
            }

            var headers = myRows[0] ?? new List<string>();
            int colCount = 0;
            int colCount2 = 1;
            foreach (var r in myRows)
                colCount = Math.Max(colCount, r?.Count ?? 0);

            for (int c = 0; c < colCount; c++)
            {
                string headerText = (c < headers.Count && !string.IsNullOrWhiteSpace(headers[c]))
                    ? headers[c]
                    : $"Column{c + 1}";

                var column = new DataGridTextColumn
                {
                    Header = headerText,
                    Binding = new System.Windows.Data.Binding($"[{c}]") // bind to indexer of List<string>
                    {
                        TargetNullValue = string.Empty,
                        FallbackValue = string.Empty
                    },
                    Width = c == 0 ? DataGridLength.Auto : columnWidth,                    
                }; 

                grid.Columns.Add(column);
            }

            // Use remaining rows as data (skip headers row)
            var data = myRows.Skip(1).ToList();
            grid.ItemsSource = data;
        }

        private void btnRAPCalc_Click(object sender, RoutedEventArgs e)
        {
            SetApprovalWindowVisibility(RapCalculatorStack);
        }
        private void btnRAPSummary_Click(object sender, RoutedEventArgs e)
        {
            SetApprovalWindowVisibility(RapSummaryStack);
        }

        private void btnRAPClosing_Click(object sender, RoutedEventArgs e)
        {
            SetApprovalWindowVisibility(ClosingStack);
        }

        private async void btnRAPApproval_Click(object sender, RoutedEventArgs e)
        {
            //Generate PDF for RAP Approval page
            string outputPdfPath = @"C:\TEAM LRAP Dropbox\6_TECHDEV\1_DASHBOARD TECHDEV\Working doc\James Files\Wrap Calc Development\Approval Process Automation\Pdfs\" + Path.GetFileNameWithoutExtension(CurrentRapFile) + ".pdf";
            //string outputPdfPath = @"C:\Users\James Lefaiver\Documents\Weekly RAP Approvals.pdf";

            await PdfWebView.EnsureCoreWebView2Async();
            PdfWebView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;

            //PdfWebView.Source = new Uri("https://www.microsoft.com");
            PdfWebView.Source = new Uri(outputPdfPath);

            SetApprovalWindowVisibility(RapApprovalStack);
        }

        private void SetApprovalWindowVisibility(StackPanel inputStack = null)
        {
            RapCalculatorStack.Visibility = Visibility.Collapsed;
            RapSummaryStack.Visibility = Visibility.Collapsed;
            ClosingStack.Visibility = Visibility.Collapsed;
            RapApprovalStack.Visibility = Visibility.Collapsed;

            if (inputStack is not null) inputStack.Visibility = Visibility.Visible;

        }

        private void btnHistoricalInsert_Click(object sender, RoutedEventArgs e)
        {
            List<int> quarters = GetQuartersFromUI();

            if (quarters.Count == 0) throw new Exception("No Quarters Found");

            InsertExcelData.InsertHistoricalData(CurrentRapFile, currentRAP, currentCalcTableData, quarters);
        }
        private void btnPaymentsInsert_Click(object sender, RoutedEventArgs e)
        {
            List<int> quarters = GetQuartersFromUI();

            if (quarters.Count == 0) throw new Exception("No Quarters Found");

            foreach (var quarter in quarters) 
            {
                PaymentSheetData pdata = GetRAPExcelData.GetPaymentSheetData(CurrentRapFile, quarter);

                InsertSQL.InsertPaymentSQL(_configFile, _connString, pdata);

                InsertExcelData.InsertPaymetSpreadsheet(CurrentRapFile, currentRAP, pdata, quarter);
            }
        }
        private void btnRAPTrackerInsert_Click(object sender, RoutedEventArgs e)
        {
            List<int> quarters = GetQuartersFromUI();

            if (quarters.Count == 0) throw new Exception("No Quarters Found");

            foreach (var quarter in quarters)
            {
                RAPTrackerData rdata = GetRAPExcelData.GetRAPData(CurrentRapFile, quarter);

                InsertSQL.InsertRAPTrackerSQL(_configFile, _connString, rdata);

                //InsertExcelData.InsertRAPTrackerSpreadsheet(CurrentRapFile, currentRAP, rdata, quarter);
            }
        }

        private void btnExGratia_Click(object sender, RoutedEventArgs e)
        {
            List<int> quarters = GetQuartersFromUI();

            if (quarters.Count == 0) throw new Exception("No Quarters Found");

            foreach (var quarter in quarters)
            {
                ExGratiaClosingData edata = GetRAPExcelData.GetExGratiaData(CurrentRapFile, quarter);

                InsertSQL.InsertExGratiaSQL(_configFile, _connString, edata);

                InsertExcelData.InsertExGratia(CurrentRapFile, currentRAP, edata, quarter);
            }
        }

        private void ExcelGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (e.AddedCells.Count == 1)
            {
                SetApprovalWindowVisibility();
                StudentDataErrorText.Visibility = Visibility.Collapsed;
                StudentDataReadyText.Visibility = Visibility.Collapsed;
                StudentDataLoadingText.Visibility = Visibility.Visible;

                StudentDataStack.Visibility = Visibility.Visible;
                StudentBtnStack.Visibility = Visibility.Collapsed;

                RapFile rapFile = (RapFile)e.AddedCells[0].Item;

                BeginApprovalProcess(rapFile.FilePath);
            }

        }

        private List<int> GetQuartersFromUI()
        {
            var quarters = new List<int>();
            if (Int32.TryParse(Quarter1Box.Text.ToString(), out int quarter1)) quarters.Add(quarter1);
            if (Int32.TryParse(Quarter2Box.Text.ToString(), out int quarter2)) quarters.Add(quarter2);
            if (Int32.TryParse(Quarter3Box.Text.ToString(), out int quarter3)) quarters.Add(quarter3);
            if (Int32.TryParse(Quarter4Box.Text.ToString(), out int quarter4)) quarters.Add(quarter4);

            return quarters;
        }
        private void Quarter1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<string> quarters = new List<string>{"", "1", "2", "3", "4" };

            if (e.AddedItems.Count > 0)
            {
                string app = ((ComboBoxItem)e.AddedItems[0]).Content?.ToString();
                quarters.Remove(app);
                Quarter2Box.ItemsSource = quarters;
                Quarter2Box.Visibility = Visibility.Visible;

                InsertExcelData.SetQuarter(CurrentRapFile, int.Parse(((ComboBoxItem)e.AddedItems[0]).Content?.ToString()));

                var displayPaymentsGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(CurrentRapFile, "Closing", 1, 5, 5, 6);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(PaymentsGrid, displayPaymentsGrid, 90));

                var displayRAPTrackerGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(CurrentRapFile, "Closing", 1, 36, 10, 11);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(RapTrackerGrid, displayRAPTrackerGrid));

                var displayExGratiaGrid = RapExcelHelper.GetRapExcelTableForDataGridSwitchRowsAndColumns(CurrentRapFile, "Closing", 1, 19, 15, 16);
                Dispatcher.Invoke(() => BindListOfListsToDataGrid(ExGratiaGrid, displayExGratiaGrid, 90));
            }
        }

        private void Quarter2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<string> quarters = new List<string> { "", "1", "2", "3", "4" };

            if (e.AddedItems.Count > 0 && e.AddedItems[0].ToString() != "")
            {
                string q2 = e.AddedItems[0].ToString();
                quarters.Remove(q2);
                string q1 = Quarter1Box.Text.ToString();
                quarters.Remove(q1);
                Quarter3Box.ItemsSource = quarters;
                Quarter3Box.Visibility = Visibility.Visible;
            }
            else if (e.AddedItems[0].ToString() == "")
            {
                Quarter3Box.ItemsSource = null;
                Quarter3Box.Visibility = Visibility.Collapsed;
            }
        }
        private void Quarter3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<string> quarters = new List<string> { "", "1", "2", "3", "4" };

            if (e.AddedItems.Count > 0 && e.AddedItems[0].ToString() != "")
            {
                string q3 = e.AddedItems[0].ToString();
                quarters.Remove(q3);
                string q2 = Quarter2Box.SelectedItem.ToString();
                quarters.Remove(q2);
                string q1 = Quarter1Box.Text.ToString();
                quarters.Remove(q1);
                Quarter4Box.ItemsSource = quarters;
                Quarter4Box.Visibility = Visibility.Visible;
            }
            else if (e.AddedItems[0].ToString() == "")
            {
                Quarter4Box.ItemsSource = null;
                Quarter4Box.Visibility = Visibility.Collapsed;
            }
        }
    }
    
    
}
