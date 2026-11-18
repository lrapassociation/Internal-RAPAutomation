using AzureStorageExplorerApp;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections;
using System.IO;
using System.Security.AccessControl;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using TechArch;

namespace FileDownload
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string _configFile;
        string _conn_string;
        List<Document> _listDocs = null;
        List<Employment> _listEmps = null;
        private int result;
        List<Document> _listDownloadDocs = null;


        public MainWindow()
        {
            InitializeComponent();

            Service objTechArch = new Service();
            Security objTechSecurity = new Security();

            _configFile = @"C:\TEAM LRAP Dropbox\4A_RAPS\1_Pending\Request Group Document Download\Config\Configuration.txt";
            _conn_string = objTechSecurity.revealConnectionString(objTechArch.getConfigVar("CONNECTION_STRING", _configFile));

        }


        private void BoxNID_TextChanged(object sender, TextChangedEventArgs e)
        {
            Service objTechArch = new Service();
            Security objTechSecurity = new Security();

            string connection_string = objTechSecurity.revealConnectionString(objTechArch.getConfigVar("CONNECTION_STRING", _configFile));

            string rapId = boxRequestId.Text.ToString().Trim();

            bool allDigits = rapId.All(char.IsDigit);

            if (int.TryParse(rapId, out result))
            {
                _listDocs = getDocs(_configFile, connection_string, result.ToString());
                _listEmps = getEmployment(_configFile, connection_string, result.ToString());

                if (_listDocs.Count > 0)
                {
                    boxStudentId.Text = _listDocs[0].StudentId;
                    boxLrapId.Text = _listDocs[0].LrapId;
                    boxStudentName.Text = _listDocs[0].Name;
                    boxStudentAddress.Text = _listDocs[0].Address;
                    boxStudentHousehold.Text = _listDocs[0].Household;
                    boxThirdParty.Text = _listDocs[0].ThirdParty;

                    EmpGrid.ItemsSource = _listEmps;
                    DocDataGrid.ItemsSource = _listDocs;
                    btnDownloadDocument.IsEnabled = true;
                    btnDownloadAllDocument.IsEnabled = true;
                }
                else
                {
                    btnDownloadDocument.IsEnabled = false;
                    btnDownloadAllDocument.IsEnabled = false;

                    DocDataGrid.ItemsSource = null;
                    EmpGrid.ItemsSource = null;

                    boxStudentId.Text = "";
                    boxLrapId.Text = "";
                    boxStudentName.Text = "";
                    boxStudentAddress.Text = "";
                    boxThirdParty.Text = "";

                    boxEmpName.Text = "";
                    boxTitle.Text = "";
                    boxHours.Text = "";
                    boxIncomeType.Text = "";
                    boxCity.Text = "";
                    boxState.Text = "";
                    boxStartDate.Text = "";
                    boxEndDate.Text = "";
                    boxOwnBusiness.Text = "";
                    boxAnnualIncome.Text = "";
                }
                
            }
        }

        public List<Document> getDocs(string configFile, string connection_string, string rapId)
        {
            ArrayList rows = null;

            try
            {
                ADODataAccess objADO = new ADODataAccess();
                objADO.connect(connection_string);

                List<Document> output = new List<Document>();

                StringBuilder sbSQL = new StringBuilder("select rd.ID as 'Doc_id',s.FIRST_NAME + ' ' + s.LAST_NAME as 'Name', lrg.STUDENT_ID, TYPE,BLOB_FILENAME as 'Blob_Filename', Modified" +
                    " ,COALESCE(ci.street_address_1, '') + CASE WHEN ci.street_address_2 IS NOT NULL THEN ', ' + ci.street_address_2 ELSE '' END + ', ' + COALESCE(ci.city, '') + ', ' + COALESCE(ci.state_province, '') + ', ' + COALESCE(ci.postal_code, '') AS Address " +
                    " ,ci.HOUSEHOLD_SIZE " +
                    " ,case when (select IS_3RD_PARTY_AUTHORIZATION_ACKNOWLEDGED FROM LSP_LRAP_ACCOUNT_INFO WHERE STUDENT_ID = LRG.STUDENT_ID) = 1 then 'Yes' else 'No' end " +
                    " ,s.LRAP_STUDENT_ID " +
                    " from LSP_LRAP_REQUEST_GROUP lrg " +
                    " join RAP_DOCUMENT rd on lrg.ID = rd.REQUEST_GROUP_ID "+
                    " join student s on lrg.STUDENT_ID = s.STUDENT_ID" +
                    " join LSP_LRAP_STUDENT_CONTACT_INFO ci on lrg.ID = ci.REQUEST_GROUP_ID " +
                    " Where lrg.ID = " + rapId +
                    " Order by Modified desc");

                rows = objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);


                foreach (object[] row in rows)
                {
                    int i = 0;
                    Document r = new Document();
                    foreach (object column in row)
                    {
                        if (i == 0) r.DocId = Int32.Parse(column.ToString());
                        else if (i == 1) r.Name = column.ToString();
                        else if (i == 2) r.StudentId = column.ToString();
                        else if (i == 3) r.Type = column.ToString();
                        else if (i == 4) r.BlobFilename = column.ToString();
                        else if (i == 5) r.Modified = column.ToString();
                        else if (i == 6) r.Address = column.ToString();
                        else if (i == 7) r.Household = column.ToString();
                        else if (i == 8) r.ThirdParty = column.ToString();
                        else if (i == 9) r.LrapId = column.ToString();

                        i++;
                    }
                    output.Add(r);
                }

                return output;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Employment> getEmployment(string configFile, string connection_string, string rapId)
        {
            ArrayList rows = null;

            try
            {
                ADODataAccess objADO = new ADODataAccess();
                objADO.connect(connection_string);

                List<Employment> output = new List<Employment>();

                StringBuilder sbSQL = new StringBuilder("select EMPLOYER_NAME, JOB_TITLE, HOURS_WORKED_PER_WEEK, case when SALARIED is null then 'N/A' when SALARIED = 1 then 'Salary' when SALARIED = 0 then 'Hourly' else 'N/A' end as 'Income_Type' " +
                    ",CITY, STATE_PROVINCE, START_DATE, ISNULL(END_DATE,'N/A') AS END_DATE, FAMILY_OWNED, ID, 'Participant', EXPECTED_ANNUAL_INCOME" +
                    " from LSP_LRAP_STUDENT_EMPLOYMENT " +
                    " Where REQUEST_GROUP_ID = " + rapId +
                    " union all " +
                    "select EMPLOYER_NAME, JOB_TITLE, HOURS_WORKED_PER_WEEK, case when SALARIED is null then 'N/A' when SALARIED = 1 then 'Salary' when SALARIED = 0 then 'Hourly' else 'N/A' end as 'Income_Type' " +
                    ",CITY, STATE_PROVINCE, START_DATE, ISNULL(END_DATE,'N/A') AS END_DATE, FAMILY_OWNED, SPOUSE_DEGREE_ID, 'Spouse', EXPECTED_ANNUAL_INCOME" +
                    " from LSP_LRAP_SPOUSE_EMPLOYMENT " +
                    " Where REQUEST_GROUP_ID = " + rapId
                    );

                rows = objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);


                foreach (object[] row in rows)
                {
                    int i = 0;
                    Employment r = new Employment();
                    foreach (object column in row)
                    {
                        if (i == 0)  r.EmployerName = column.ToString();
                        else if (i == 1) r.JobTitle = column.ToString();
                        else if (i == 2) r.WeeklyHours = column.ToString();
                        else if (i == 3) r.Salaried = column.ToString();
                        else if (i == 4) r.City = column.ToString();
                        else if (i == 5) r.State = column.ToString();
                        else if (i == 6) r.StartDate = column.ToString();
                        else if (i == 7) r.EndDate = column.ToString();
                        else if (i == 8) r.FamilyOwned = column.ToString();
                        else if (i == 9) r.JobId = column.ToString();
                        else if (i == 10) r.IsSpouse = column.ToString();
                        else if (i == 11) r.ExpectedAnnualIncome = column.ToString();

                        i++;
                    }
                    output.Add(r);
                }

                return output;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void BtnDownloadDocument_Click(object sender, RoutedEventArgs e)
        {
            CStorage objStorage1 = new AzureStorageExplorerApp.CStorage();
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\TEAM LRAP Dropbox";
            dialog.IsFolderPicker = true;
            string container_name = "rap-document/";

            _listDownloadDocs = new List<Document>();



            _listDownloadDocs = DocDataGrid.SelectedItems.Cast<Document>().ToList();



            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string tempfilename;
                foreach (Document downloadDoc in _listDownloadDocs)
                {
                    tempfilename = downloadDoc.BlobFilename;


                    objStorage1.DownloadBlobAsync(false, tempfilename, dialog.FileName + "\\" + Path.GetFileName(tempfilename), container_name);
                    Thread.Sleep(300);
                }

            }
        }

        private void BtnDownloadAllDocument_Click(object sender, RoutedEventArgs e)
        {
            CStorage objStorage1 = new AzureStorageExplorerApp.CStorage();
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\TEAM LRAP Dropbox";
            dialog.IsFolderPicker = true;
            string container_name = "rap-document/";

            _listDownloadDocs = new List<Document>();


            //if (MessageBox.Show("Do you want to Download All of this students Documents?",
            //    "Download",
            //    MessageBoxButton.YesNo,
            //    MessageBoxImage.Question) == MessageBoxResult.Yes)
            //{
                _listDownloadDocs = _listDocs;

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    string tempfilename;

                    string folderpath = dialog.FileName + "\\" + boxRequestId.Text.ToString().Trim();
                    string zipPath = folderpath + ".zip";

                    Directory.CreateDirectory(folderpath);

                    foreach (Document downloadDoc in _listDownloadDocs)
                    {
                        tempfilename = downloadDoc.BlobFilename;

                        objStorage1.DownloadBlobAsync(false, tempfilename, folderpath + "\\" + Path.GetFileName(tempfilename), container_name);
                        Thread.Sleep(300);
                    }

                    //ZipFile.CreateFromDirectory(folderpath, zipPath);
                }
            //}
        }

        private void EmpGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (e.AddedCells.Count == 1)
            {
                Employment emp = (Employment)e.AddedCells[0].Item;

                boxEmpName.Text = emp.EmployerName;
                boxTitle.Text = emp.JobTitle;
                boxHours.Text = emp.WeeklyHours;
                boxIncomeType.Text = emp.Salaried;
                boxCity.Text = emp.City;
                boxState.Text = emp.State;
                boxStartDate.Text = emp.StartDate;
                boxEndDate.Text = emp.EndDate;
                boxOwnBusiness.Text = emp.FamilyOwned;
                boxAnnualIncome.Text = emp.ExpectedAnnualIncome;

            }

        }

        private void btnApprovalProcess_Click(object sender, RoutedEventArgs e)
        {
            ApprovalWindow window = new ApprovalWindow(_configFile, _conn_string);
            window.Show();
        }

        //private void btnUpdateStatus_Click(object sender, RoutedEventArgs e)
        //{
        //    string statusToUpdate = boxStatus.Text.ToString();
        //    string requestGroupId = boxRequestId.Text.ToString();
        //    ArrayList rows = null;

        //    Service objTechArch = new Service();
        //    Security objTechSecurity = new Security();

        //    string connection_string = objTechSecurity.revealConnectionString(objTechArch.getConfigVar("CONNECTION_STRING", _configFile));

        //    try
        //    {
        //        ADODataAccess objADO = new ADODataAccess();
        //        objADO.connect(connection_string);

        //        List<Document> output = new List<Document>();

        //        StringBuilder sbSQL = new StringBuilder("update LSP_REQUEST_GROUP set status = "+ statusToUpdate +" where id = " + requestGroupId);

        //        rows = objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);


        //        foreach (object[] row in rows)
        //        {
        //            int i = 0;
        //            Document r = new Document();
        //            foreach (object column in row)
        //            {
        //                if (i == 0) r.DocId = Int32.Parse(column.ToString());
        //                else if (i == 1) r.Name = column.ToString();
        //                else if (i == 2) r.StudentId = column.ToString();
        //                else if (i == 3) r.Type = column.ToString();
        //                else if (i == 4) r.BlobFilename = column.ToString();
        //                else if (i == 5) r.Modified = column.ToString();
        //                else if (i == 6) r.Address = column.ToString();

        //                i++;
        //            }
        //            output.Add(r);
        //        }

        //        return output;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}
    }

    public class Document
    {
        public int DocId { get; set; }
        public string Name { get; set; }
        public string StudentId { get; set; }
        public string LrapId { get; set; }
        public string Type { get; set; }
        public string BlobFilename { get; set; }
        public string Modified { get; set; }
        public string Address { get; set; }
        public string EmployerName { get; set; }
        public string JobTitle { get; set; }
        public string WeeklyHours { get; set; }
        public string Salaried { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string FamilyOwned { get; set; }
        public string? Household { get; internal set; }
        public string ThirdParty { get; set; }

        public Document() { }
    }

    public class Employment
    {
        public string EmployerName { get; set; }
        public string JobTitle { get; set; }
        public string WeeklyHours { get; set; }
        public string Salaried { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string FamilyOwned { get; set; }
        public string JobId { get; set; }
        public string IsSpouse { get; set; }
        public string ExpectedAnnualIncome { get; set; }


        public Employment() { }
    }
}