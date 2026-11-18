using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TechArch;

namespace FileDownload
{
    internal class InsertSQL
    {
        public static void InsertPaymentSQL(string configFile, string connString, PaymentSheetData data)
        {
            ADODataAccess objADO = new ADODataAccess();
            objADO.connect(connString);

            //insert statement here
            StringBuilder sbSQL = new StringBuilder(@"Insert Into RAP_Payments(columns)
            values()");


            objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);

        }

        public static void InsertExGratiaSQL(string configFile, string connString, ExGratiaClosingData data)
        {
            ADODataAccess objADO = new ADODataAccess();
            objADO.connect(connString);

            //insert statement here
            StringBuilder sbSQL = new StringBuilder(@"Insert Into RAP_Payments(columns)
            values()");


            objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);
        }

        internal static void InsertRAPTrackerSQL(string configFile, string connString, RAPTrackerData rdata)
        {
            ADODataAccess objADO = new ADODataAccess();
            objADO.connect(connString);

            //insert statement here
            StringBuilder sbSQL = new StringBuilder(@"Insert Into RAP_Payments(columns)
            values()");


            objADO.ExecuteQuery(sbSQL.ToString(), System.Data.CommandType.Text, null);
        }
    }



}
