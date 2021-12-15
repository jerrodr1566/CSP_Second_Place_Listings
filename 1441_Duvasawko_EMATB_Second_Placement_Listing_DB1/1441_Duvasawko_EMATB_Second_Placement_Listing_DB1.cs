using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Windows.Markup;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.IO;
using AmcClientLibrary;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1
{
    class _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1
    {
        DataSet dsExcelData = new DataSet();
        DataTable dtAllData = new DataTable();

        string TestModePrefix;
        string DataBasePrefix;
        string SavePath = @"G:\Clients\Duvasawko\EMATB Second Placement\Sent to PARC\";
        string ClientName = "Duvasawko_EMATB_Second_Placement";

        AMC_Functions.GeneralFunctions oGenFun = new AMC_Functions.GeneralFunctions();

        DateTime startDate = DateTime.Today.AddDays(-1);
        DateTime endDate = DateTime.Today.AddDays(-1);

        public _1441_Duvasawko_EMATB_Second_Placement_Listing_DB1(bool inTestMode, bool inisSQL, string inDataBase)
        {
            string fileName = "EMATB_Cancel Second Placement " + DateTime.Now.ToString("MMddyyyy");
            string fileExt = ".xlsx";

            if (inTestMode)
            {
                TestModePrefix = "TEST_";
            }
            else
            {
                TestModePrefix = "";
            }

            DataBasePrefix = inDataBase;

            // call the CSPs!
            AMC_Functions.CSPProcessListing oCSP_Duvo = new AMC_Functions.CSPProcessListing("DSEMATB-G", "Duvasawko EMATB", inisSQL, startDate, endDate, SavePath, $"EMATB_Cancel Second Placement {DateTime.Today.ToString("yyyyMMdd")}.xlsx", inTestMode);

            // get the data out of it, as well as the excel sheet - in case there is additional formatting needed
            dsExcelData = oCSP_Duvo.returnExcelData();
            dtAllData = oCSP_Duvo.returnAllData();

            CopyDataOverAsStrings();
            EvaluateForAdditionalSteps(inTestMode, inisSQL);

            // export data, if there is anyhting to actually export
            if (dtAllData.Rows.Count >= 1)
            {
                ExportData(SavePath, fileName, fileExt, "ashleys@americollect.com", ClientName, inTestMode);
            }
        }

        private void CopyDataOverAsStrings()
        {
            DataTable dtTemp = new DataTable();

            foreach (DataColumn dc in dtAllData.Columns)
            {
                dtTemp.Columns.Add(dc.ColumnName);
            }

            // transfer the data over, and format it correctly based off of the type - mainly for dates
            foreach (DataRow dr in dtAllData.Rows)
            {
                string[] dataToAdd = new string[dtAllData.Columns.Count];
                long currentArrayCounter = 0;

                foreach (DataColumn dc in dtAllData.Columns)
                {
                    switch (dc.DataType.ToString())
                    {
                        case "System.DateTime":

                            if (dc.ColumnName == "Last Pay Date" && dr[dc.ColumnName].ToString() != string.Empty)
                            {
                                if (Convert.ToDateTime(dr[dc.ColumnName].ToString()) == Convert.ToDateTime("01/01/1900").Date)
                                {
                                    dataToAdd[currentArrayCounter] = string.Empty;
                                }
                                else
                                {
                                    // function as normal
                                    dataToAdd[currentArrayCounter] = Convert.ToDateTime(dr[dc.ColumnName].ToString()).ToString("MM/dd/yyyy");
                                }
                            }
                            else
                            {
                                // function as normal
                                dataToAdd[currentArrayCounter] = Convert.ToDateTime(dr[dc.ColumnName].ToString()).ToString("MM/dd/yyyy");
                            }

                            break;
                        case "System.Double":
                            dataToAdd[currentArrayCounter] = Convert.ToDouble(dr[dc.ColumnName].ToString()).ToString("0.00");
                            break;

                        case "System.Decimal":
                            if (dc.ColumnName == "Last Pay Amount" && dr[dc.ColumnName].ToString() != string.Empty)
                            {
                                if (Convert.ToDecimal(dr[dc.ColumnName].ToString()) == 0)
                                {
                                    dataToAdd[currentArrayCounter] = string.Empty;
                                }
                                else
                                {
                                    // proceed as normal
                                    dataToAdd[currentArrayCounter] = Convert.ToDecimal(dr[dc.ColumnName].ToString()).ToString("0.00");
                                }
                            }
                            else
                            {
                                // proceed as normal
                                dataToAdd[currentArrayCounter] = Convert.ToDecimal(dr[dc.ColumnName].ToString()).ToString("0.00");
                            }
                            break;

                        default:
                            dataToAdd[currentArrayCounter] = dr[dc.ColumnName].ToString();
                            break;

                    }

                    currentArrayCounter++;
                }


                dtTemp.Rows.Add(dataToAdd);
            }

            // clear the table itself, and transfer the data over
            dtAllData = new DataTable();

            // transfer the columns back in, and then the data
            foreach (DataColumn dc in dtTemp.Columns)
            {
                dtAllData.Columns.Add(dc.ColumnName);
            }

            foreach (DataRow dr in dtTemp.Rows)
            {
                object[] dataToAdd = dr.ItemArray.Cast<object>().ToArray();

                dtAllData.Rows.Add(dataToAdd);
            }


        }

        /// <summary>
        /// Check for any additional steps - total payments, adjustments, etc
        /// </summary>
        /// <param name="inTestMode"></param>
        /// <param name="inisSQL"></param>
        private void EvaluateForAdditionalSteps(bool inTestMode, bool inisSQL)
        {
            DataTable dt = dsExcelData.Tables[0];

            bool additionalLogicNeeded = false;
            bool additionalLogicColFound = false;
            string additionallogicText = "";

            Dictionary<string, string> dict_addLogic = new Dictionary<string, string>();

            // make sure the additional column is actually found, if not, then this step isn't needed
            foreach (DataColumn dc in dt.Columns)
            {
                if (dc.ColumnName.ToUpper().Contains("ADDITIONALLOGIC"))
                {
                    additionalLogicColFound = true;
                    additionallogicText = dc.ColumnName;
                }
            }

            if (additionalLogicColFound)
            {
                // now iterate through the rows in there and see if there are any that require additional work to make formatted correctly
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr[additionallogicText].ToString() != string.Empty)
                    {
                        additionalLogicNeeded = true;

                        if (!dict_addLogic.ContainsKey(dr["ClientHeader"].ToString()))
                        {
                            dict_addLogic.Add(dr["ClientHeader"].ToString(), dr["ClientHeader"].ToString());
                        }
                    }
                }

            }

            Console.WriteLine("Evaluating additional logic portion..");

            if (additionalLogicNeeded)
            {
                // connect to the database, pending what is set for it
                string connectionString;
                AMC_Functions.DetermineDSNFile oDSN = new AMC_Functions.DetermineDSNFile();

                if (inisSQL)
                {
                    connectionString = oDSN.getDSNFile("jerrodr", "Reporting DB1", false);
                }
                else
                {
                    connectionString = inTestMode ? oDSN.getDSNFile("jerrodr", "Training DB1", false) : oDSN.getDSNFile("jerrodr", "DB1", false);

                    using (OdbcConnection con = new OdbcConnection(connectionString))
                    {
                        con.Open();

                        long iLoop = 0;

                        ///////////**********************************////////////////////////
                        ///this part will be partially customized per client
                        foreach (DataRow dr in dtAllData.Rows)
                        {

                            string percentage = (Convert.ToDouble(iLoop) / Convert.ToDouble(dtAllData.Rows.Count)).ToString("P");

                            string updateString = "Currently on " + iLoop + " of " + dtAllData.Rows.Count + " (" + percentage + " completed)";

                            string backupData = "";

                            // update the total percentage
                            if (iLoop >= 1)
                            {
                                // create the string to re-write blanks, and then update
                                backupData = new string('\b', updateString.Length);
                                Console.Write(backupData);

                                // now, write the updated
                                Console.Write(updateString);
                            }
                            else
                            {
                                // no need to re-write, so just write the normal text
                                Console.Write(updateString);
                            }


                            const string select_TotalPay = @"SELECT sum(baamount) as 'TotalPayments' 
from PUB.tranmstr 
JOIN PUB.balances on PUB.balances.baserial = tmtserial 
JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber 
WHERE amanumber = ? and tmrcptcode IN ('A', 'C', 'D', 'M', 'O', 'R', 'W', 'X') and tmtrancode = 'C' WITH (NOLOCK)";

                            const string select_TotalIns = @"SELECT sum(baamount) as 'TotalInsPayments' 
from PUB.tranmstr
JOIN PUB.balances on PUB.balances.baserial = tmtserial
JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber
WHERE amanumber = ? and tmrcptcode IN('N', 'S', 'Z') and tmtrancode = 'C' WITH (NOLOCK)";

                            const string select_TotalAdj = @"SELECT sum(baamount) as 'TotalAdjusts' 
from PUB.tranmstr 
JOIN PUB.balances on PUB.balances.baserial = tmtserial 
JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber 
WHERE amanumber = ? and tmtrancode = 'A' WITH (NOLOCK)";

                            decimal TotalInsurancePay = 0;
                            decimal TotalPatientPay = 0;

                            string[] patientName;

                            // check if the patient is the guarantor, if not, need to clear some data out
                            bool ptEqualGuar = false;

                            // patient name isn't fixed yet, so it's using ADATA:AH1
                            if ((dr["Guarantor Last Name"].ToString() + ", " + dr["Guarantor First Name"].ToString()).ToUpper() == dr["Patient First Name"].ToString().ToUpper() ||
                                (dr["Guarantor Last Name"].ToString() + "," + dr["Guarantor First Name"].ToString()).ToUpper() == dr["Patient First Name"].ToString().ToUpper())
                            {
                                ptEqualGuar = true;
                            }

                            foreach (string sValue in dict_addLogic.Values)
                            {
                                switch (sValue.ToUpper())
                                {
                                    case "FLEX TAG":
                                        if (dr[sValue].ToString().ToUpper() == "TRUE")
                                        {
                                            dr[sValue] = "Y";
                                        }
                                        else
                                        {
                                            dr[sValue] = string.Empty;
                                        }
                                        break;

                                    case "TOTAL INSURANCE PAYMENTS":
                                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalIns, con))
                                        {
                                            SelectCMD.Parameters.Add("@acct", OdbcType.VarChar).Value = dr["xxamanumber"].ToString();

                                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                                            {
                                                while (Reader.Read())
                                                {
                                                    if (Reader["TotalInsPayments"].ToString() != string.Empty)
                                                    {
                                                        dr[sValue] = Convert.ToDecimal(Reader["TotalInsPayments"].ToString()).ToString("0.00");

                                                        TotalInsurancePay = Convert.ToDecimal(Reader["TotalInsPayments"].ToString());
                                                    }
                                                    else
                                                    {
                                                        dr[sValue] = Convert.ToDecimal(0).ToString("0.00");

                                                        TotalInsurancePay = Convert.ToDecimal(0);
                                                    }

                                                }
                                            }
                                        }

                                        break;

                                    case "TOTAL PATIENT PAYMENTS":
                                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalPay, con))
                                        {
                                            SelectCMD.Parameters.Add("@acct", OdbcType.VarChar).Value = dr["xxamanumber"].ToString();

                                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                                            {
                                                while (Reader.Read())
                                                {
                                                    if (Reader["TotalPayments"].ToString() != string.Empty)
                                                    {
                                                        dr[sValue] = Convert.ToDecimal(Reader["TotalPayments"].ToString()).ToString("0.00");

                                                        TotalPatientPay = Convert.ToDecimal(Reader["TotalPayments"].ToString());
                                                    }
                                                    else
                                                    {
                                                        dr[sValue] = Convert.ToDecimal(0).ToString("0.00");

                                                        TotalPatientPay = Convert.ToDecimal(0);
                                                    }

                                                }
                                            }
                                        }
                                        break;

                                    case "TOTAL PAYMENTS":
                                        // add the previous 2 together
                                        dr[sValue] = (TotalInsurancePay + TotalPatientPay).ToString("0.00");
                                        break;
                                    case "TOTAL ADJUSTMENTS":
                                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalAdj, con))
                                        {
                                            SelectCMD.Parameters.Add("@acct", OdbcType.VarChar).Value = dr["xxamanumber"].ToString();

                                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                                            {
                                                while (Reader.Read())
                                                {
                                                    if (Reader["TotalAdjusts"].ToString() != string.Empty)
                                                    {
                                                        // updated 5-27-2020 JAR, need to inverse it, as per AAS 5-27-2020
                                                        dr[sValue] = (Convert.ToDecimal(Reader["TotalAdjusts"].ToString()) * -1).ToString("0.00");
                                                        //dr[sValue] = Convert.ToDecimal(Reader["TotalAdjusts"].ToString()).ToString("0.00");
                                                    }
                                                    else
                                                    {
                                                        dr[sValue] = Convert.ToDecimal(0).ToString("0.00");
                                                    }

                                                }
                                            }
                                        }
                                        break;

                                    case "PATIENT FIRST NAME":
                                        // split and get first name
                                        patientName = dr[sValue].ToString().Split(',');

                                        dr[sValue] = patientName[patientName.Length - 1].Trim();

                                        break;
                                    case "PATIENT LAST NAME":
                                        // split and get last name
                                        patientName = dr[sValue].ToString().Split(',');

                                        dr[sValue] = patientName[0].Trim();

                                        break;
                                    case "PATIENT ADDRESS1":
                                    case "PATIENT ADDRESS2":
                                    case "PATIENT CITY":
                                    case "PATIENT ST":
                                    case "PATIENT ZIP":
                                        // only keep if the patient is the same as the guar
                                        if (!ptEqualGuar)
                                        {
                                            dr[sValue] = string.Empty;
                                        }
                                        break;
                                    case "PATIENT EMPLOYER NAME":
                                        // first, check if patient is guarantor, if so, then evaluate further,otherwise blank
                                        if (!ptEqualGuar)
                                        {
                                            dr[sValue] = string.Empty;
                                        }
                                        else
                                        {
                                            // if the first letter is #, replace that
                                            if (dr[sValue].ToString().StartsWith("#"))
                                            {
                                                dr[sValue] = dr[sValue].ToString().TrimStart('#');
                                            }

                                            // if it has "NON ATTY REP", clear that
                                            if (dr[sValue].ToString().ToUpper().Contains("NON ATTY REP"))
                                            {
                                                dr[sValue] = string.Empty;
                                            }
                                        }
                                        break;
                                }

                            }

                            // separate special cases
                            if (dr["Employer"].ToString() != string.Empty)
                            {
                                // if the first letter is #, replace that
                                if (dr["Employer"].ToString().StartsWith("#"))
                                {
                                    dr["Employer"] = dr["Employer"].ToString().TrimStart('#');
                                }

                                // if it has "NON ATTY REP", clear that
                                if (dr["Employer"].ToString().ToUpper().Contains("NON ATTY REP"))
                                {
                                    dr["Employer"] = string.Empty;
                                }
                            }

                            if (dr["Guarantor DOB"].ToString() != string.Empty)
                            {
                                try
                                {
                                    dr["Guarantor DOB"] = DateTime.ParseExact(dr["Guarantor DOB"].ToString(), "MMddyyyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                                }
                                catch
                                {
                                    dr["Guarantor DOB"] = Convert.ToDateTime(dr["Guarantor DOB"].ToString()).ToString("MM/dd/yyyy");
                                }
                            }

                            // if the first letter is #, replace that
                            if (dr["Employer"].ToString().StartsWith("#"))
                            {
                                dr["Employer"] = dr["Employer"].ToString().TrimStart('#');
                            }

                            // if it has "NON ATTY REP", clear that
                            if (dr["Employer"].ToString().ToUpper().Contains("NON ATTY REP"))
                            {
                                dr["Employer"] = string.Empty;
                            }

                            iLoop++;
                        }


                        ///////////**********************************////////////////////////

                        con.Close();
                    }
                }
            }

            bool colToRemove = false;
            string colNameExact = "";

            // see if the table has the xxamanumber column in it, if so, remove that
            foreach (DataColumn dc in dtAllData.Columns)
            {
                if (dc.ColumnName.ToUpper() == "XXAMANUMBER")
                {
                    colToRemove = true;
                    colNameExact = dc.ColumnName;
                }
            }

            if (colToRemove)
            {
                dtAllData.Columns.Remove(colNameExact);
            }

        }

        /// <summary>
        /// Save the data into the format expected
        /// </summary>
        /// <param name="inFilePath"></param>
        /// <param name="inFileName"></param>
        /// <param name="inFileExt"></param>
        /// <param name="inEmailAddress"></param>
        /// <param name="inClientName"></param>
        private void ExportData(string inFilePath, string inFileName, string inFileExt, string inEmailAddress, string inClientName, bool inTestMode)
        {
            string fullFileName = inFilePath + inFileName;

            string sharefileAddress = "https://parcassets.sharefile.com/r-rdb418dc19c14b339";

            switch (inFileExt.ToUpper())
            {
                case ".XLSX":
                    // dataset detail
                    // add the data tables to a dataset and export that way
                    DataSet ds = new DataSet("ExcelExport");

                    // get the table names back in case they cleaned up at all during run time
                    dtAllData.TableName = "SecondPlacements";

                    ds.Tables.Add(dtAllData);

                    ExportDataSetToExcel(ds, inFilePath, inFileName + inFileExt);

                    break;
                case ".CSV":
                    // similar export, but save a csv
                    ToCSV(dtAllData, inFilePath + inFileName + inFileExt);
                    break;
                default:
                    // anything else-  basically text file
                    AmcClientLibrary.JAR_NewBusinessFunctions.WriteNBFile(inTestMode, DataBasePrefix, dtAllData, inFilePath, inFilePath, ClientName, TestModePrefix, inFileName);
                    break;
            }

            oGenFun.SendEmail(inEmailAddress, "", ClientName + " CSP Listing", $"Hello!</br></br>Here is the listing for the CSP accounts.</br></br><a href= '{fullFileName}'>{fullFileName}</a></br></br>Please place on the PARC Sharefile: <a href ='{sharefileAddress}'>{sharefileAddress}</a></br></br>Email Shari & Chris letting them know the file is available for pickup.</br></br>sirsay@parcassets.net</br>cconway@parcassets.net</br></br>", "macro@americollect.com", "", false, false);

        }

        /// <summary>
        /// Place on the FTP if allowed
        /// </summary>
        /// <param name="inAllow"></param>
        private void PlaceOnFTP(bool inAllow)
        {

        }


        private void ExportDataSetToExcel(DataSet ds, string inFilePath, string inFileName)
        {
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            // turned visibilty on originally to make sure was working as expected
            //excelApp.Visible = true;

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                ADODB.Recordset rs = new ADODB.Recordset();

                rs = AmcLibrary_KLV.ExcelFunctions.ConvertToRecordset(table);

                // add the column headers
                int intColIndex;
                for (intColIndex = 0; intColIndex <= rs.Fields.Count - 1; intColIndex++)
                {
                    excelWorkSheet.Cells[1, intColIndex + 1] = rs.Fields[intColIndex].Name;
                }


                excelWorkSheet.Range["A2"].CopyFromRecordset(AmcLibrary_KLV.ExcelFunctions.ConvertToRecordset(table));

                // added in autofit 4-15-2020
                excelWorkSheet.Columns.AutoFit();
            }

            // check if there is a sheet1 in there, since it usually gets added by default, but only if there is more than 1 sheet
            if (excelWorkBook.Worksheets.Count > 1)
            {
                foreach (Excel.Worksheet worksheet in excelWorkBook.Worksheets)
                {
                    if (worksheet.Name.ToUpper() == "SHEET1")
                    {
                        worksheet.Delete();
                    }
                }
            }

            excelWorkBook.SaveAs(inFilePath + inFileName);


            excelWorkBook.Close();
            excelApp.Quit();

        }

        public void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}
