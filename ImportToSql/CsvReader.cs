using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using ExcelDataReader;
using System.IO;
using System.Data;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using ExcelApp = Microsoft.Office.Interop.Excel;


namespace ImportToSql
{
    class CsvReader
    {
        public static bool GetDataTabletFromCSVFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DataRow myIncompleteDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            int IncompleteColumn = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = false;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        if (datacolumn.ColumnName.Contains("Cleared/Open Items Symbol"))
                        {
                            datacolumn.ColumnName = "Symbol";
                        }
                        else
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        }
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Document_Date") || datacolumn.ColumnName.Contains("Posting_Date"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");
                        else if (datacolumn.ColumnName.Contains("Amount_in_Local_Currency"))
                        {
                            datacolumn.DataType = System.Type.GetType("System.Decimal");
                        }
                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        
                        if (fieldData[0].StartsWith("@") || fieldData[0].StartsWith("\"@"))
                        {
                            myDataRow = csvData.NewRow();
                            IncompleteColumn = 0;
                            //Making empty value as null
                            for (int i = 0; i < fieldData.Length; i++)
                            {
                                IncompleteColumn = i;
                                if (i == 7 && fieldData[i] != "")
                                {
                                    if (fieldData[i].Contains("-"))
                                        fieldData[i] = "-" + fieldData[i].Replace("-", "").Replace(",00", "00").Replace("\"", "");
                                    else
                                        fieldData[i] = fieldData[i].Replace(",00", "00");
                                    myDataRow[i] = fieldData[i].Replace(",,", "");
                                }
                                else if (fieldData[i] != "")
                                {
                                    myDataRow[i] = fieldData[i];     // Added later .Replace(",,", "");
                                }
                                else if (fieldData[i] == "")
                                {
                                    if (i == 14)
                                        myDataRow[i] = DBNull.Value;
                                    else
                                        myDataRow[i] = null;
                                }
                            }
                            myDataRow["Date_Created"] = dtnow;
                            myIncompleteDataRow = myDataRow;
                            csvData.Rows.Add(myDataRow);
                            RowCount++;
                        }
                        else if (fieldData[0].StartsWith(",,"))
                        {
                            myDataRow = csvData.Rows[csvData.Rows.Count - 1];
                            myDataRow[IncompleteColumn] = myDataRow[IncompleteColumn] + fieldData[0];
                            csvData.AcceptChanges();
                        }
                        else 
                        {
                            myDataRow = csvData.Rows[csvData.Rows.Count - 1];
                            myDataRow[IncompleteColumn] = myDataRow[IncompleteColumn] + fieldData[0];

                            for (int i = 1; i < fieldData.Length; i++)
                            {
                                if (i + IncompleteColumn == 7 && fieldData[i] != "")
                                {
                                    if (fieldData[i].Contains("-"))
                                        fieldData[i] = "-" + fieldData[i].Replace("-", "").Replace(",00", "00").Replace("\"", "");
                                    else
                                        fieldData[i] = fieldData[i].Replace(",00", "00");
                                    myDataRow[i] = fieldData[i].Replace(",,", "");
                                }
                                else if (fieldData[i] != "")
                                {
                                    myDataRow[i + IncompleteColumn] = fieldData[i];     // Added later.Replace(",,", "");
                                }
                                else if (fieldData[i] == "")
                                {
                                    if (i + IncompleteColumn == 14)
                                        myDataRow[i + IncompleteColumn] = DBNull.Value;
                                    else
                                        myDataRow[i + IncompleteColumn] = null;
                                }
                            }
                            csvData.AcceptChanges();
                        }
                    }
                }
                InsertIntoSQLServer(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }
        
        public static void InsertIntoSQLServer(DataTable dt, ErrorLog oErrorLog)
        {
            string tableName = ConfigurationManager.AppSettings["FinancetableName"];
            string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

            oErrorLog.WriteErrorLog("Connecting Database");
            SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
            bulkcopy.DestinationTableName = tableName;
            SqlConnection con = new SqlConnection(ssqlconnectionstring);
            try
            {
                con.Open();
                bulkcopy.ColumnMappings.Add("Symbol", "Symbol");
                bulkcopy.ColumnMappings.Add("Assignment", "Assignment");
                bulkcopy.ColumnMappings.Add("Document_Number", "Document_Number");
                bulkcopy.ColumnMappings.Add("Business_Area", "Business_Area");
                bulkcopy.ColumnMappings.Add("Document_type", "Document_type");
                bulkcopy.ColumnMappings.Add("Document_Date", "Document_Date");
                bulkcopy.ColumnMappings.Add("Posting_Key", "Posting_Key");
                bulkcopy.ColumnMappings.Add("Amount_in_Local_Currency", "Amount_in_Local_Currency");
                bulkcopy.ColumnMappings.Add("Local_Currency", "Local_Currency");
                bulkcopy.ColumnMappings.Add("Tax_Code", "Tax_Code");
                bulkcopy.ColumnMappings.Add("Clearing_Document", "Clearing_Document");
                bulkcopy.ColumnMappings.Add("Text", "Text");
                bulkcopy.ColumnMappings.Add("Asset", "Asset");
                bulkcopy.ColumnMappings.Add("Order", "Order");
                bulkcopy.ColumnMappings.Add("Posting_Date", "Posting_Date");
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Line_item", "Line_item");
                bulkcopy.ColumnMappings.Add("Fiscal_Year", "Fiscal_Year");
                bulkcopy.ColumnMappings.Add("Account_type", "Account_type");
                bulkcopy.ColumnMappings.Add("Account", "Account");
                bulkcopy.ColumnMappings.Add("Cost_Center", "Cost_Center");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("DebitCredit_ind", "DebitCredit_ind");
                bulkcopy.ColumnMappings.Add("GL_Account", "GL_Account");
                bulkcopy.ColumnMappings.Add("Offsetting_Account", "Offsetting_Account");
                bulkcopy.ColumnMappings.Add("Personnel_Number", "Personnel_Number");
                bulkcopy.ColumnMappings.Add("Account_ID", "Account_ID");
                bulkcopy.ColumnMappings.Add("House_bank", "House_bank");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created"); 
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import Finance CSV to database");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        public static void DeleteFile(string filePath, ErrorLog oErrorLog)
        {
            try
            {
                // Check if file exists with its full path    
                if (File.Exists(filePath))
                {
                    // If file found, delete it    
                    File.Delete(filePath);
                    Console.WriteLine("File deleted.");
                }
                else Console.WriteLine("File not found");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }

        }

        public static bool GetDTFromAccountsReceivableFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "`" });
                    csvReader.HasFieldsEnclosedInQuotes = false;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        if (datacolumn.ColumnName.Contains("Cleared/Open Items Symbol"))
                        {
                            datacolumn.ColumnName = "Symbol";
                        }
                        else
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        }
                        datacolumn.AllowDBNull = true;

                        if (datacolumn.ColumnName.ToString().ToLower() == "net_due_date" || datacolumn.ColumnName.Contains("Document_Date") || datacolumn.ColumnName.Contains("Clearing_Date")
                            || datacolumn.ColumnName.Contains("Posting_Date") || datacolumn.ColumnName.Contains("Value_Date") || datacolumn.ColumnName.Contains("Payment_Date"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 2 || i == 6 || i == 15 || i == 16 || i == 22 || i == 29)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteAccountsReceivable("", "", "", oErrorLog);
                InsertAccountsReceivable(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertAccountsReceivable(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["AccReceivabletableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Symbol", "Symbol");
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Net_Due_Date", "Net_Due_Date");
                bulkcopy.ColumnMappings.Add("Assignment", "Assignment");
                bulkcopy.ColumnMappings.Add("Document_Number", "Document_Number");
                bulkcopy.ColumnMappings.Add("Document_type", "Document_type");
                bulkcopy.ColumnMappings.Add("Document_Date", "Document_Date");
                bulkcopy.ColumnMappings.Add("Special_GL_Ind", "Special_GL_Ind");
                bulkcopy.ColumnMappings.Add("Net_Due_Date_Symbol", "Net_Due_Date_Symbol");
                bulkcopy.ColumnMappings.Add("Amount_in_Local_Currency", "Amount_in_Local_Currency");
                bulkcopy.ColumnMappings.Add("Local_Currency", "Local_Currency");
                bulkcopy.ColumnMappings.Add("Clearing_Document", "Clearing_Document");
                bulkcopy.ColumnMappings.Add("Text", "Text");
                bulkcopy.ColumnMappings.Add("Document_Header_Text", "Document_Header_Text");
                bulkcopy.ColumnMappings.Add("Reference", "Reference");
                bulkcopy.ColumnMappings.Add("Clearing_Date", "Clearing_Date");
                bulkcopy.ColumnMappings.Add("Posting_Date", "Posting_Date");
                bulkcopy.ColumnMappings.Add("Purchasing_Document", "Purchasing_Document");
                bulkcopy.ColumnMappings.Add("Cost_Center", "Cost_Center");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("DebitCredit_ind", "DebitCredit_ind");
                bulkcopy.ColumnMappings.Add("Invoice_Reference", "Invoice_Reference");
                bulkcopy.ColumnMappings.Add("Value_Date", "Value_Date");
                bulkcopy.ColumnMappings.Add("Billing_Document", "Billing_Document");
                bulkcopy.ColumnMappings.Add("Sales_Document", "Sales_Document");
                bulkcopy.ColumnMappings.Add("Discount_Amount", "Discount_Amount");
                bulkcopy.ColumnMappings.Add("Trading_Partner_No", "Trading_Partner_No");
                bulkcopy.ColumnMappings.Add("Contract_Number", "Contract_Number");
                bulkcopy.ColumnMappings.Add("Contract_Type", "Contract_Type");
                bulkcopy.ColumnMappings.Add("Payment_Date", "Payment_Date");
                bulkcopy.ColumnMappings.Add("Disputed_Item", "Disputed_Item");
                bulkcopy.ColumnMappings.Add("Payment_Method", "Payment_Method");
                bulkcopy.ColumnMappings.Add("Payment_terms", "Payment_terms");
                bulkcopy.ColumnMappings.Add("Reason_Code", "Reason_Code");
                bulkcopy.ColumnMappings.Add("GL_Account", "GL_Account");
                bulkcopy.ColumnMappings.Add("Payment_Sent", "Payment_Sent");
                bulkcopy.ColumnMappings.Add("Pmnt_currency", "Pmnt_currency");
                bulkcopy.ColumnMappings.Add("Amt_in_Payment_Currency", "Amt_in_Payment_Currency");
                bulkcopy.ColumnMappings.Add("Payment_Order", "Payment_Order");
                bulkcopy.ColumnMappings.Add("Reverse_Clearing", "Reverse_Clearing");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.ColumnMappings.Add("Account", "Account");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import Accounts Receivable CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static void DeleteAccountsReceivable(string Year, string Period, string Company_Code, ErrorLog oErrorLog)
        {
            string tableName = ConfigurationManager.AppSettings["AccReceivabletableName"];
            string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();
            SqlConnection connection = new SqlConnection(ssqlconnectionstring);
            oErrorLog.WriteErrorLog("Connected to Database successfully.");

            string sqlStatement = " DELETE FROM [Accounts_Receivable]";
            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                //cmd.Parameters.AddWithValue("@Period", Period);
                //cmd.Parameters.AddWithValue("@Year", Year);
                //cmd.Parameters.AddWithValue("@Company_Code", Company_Code);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                oErrorLog.WriteErrorLog("Deleted the record from database table successfully [Period] = " + Period + " and [Year] = " + Year + " and [Company_Code] =" + Company_Code);
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        public static bool GetDTFromAccountsPayableFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "`" });
                    csvReader.HasFieldsEnclosedInQuotes = false;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        if (datacolumn.ColumnName.Contains("Cleared/Open Items Symbol"))
                        {
                            datacolumn.ColumnName = "Symbol";
                        }
                        else
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        }
                        datacolumn.AllowDBNull = true;

                        if (datacolumn.ColumnName.Contains("Document_Date") || datacolumn.ColumnName.Contains("Clearing_Date") || datacolumn.ColumnName.Contains("Posting_Date")
                            || datacolumn.ColumnName.ToString().ToLower() == "net_due_date" || datacolumn.ColumnName.Contains("Payment_Date"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();

                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 4 || i == 14 || i == 16 || i == 17 || i == 23)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteAccountsPayable("", "", "", oErrorLog);
                InsertAccountsPayable(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertAccountsPayable(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["AccPayabletableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Symbol", "Symbol");
                bulkcopy.ColumnMappings.Add("Assignment", "Assignment");
                bulkcopy.ColumnMappings.Add("Document_Number", "Document_Number");
                bulkcopy.ColumnMappings.Add("Document_type", "Document_type");
                bulkcopy.ColumnMappings.Add("Document_Date", "Document_Date");
                bulkcopy.ColumnMappings.Add("Special_GL_Ind", "Special_GL_Ind");
                bulkcopy.ColumnMappings.Add("Net_Due_Date_Symbol", "Net_Due_Date_Symbol");
                bulkcopy.ColumnMappings.Add("Amount_in_Local_Currency", "Amount_in_Local_Currency");
                bulkcopy.ColumnMappings.Add("Local_Currency", "Local_Currency");
                bulkcopy.ColumnMappings.Add("Clearing_Document", "Clearing_Document");
                bulkcopy.ColumnMappings.Add("Text", "Text");
                bulkcopy.ColumnMappings.Add("Check_Number_From", "Check_Number_From");
                bulkcopy.ColumnMappings.Add("Document_Header_Text", "Document_Header_Text");
                bulkcopy.ColumnMappings.Add("Reference", "Reference");
                bulkcopy.ColumnMappings.Add("Clearing_Date", "Clearing_Date");
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Posting_Date", "Posting_Date");
                bulkcopy.ColumnMappings.Add("Net_Due_Date", "Net_Due_Date");
                bulkcopy.ColumnMappings.Add("Purchasing_Document", "Purchasing_Document");
                bulkcopy.ColumnMappings.Add("Cost_Center", "Cost_Center");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("DebitCredit_ind", "DebitCredit_ind");
                bulkcopy.ColumnMappings.Add("Contract_Number", "Contract_Number");
                bulkcopy.ColumnMappings.Add("Contract_Type", "Contract_Type");
                bulkcopy.ColumnMappings.Add("Payment_Date", "Payment_Date");
                bulkcopy.ColumnMappings.Add("Payment_Method", "Payment_Method");
                bulkcopy.ColumnMappings.Add("Reason_Code", "Reason_Code");
                bulkcopy.ColumnMappings.Add("GL_Account", "GL_Account");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.ColumnMappings.Add("Reverse_Clearing", "Reverse_Clearing");
                bulkcopy.ColumnMappings.Add("Account", "Account");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import Accounts Payable CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static void DeleteAccountsPayable(string Year, string Period, string Company_Code, ErrorLog oErrorLog)
        {
            string tableName = ConfigurationManager.AppSettings["AccPayabletableName"];
            string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();
            SqlConnection connection = new SqlConnection(ssqlconnectionstring);
            oErrorLog.WriteErrorLog("Connected to Database successfully.");

            string sqlStatement = " DELETE FROM [Accounts_Payable]  ";
            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                oErrorLog.WriteErrorLog("Deleted the record from database table successfully [Period] = " + Period + " and [Year] = " + Year + " and [Company_Code] =" + Company_Code);
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        public static void DeleteDatabaseTable(string tablename, ErrorLog oErrorLog)
        {
            string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();
            SqlConnection connection = new SqlConnection(ssqlconnectionstring);
            oErrorLog.WriteErrorLog("Connected to Database successfully.");

            string sqlStatement = " DELETE FROM ["+tablename +"]  ";
            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                oErrorLog.WriteErrorLog("Deleted the record from database table successfully.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        public static bool GetDTFromPLCSVFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            string strCompanyCode = string.Empty;
            string strYear = string.Empty;
            string strPeriod = string.Empty;
            string ReportingDate = string.Empty;
            DateTime dtnow = DateTime.Now;

            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { @"\n" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    //read column names
                    string[] colFields = new string[] { "NoUsedCol1", "F_C", "Company_Code", "Business_Area", "Text", "Amount", "NoUsedCol2", "NoUsedCol3", "NoUsedCol4", "Summary_Level", "Date_Created", "Period", "Year" };

                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "");
                        datacolumn.AllowDBNull = true;
                        csvData.Columns.Add(datacolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        bool TextNotInsert = false;

                        if (fieldData[0].ToString().StartsWith("|F|"))
                        {
                            string[] strarrData = fieldData[0].Split(new char[] { '|' });
                            for (int colCount = 1; colCount < strarrData.Length; colCount++)
                            {
                                if (colCount == 5 && String.IsNullOrEmpty(strCompanyCode))
                                {
                                    ReportingDate = strarrData[colCount].Trim().Replace("(", "").Replace(")", "");
                                    string[] dateRange = ReportingDate.Split(new char[] { '-', '.' });
                                    strPeriod = dateRange[0];
                                    strYear = dateRange[1];
                                }
                            }
                        }
                        else if (fieldData[0].ToString().StartsWith("| |"))
                        {
                            string[] strarrData = fieldData[0].Split(new char[] { '|' });
                            myDataRow = csvData.NewRow();

                            for (int colCount = 1; colCount < strarrData.Length; colCount++)
                            {
                                if (colCount == 2 && String.IsNullOrEmpty(strCompanyCode))
                                    strCompanyCode = strarrData[colCount];
                                if (colCount == 4 && String.IsNullOrWhiteSpace(strarrData[colCount]))
                                    TextNotInsert = true;
                                if (colCount >= 5 && strarrData[colCount] != "" && strarrData[colCount].Contains("-"))
                                    strarrData[colCount] = "-" + strarrData[colCount].Trim().Replace("-", "");

                                if (colCount == 2 && string.IsNullOrEmpty(strarrData[colCount].Trim()))
                                    strarrData[colCount] = strCompanyCode;

                                if (strarrData[colCount] != "")
                                    myDataRow[colCount] = strarrData[colCount].Trim().Replace("*", "");
                                else
                                    myDataRow[colCount] = null;

                                myDataRow["Date_Created"] = dtnow;
                                myDataRow["Period"] = strPeriod;
                                myDataRow["Year"] = strYear;
                            }
                            if (!TextNotInsert)
                            {
                                csvData.Rows.Add(myDataRow);
                                TextNotInsert = false;
                            }
                        }
                    }
                }

                DeleteCompanyCodeProfitLost(strYear, strPeriod, strCompanyCode, oErrorLog);
                InsertProfitLost(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                return false;
            }
        }

        public static void DeleteCompanyCodeProfitLost(string Year, string Period, string Company_Code, ErrorLog oErrorLog)
        {
            string tableName = ConfigurationManager.AppSettings["ProfitLosttableName"];
            string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();
            SqlConnection connection = new SqlConnection(ssqlconnectionstring);
            oErrorLog.WriteErrorLog("Connected to Database successfully.");

            string sqlStatement = " DELETE FROM [FinancialStatement] WHERE [Period] = @Period and [Year] = @Year and [Company_Code] = @Company_Code ";
            try
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sqlStatement, connection);
                cmd.Parameters.AddWithValue("@Period", Period);
                cmd.Parameters.AddWithValue("@Year", Year);
                cmd.Parameters.AddWithValue("@Company_Code", Company_Code);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                oErrorLog.WriteErrorLog("Deleted the record from database table successfully [Period] = " + Period + " and [Year] = " + Year + " and [Company_Code] =" + Company_Code);
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        public static void InsertProfitLost(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["ProfitLosttableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Year", "Year");
                bulkcopy.ColumnMappings.Add("Period", "Period");
                bulkcopy.ColumnMappings.Add("F_C", "F_C");
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Business_Area", "Business_Area");
                bulkcopy.ColumnMappings.Add("Text", "Text");
                bulkcopy.ColumnMappings.Add("Amount", "Amount");
                bulkcopy.ColumnMappings.Add("Summary_Level", "Summary_Level");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import PnL CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromROMasterFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = false;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Row_Valid_From") || datacolumn.ColumnName.Contains("Row_Valid_To") || datacolumn.ColumnName.Contains("Cash_Flow_From"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 9 || i == 17 || i == 18)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["ROMastertableName"], oErrorLog);
                InsertROMaster(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertROMaster(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["ROMastertableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Company_Name", "Company_Name");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Business_Entity_Name", "Business_Entity_Name");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Building_Name", "Building_Name");
                bulkcopy.ColumnMappings.Add("Rental_Object", "Rental_Object");
                bulkcopy.ColumnMappings.Add("Rental_Object_Name", "Rental_Object_Name");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Usage_type_of_rental_unit", "Usage_type_of_rental_unit");
                bulkcopy.ColumnMappings.Add("Cash_Flow_From", "Cash_Flow_From");
                bulkcopy.ColumnMappings.Add("Neighborhood", "Neighborhood");
                bulkcopy.ColumnMappings.Add("Floor_shrt_nme", "Floor_shrt_nme");
                bulkcopy.ColumnMappings.Add("Floor_long_name", "Floor_long_name");
                bulkcopy.ColumnMappings.Add("City", "City");
                bulkcopy.ColumnMappings.Add("Country_Key", "Country_Key");
                bulkcopy.ColumnMappings.Add("RU_No_Old", "RU_No_Old");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Row_Valid_From", "Row_Valid_From");
                bulkcopy.ColumnMappings.Add("Row_Valid_To", "Row_Valid_To");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromExcelFileSource2(string csv_file_path, ErrorLog oErrorLog)
        {
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;

            try
            {
                //Create COM Objects.
                ExcelApp.Application excelApp = new ExcelApp.Application();
                DataRow myNewRow;
                DataTable MyDataTable = new DataTable();
                DateTime conv;
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return true;
                }

                //Notice: Change this path to your real excel file path
                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(csv_file_path);
                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                ExcelApp.Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                //read column names
                string[] colFields = new string[]  { "Cleared/Open Items Symbol","Assignment","Document Number","Business Area","Document type","Document Date",
                    "Posting Key","Amount in Local Currency","Local Currency","Material","Profit Center","Segment","Text","Offsetting Account","Quantity",
                        "Plant","Posting Date","Company Code","Order","Clearing Date","Fiscal Year","Cost Center","G/L Account"};

                foreach (string column in colFields)
                {
                    DataColumn datacolumn = new DataColumn(column);
                    if (datacolumn.ColumnName.Contains("Cleared/Open Items Symbol"))
                    {
                        datacolumn.ColumnName = "Symbol";
                    }
                    else
                    {
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                    }
                    datacolumn.AllowDBNull = true;

                    if (datacolumn.ColumnName.Contains("Document_Date") || datacolumn.ColumnName.Contains("Clearing_Date") || datacolumn.ColumnName.Contains("Posting_Date")
                            || datacolumn.ColumnName.ToString().ToLower() == "net_due_date" || datacolumn.ColumnName.Contains("Payment_Date"))
                        datacolumn.DataType = System.Type.GetType("System.DateTime");

                    MyDataTable.Columns.Add(datacolumn);
                }

                DataColumn dcCreatedDate = new DataColumn("Date_Created");
                dcCreatedDate.AllowDBNull = true;
                MyDataTable.Columns.Add(dcCreatedDate);

                //first row using for heading, start second row for data
                for (int i = 2; i <= rows; i++)
                {
                    myNewRow = MyDataTable.NewRow();
                    myNewRow["Symbol"] = "1"; excelRange.Cells[i, 1].Value2.ToString();
                    myNewRow["Assignment"] = excelRange.Cells[i, 2].Value2.ToString() != "" ? excelRange.Cells[i, 2].Value2.ToString() : null;
                    myNewRow["Document_Number"] = excelRange.Cells[i, 3].Value2.ToString() != "" ? excelRange.Cells[i, 3].Value2.ToString() : null;
                    myNewRow["Business_Area"] = excelRange.Cells[i, 4].Value2.ToString() != "" ? excelRange.Cells[i, 4].Value2.ToString() : null;
                    myNewRow["Document_type"] = excelRange.Cells[i, 5].Value2.ToString() != "" ? excelRange.Cells[i, 5].Value2.ToString() : null;
                    if (excelRange.Cells[i, 6].Value2 != null)
                    {
                        conv = DateTime.FromOADate(double.Parse(excelRange.Cells[i, 6].Value2.ToString()));
                        myNewRow["Document_Date"] = conv;
                    }
                    myNewRow["Posting_Key"] = excelRange.Cells[i, 7].Value2.ToString() != "" ? excelRange.Cells[i, 7].Value2.ToString() : null;
                    myNewRow["Amount_in_Local_Currency"] = excelRange.Cells[i, 8].Value2.ToString() != "" ? excelRange.Cells[i, 8].Value2.ToString() : null;
                    myNewRow["Local_Currency"] = excelRange.Cells[i, 9].Value2.ToString() != "" ? excelRange.Cells[i, 9].Value2.ToString() : null;
                    myNewRow["Material"] = excelRange.Cells[i, 10].Value2.ToString() != "" ? excelRange.Cells[i, 10].Value2.ToString() : null;
                    myNewRow["Profit_Center"] = excelRange.Cells[i, 11].Value2.ToString() != "" ? excelRange.Cells[i, 11].Value2.ToString() : null;
                    myNewRow["Segment"] = excelRange.Cells[i, 12].Value2.ToString() != "" ? excelRange.Cells[i, 12].Value2.ToString() : null;
                    myNewRow["Text"] = excelRange.Cells[i, 13].Value2.ToString() != "" ? excelRange.Cells[i, 13].Value2.ToString() : null;
                    myNewRow["Offsetting_Account"] = excelRange.Cells[i, 14].Value2.ToString() != "" ? excelRange.Cells[i, 14].Value2.ToString() : null;
                    myNewRow["Quantity"] = excelRange.Cells[i, 15].Value2 != null ? excelRange.Cells[i, 15].Value2.ToString() : null;
                    myNewRow["Plant"] = excelRange.Cells[i, 16].Value2.ToString() != "" ? excelRange.Cells[i, 16].Value2.ToString() : null;
                    if (excelRange.Cells[i, 17].Value2 != null)
                    {
                        conv = DateTime.FromOADate(double.Parse(excelRange.Cells[i, 17].Value2.ToString()));
                        myNewRow["Posting_Date"] = conv;
                    }
                    myNewRow["Company_Code"] = excelRange.Cells[i, 18].Value2.ToString() != "" ? excelRange.Cells[i, 18].Value2.ToString() : null;
                    myNewRow["Order"] = excelRange.Cells[i, 19].Value2.ToString() != "" ? excelRange.Cells[i, 19].Value2.ToString() : null;
                    if (excelRange.Cells[i, 20].Value2 != null)
                    {
                        conv = DateTime.FromOADate(double.Parse(excelRange.Cells[i, 20].Value2.ToString()));
                        myNewRow["Clearing_Date"] = conv;
                    }
                    myNewRow["Fiscal_Year"] = excelRange.Cells[i, 21].Value2.ToString() != "" ? excelRange.Cells[i, 21].Value2.ToString() : null;
                    myNewRow["Cost_Center"] = excelRange.Cells[i, 22].Value2.ToString() != "" ? excelRange.Cells[i, 22].Value2.ToString() : null;
                    myNewRow["GL_Account"] = excelRange.Cells[i, 23].Value2.ToString() != "" ? excelRange.Cells[i, 23].Value2.ToString() : null;
                    myNewRow["Date_Created"] = dtnow;
                    MyDataTable.Rows.Add(myNewRow);
                    Console.WriteLine("Successfully Validated Row number:" + RowCount);
                    RowCount++;
                }

                InsertFinance_Source2(MyDataTable, oErrorLog);

                //after reading, relaase the excel project
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                Console.ReadLine();

                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertFinance_Source2(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["Finance_Source2"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Symbol", "Symbol");
                bulkcopy.ColumnMappings.Add("Assignment", "Assignment");
                bulkcopy.ColumnMappings.Add("Document_Number", "Document_Number");
                bulkcopy.ColumnMappings.Add("Business_Area", "Business_Area");
                bulkcopy.ColumnMappings.Add("Document_type", "Document_type");
                bulkcopy.ColumnMappings.Add("Document_Date", "Document_Date");
                bulkcopy.ColumnMappings.Add("Posting_Key", "Posting_Key");
                bulkcopy.ColumnMappings.Add("Amount_in_Local_Currency", "Amount_in_Local_Currency");
                bulkcopy.ColumnMappings.Add("Local_Currency", "Local_Currency");
                bulkcopy.ColumnMappings.Add("Material", "Material");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Segment", "Segment");
                bulkcopy.ColumnMappings.Add("Text", "Text");
                bulkcopy.ColumnMappings.Add("Offsetting_Account", "Offsetting_Account");
                bulkcopy.ColumnMappings.Add("Quantity", "Quantity");
                bulkcopy.ColumnMappings.Add("Plant", "Plant");
                bulkcopy.ColumnMappings.Add("Posting_Date", "Posting_Date");
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Order", "Order");
                bulkcopy.ColumnMappings.Add("Clearing_Date", "Clearing_Date");
                bulkcopy.ColumnMappings.Add("Fiscal_Year", "Fiscal_Year");
                bulkcopy.ColumnMappings.Add("Cost_Center", "Cost_Center");
                bulkcopy.ColumnMappings.Add("GL_Account", "GL_Account");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully imported Finance Source 2 Excel Document to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromRO_FFMFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Row_Valid_From") || datacolumn.ColumnName.Contains("Row_Valid_To") || datacolumn.ColumnName.Contains("Cash_Flow_From"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 9 || i == 17 || i == 18)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["FFMtableName"], oErrorLog);
                InsertRO_FFM(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertRO_FFM(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["FFMtableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Company_Name", "Company_Name");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Business_Entity_Name", "Business_Entity_Name");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Building_Name", "Building_Name");
                bulkcopy.ColumnMappings.Add("Rental_Object", "Rental_Object");
                bulkcopy.ColumnMappings.Add("Rental_Object_Name", "Rental_Object_Name");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Usage_type_of_rental_unit", "Usage_type_of_rental_unit");
                bulkcopy.ColumnMappings.Add("Cash_Flow_From", "Cash_Flow_From");
                bulkcopy.ColumnMappings.Add("Neighborhood", "Neighborhood");
                bulkcopy.ColumnMappings.Add("Floor_shrt_nme", "Floor_shrt_nme");
                bulkcopy.ColumnMappings.Add("Floor_long_name", "Floor_long_name");
                bulkcopy.ColumnMappings.Add("City", "City");
                bulkcopy.ColumnMappings.Add("Country_Key", "Country_Key");
                bulkcopy.ColumnMappings.Add("RU_No_Old", "RU_No_Old");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Row_Valid_From", "Row_Valid_From");
                bulkcopy.ColumnMappings.Add("Row_Valid_To", "Row_Valid_To");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromROMeasureFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        datacolumn.AllowDBNull = true;
                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["MeasurementsMastertableName"], oErrorLog);
                InsertROMeasure(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertROMeasure(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["MeasurementsMastertableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Pooled_SpaceRU", "Pooled_SpaceRU");
               // bulkcopy.ColumnMappings.Add("Rental_Space", "Rental_Space");
                bulkcopy.ColumnMappings.Add("Object_Type", "Object_Type");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Name", "Name");
               // bulkcopy.ColumnMappings.Add("Status", "Status");
                bulkcopy.ColumnMappings.Add("Med_Name_of_Meas_Type", "Med_Name_of_Meas_Type");
                bulkcopy.ColumnMappings.Add("Amount", "Amount");
                bulkcopy.ColumnMappings.Add("Capacity", "Capacity");
                bulkcopy.ColumnMappings.Add("Units_in", "Units_in");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromRO_Conditions_MasterFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    int Condition = 0; 
                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        if (datacolumn.ColumnName.Contains("Condition Type"))
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_") + "_" + Condition;
                            Condition++;
                        }
                        else
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        }
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Relnship_Valid_From") || datacolumn.ColumnName.Contains("Relnship_Valid_To"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 11 || i == 12)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["ConditionsMastertableName"], oErrorLog);
                InsertRO_Conditions_Master(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertRO_Conditions_Master(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["ConditionsMastertableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Pooled_SpaceRU", "Pooled_SpaceRU");
                bulkcopy.ColumnMappings.Add("Object_Type", "Object_Type");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Name", "Name");
                bulkcopy.ColumnMappings.Add("Condition_Type_0", "Condition_Type");
                bulkcopy.ColumnMappings.Add("Condition_Type_1", "Condition_Type_Text");
                bulkcopy.ColumnMappings.Add("Condition_Amount", "Condition_Amount");
                bulkcopy.ColumnMappings.Add("Currency", "Currency");
                bulkcopy.ColumnMappings.Add("Relnship_Valid_From", "Relnship_Valid_From");
                bulkcopy.ColumnMappings.Add("Relnship_Valid_To", "Relnship_Valid_To");
                bulkcopy.ColumnMappings.Add("Frequency", "Frequency");
                bulkcopy.ColumnMappings.Add("Frequency_Unit", "Frequency_Unit");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromCashflowFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0, Condition = 1, ObjectId = 1, RoleCategory = 1, AddressType = 1;

            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "`" });
                    csvReader.HasFieldsEnclosedInQuotes = false;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        if (datacolumn.ColumnName.Contains("Condition Type"))
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "") + "_" + Condition;
                            Condition++;
                        }
                        else if (datacolumn.ColumnName.Contains("Object ID"))
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "") + "_" + ObjectId;
                            ObjectId++;
                        }
                        else if (datacolumn.ColumnName.Contains("Role Category Categ"))
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "") + "_" + RoleCategory;
                            RoleCategory++;
                        }
                        else if (datacolumn.ColumnName.Contains("Address Type"))
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "") + "_" + AddressType;
                            AddressType++;
                        }
                        else
                        {
                            datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "").Replace(":", "");
                        }
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Contract_start_date") || datacolumn.ColumnName.Contains("Calculation_From") || datacolumn.ColumnName.Contains("Calculation_to")
                            || datacolumn.ColumnName.Contains("Active_from") || datacolumn.ColumnName.Contains("Due_date") || datacolumn.ColumnName.Contains("Cash_Flow_From")
                            || datacolumn.ColumnName.Contains("Calculation_Date") || datacolumn.ColumnName.Contains("Row_Valid_From") || datacolumn.ColumnName.Contains("Row_Valid_To")
                            || datacolumn.ColumnName.Contains("Object_Valid_To") || datacolumn.ColumnName.Contains("First_Contract_End"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 11 || i == 12 || i == 13 || i == 16 || i == 17 || i == 20 || i == 25 || i == 58 || i == 59 || i == 68 || i == 69)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["CashflowtableName"], oErrorLog);
                InsertCashflow(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertCashflow(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["CashflowtableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Object_Company_Code", "Object_Company_Code");
                bulkcopy.ColumnMappings.Add("Real_Estate_Key", "Real_Estate_Key");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Pooled_SpaceRU", "Pooled_SpaceRU");
                bulkcopy.ColumnMappings.Add("Status", "Status");
                bulkcopy.ColumnMappings.Add("Contrtype_text", "Contrtype_text");
                bulkcopy.ColumnMappings.Add("Object_ID_1", "Object_ID");
                bulkcopy.ColumnMappings.Add("Contract", "Contract");
                bulkcopy.ColumnMappings.Add("1st_Main_Contractual_Partner_Name", "First_Main_Contractual_Partner_Name");
                bulkcopy.ColumnMappings.Add("Object_ID_2", "Object_ID2");
                bulkcopy.ColumnMappings.Add("Contract_start_date", "Contract_start_date");
                bulkcopy.ColumnMappings.Add("Calculation_From", "Calculation_From");
                bulkcopy.ColumnMappings.Add("Calculation_to", "Calculation_to");
                bulkcopy.ColumnMappings.Add("Condition_Type_1", "Condition_Type");
                bulkcopy.ColumnMappings.Add("Condition_Type_2", "Condition_Type_Text");
                bulkcopy.ColumnMappings.Add("Active_from", "Active_from");
                bulkcopy.ColumnMappings.Add("Due_date", "Due_date");
                bulkcopy.ColumnMappings.Add("Net_in_Local_Crcy", "Net_in_Local_Crcy");
                bulkcopy.ColumnMappings.Add("Contract_Currency", "Contract_Currency");
                bulkcopy.ColumnMappings.Add("Cash_Flow_From", "Cash_Flow_From");
                bulkcopy.ColumnMappings.Add("BE_for_Contract", "BE_for_Contract");
                bulkcopy.ColumnMappings.Add("Company_Name", "Company_Name");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Tenancy_law_description", "Tenancy_law_description");
                bulkcopy.ColumnMappings.Add("Calculation_Date", "Calculation_Date");
                bulkcopy.ColumnMappings.Add("WBS_Element", "WBS_Element");
                bulkcopy.ColumnMappings.Add("Valuation_Relevance", "Valuation_Relevance");
                bulkcopy.ColumnMappings.Add("Name_of_Cost_Center", "Name_of_Cost_Center");
                bulkcopy.ColumnMappings.Add("Order_Name", "Order_Name");
                bulkcopy.ColumnMappings.Add("Description_of_WBS_element", "Description_of_WBS_element");
                bulkcopy.ColumnMappings.Add("Local_Currency", "Local_Currency");
                bulkcopy.ColumnMappings.Add("Net_CndCrcy", "Net_CndCrcy");
                bulkcopy.ColumnMappings.Add("Gross_CndCrcy", "Gross_CndCrcy");
                bulkcopy.ColumnMappings.Add("Sales_Tax_CndCrcy", "Sales_Tax_CndCrcy");
                bulkcopy.ColumnMappings.Add("Condition_currency", "Condition_currency");
                bulkcopy.ColumnMappings.Add("Gross", "Gross");
                bulkcopy.ColumnMappings.Add("Posting_Term", "Posting_Term");
                bulkcopy.ColumnMappings.Add("Frequency_Term", "Frequency_Term");
                bulkcopy.ColumnMappings.Add("Org_Term", "Org_Term");
                bulkcopy.ColumnMappings.Add("Sales_Term", "Sales_Term");
                bulkcopy.ColumnMappings.Add("Peak_Sales_Term", "Peak_Sales_Term");
                bulkcopy.ColumnMappings.Add("Withholding_Tax_Term", "Withholding_Tax_Term");
                bulkcopy.ColumnMappings.Add("Crcy_Transl_Rule", "Crcy_Transl_Rule");
                bulkcopy.ColumnMappings.Add("Gross_in_Local_Crcy", "Gross_in_Local_Crcy");
                bulkcopy.ColumnMappings.Add("LC_Tax_Amount", "LC_Tax_Amount");
                bulkcopy.ColumnMappings.Add("Exchange_Rate", "Exchange_Rate");
                bulkcopy.ColumnMappings.Add("Post_in_Cond_Crcy", "Post_in_Cond_Crcy");
                bulkcopy.ColumnMappings.Add("Planned_TrslDate", "Planned_TrslDate");
                bulkcopy.ColumnMappings.Add("Actual_TranslDate", "Actual_TranslDate");
                bulkcopy.ColumnMappings.Add("Doc_Reference_Key", "Doc_Reference_Key");
                bulkcopy.ColumnMappings.Add("Document", "Document");
                bulkcopy.ColumnMappings.Add("Reference_Flow_Type", "Reference_Flow_Type");
                bulkcopy.ColumnMappings.Add("Origin", "Origin");
                bulkcopy.ColumnMappings.Add("Role_Category_Categ_1", "Role_Category_Categ");
                bulkcopy.ColumnMappings.Add("Address_Type_1", "Address_Type");
                bulkcopy.ColumnMappings.Add("Role_Category_Categ_2", "Role_Category_Categ2");
                bulkcopy.ColumnMappings.Add("Address_Type_2", "Address_Type2");
                bulkcopy.ColumnMappings.Add("Row_Valid_From", "Row_Valid_From");
                bulkcopy.ColumnMappings.Add("Row_Valid_To", "Row_Valid_To");
                bulkcopy.ColumnMappings.Add("Acct_DetermValue", "Acct_DetermValue");
                bulkcopy.ColumnMappings.Add("PBCISR_number", "PBCISR_number");
                bulkcopy.ColumnMappings.Add("ISRQR_Reference_Number", "ISRQR_Reference_Number");
                bulkcopy.ColumnMappings.Add("Text_of_Vendor_Invoice", "Text_of_Vendor_Invoice");
                bulkcopy.ColumnMappings.Add("Delete_FI_Text", "Delete_FI_Text");
                bulkcopy.ColumnMappings.Add("Position", "Position");
                bulkcopy.ColumnMappings.Add("Cash_Flow_Partner", "Cash_Flow_Partner");
                bulkcopy.ColumnMappings.Add("Descr_Cash_Flow_BP", "Cash_Flow_Partner_Name");
                bulkcopy.ColumnMappings.Add("Object_Valid_To", "Object_Valid_To");
                bulkcopy.ColumnMappings.Add("1st_Contract_End", "First_Contract_End");
                bulkcopy.ColumnMappings.Add("Contract_Conclusion", "Contract_Conclusion");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import Cashflow CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromProfitabiltyFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Row_Valid_From") || datacolumn.ColumnName.Contains("Row_Valid_To") || datacolumn.ColumnName.Contains("Cash_Flow_From"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 9 || i == 17 || i == 18)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["ROMastertableName"], oErrorLog);
                InsertROMaster(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertProfitabilty(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["ROMastertableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Company_Name", "Company_Name");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Business_Entity_Name", "Business_Entity_Name");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Building_Name", "Building_Name");
                bulkcopy.ColumnMappings.Add("Rental_Object", "Rental_Object");
                bulkcopy.ColumnMappings.Add("Rental_Object_Name", "Rental_Object_Name");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Usage_type_of_rental_unit", "Usage_type_of_rental_unit");
                bulkcopy.ColumnMappings.Add("Cash_Flow_From", "Cash_Flow_From");
                bulkcopy.ColumnMappings.Add("Neighborhood", "Neighborhood");
                bulkcopy.ColumnMappings.Add("Floor_shrt_nme", "Floor_shrt_nme");
                bulkcopy.ColumnMappings.Add("Floor_long_name", "Floor_long_name");
                bulkcopy.ColumnMappings.Add("City", "City");
                bulkcopy.ColumnMappings.Add("Country_Key", "Country_Key");
                bulkcopy.ColumnMappings.Add("RU_No_Old", "RU_No_Old");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Row_Valid_From", "Row_Valid_From");
                bulkcopy.ColumnMappings.Add("Row_Valid_To", "Row_Valid_To");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }

        public static bool GetDTFromRO_OccupancyFile(string csv_file_path, ErrorLog oErrorLog)
        {
            DataTable csvData = new DataTable();
            DataRow myDataRow;
            DateTime dtnow = DateTime.Now;
            int RowCount = 0;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "|" });
                    csvReader.HasFieldsEnclosedInQuotes = true;

                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datacolumn = new DataColumn(column);
                        datacolumn.ColumnName = datacolumn.ColumnName.Replace(" ", "_").Replace("/", "").Replace(".", "").Replace(",,", "");
                        datacolumn.AllowDBNull = true;
                        if (datacolumn.ColumnName.Contains("Row_Valid_From") || datacolumn.ColumnName.Contains("Row_Valid_To") || datacolumn.ColumnName.Contains("Cash_Flow_From"))
                            datacolumn.DataType = System.Type.GetType("System.DateTime");

                        csvData.Columns.Add(datacolumn);
                    }

                    DataColumn dcCreatedDate = new DataColumn("Date_Created");
                    dcCreatedDate.AllowDBNull = true;
                    csvData.Columns.Add(dcCreatedDate);
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        myDataRow = csvData.NewRow();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] != "" && fieldData[i] != "00.00.0000")
                            {
                                myDataRow[i] = fieldData[i];
                            }
                            else if (fieldData[i] == "" || fieldData[i] == "00.00.0000")
                            {
                                if (i == 9 || i == 17 || i == 18)
                                    myDataRow[i] = DBNull.Value;
                                else
                                    myDataRow[i] = null;
                            }
                        }
                        myDataRow["Date_Created"] = dtnow;
                        csvData.Rows.Add(myDataRow);
                        RowCount++;
                    }
                }
                DeleteDatabaseTable(ConfigurationManager.AppSettings["ROMastertableName"], oErrorLog);
                InsertROMaster(csvData, oErrorLog);
                return true;
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
                oErrorLog.WriteErrorLog("Something went wrong on line number: " + RowCount + "in CSV file");
                return false;
            }
        }

        public static void InsertRO_Occupancy(DataTable dt, ErrorLog oErrorLog)
        {
            try
            {
                string tableName = ConfigurationManager.AppSettings["ROMastertableName"];
                string ssqlconnectionstring = ConfigurationManager.ConnectionStrings["DB_ConnectionString"].ToString();

                oErrorLog.WriteErrorLog("Connected to Database successfully.");
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = tableName;
                SqlConnection con = new SqlConnection(ssqlconnectionstring);
                con.Open();
                bulkcopy.ColumnMappings.Add("Company_Code", "Company_Code");
                bulkcopy.ColumnMappings.Add("Company_Name", "Company_Name");
                bulkcopy.ColumnMappings.Add("Business_Entity", "Business_Entity");
                bulkcopy.ColumnMappings.Add("Business_Entity_Name", "Business_Entity_Name");
                bulkcopy.ColumnMappings.Add("Building", "Building");
                bulkcopy.ColumnMappings.Add("Building_Name", "Building_Name");
                bulkcopy.ColumnMappings.Add("Rental_Object", "Rental_Object");
                bulkcopy.ColumnMappings.Add("Rental_Object_Name", "Rental_Object_Name");
                bulkcopy.ColumnMappings.Add("Object_ID", "Object_ID");
                bulkcopy.ColumnMappings.Add("Usage_type_of_rental_unit", "Usage_type_of_rental_unit");
                bulkcopy.ColumnMappings.Add("Cash_Flow_From", "Cash_Flow_From");
                bulkcopy.ColumnMappings.Add("Neighborhood", "Neighborhood");
                bulkcopy.ColumnMappings.Add("Floor_shrt_nme", "Floor_shrt_nme");
                bulkcopy.ColumnMappings.Add("Floor_long_name", "Floor_long_name");
                bulkcopy.ColumnMappings.Add("City", "City");
                bulkcopy.ColumnMappings.Add("Country_Key", "Country_Key");
                bulkcopy.ColumnMappings.Add("RU_No_Old", "RU_No_Old");
                bulkcopy.ColumnMappings.Add("Profit_Center", "Profit_Center");
                bulkcopy.ColumnMappings.Add("Row_Valid_From", "Row_Valid_From");
                bulkcopy.ColumnMappings.Add("Row_Valid_To", "Row_Valid_To");
                bulkcopy.ColumnMappings.Add("Date_Created", "Date_Created");
                bulkcopy.WriteToServer(dt);
                con.Close();
                oErrorLog.WriteErrorLog("Successfully import RO Master CSV to database table.");
            }
            catch (Exception ex)
            {
                oErrorLog.WriteErrorLog(ex.Message);
            }
        }
    }
}
