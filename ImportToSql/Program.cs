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
namespace ImportToSql
{
    class Program
    {
        static void Main(string[] args)
        {
            ErrorLog oErrorLog = new ErrorLog();
            string Financefilepath = ConfigurationManager.AppSettings["FinancefilePath"];
            string PLfilepath = ConfigurationManager.AppSettings["ProfitLossfilePath"];
            string AllowToRunFinance = ConfigurationManager.AppSettings["AllowToRunFinance"];
            string AllowToRunPnL = ConfigurationManager.AppSettings["AllowToRunProfitLost"];

            oErrorLog.WriteErrorLog(" ");
            oErrorLog.WriteErrorLog("----------------------------------------");
            oErrorLog.WriteErrorLog("Import task starting...");
            oErrorLog.WriteErrorLog("Open CSV file to read data");
            Console.WriteLine("Data Extraction has been started.");

            if (AllowToRunFinance == "true")
            {
                var dt = CsvReader.GetDataTabletFromCSVFile(Financefilepath, oErrorLog);
            }
            Console.WriteLine("Finance CSV file has been completed.");

            if (AllowToRunPnL == "true")
            {
                var dtProfitLost = CsvReader.GetDTFromPLCSVFile(PLfilepath, oErrorLog);
            }
            Console.WriteLine("Profit And Lost CSV file has been completed.");
            Console.Read();
        }
    }
}

