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
            string Aurafilepath = ConfigurationManager.AppSettings["AurafilePath"];
            string Auragroupfilepath = ConfigurationManager.AppSettings["AuragroupfilePath"];
            string ElegenciPalfilepath = ConfigurationManager.AppSettings["ElegenciPalfilePath"];
            string Yemekfilepath = ConfigurationManager.AppSettings["YemekfilePath"];
            string RO_MasterfilePath = ConfigurationManager.AppSettings["RO_MasterfilePath"];
            string FFMfilePath = ConfigurationManager.AppSettings["FFMfilePath"];
            string MMfilePath = ConfigurationManager.AppSettings["Measurements_MasterfilePath"];
            string CMfilePath = ConfigurationManager.AppSettings["Conditions_MasterfilePath"];
            string OccupancyfilePath = ConfigurationManager.AppSettings["OccupancyfilePath"];
            string CashflowfilePath = ConfigurationManager.AppSettings["CashflowfilePath"];
            string ProfitabiltyfilePath = ConfigurationManager.AppSettings["ProfitabiltyfilePath"];
            string AccReceivablefilepath = ConfigurationManager.AppSettings["AccReceivablefilePath"];
            string AccPayablefilepath = ConfigurationManager.AppSettings["AccPayablefilePath"];
            string PLfilepath = ConfigurationManager.AppSettings["ProfitLossfilePath"];
            string FS2filepath = ConfigurationManager.AppSettings["FS2filePath"];

            string AllowToRunAura = ConfigurationManager.AppSettings["AllowToRunAura"];
            string AllowToRunPnL = ConfigurationManager.AppSettings["AllowToRunProfitLost"];
            string AllowToRunAG = ConfigurationManager.AppSettings["AllowToRunAuraGroup"];
            string AllowToRunEP = ConfigurationManager.AppSettings["AllowToRunElegenciPal"];
            string AllowToRunYemek = ConfigurationManager.AppSettings["AllowToRunYemek"];
            string AllowToRunRO_Master = ConfigurationManager.AppSettings["AllowToRunRO_Master"];
            string AllowToFFM = ConfigurationManager.AppSettings["AllowToRunFFM"];
            string AllowToRunMM = ConfigurationManager.AppSettings["AllowToRunMeasurements_Master"];
            string AllowToRunCM = ConfigurationManager.AppSettings["AllowToRunConditions_Master"];
            string AllowToRunOccupancy = ConfigurationManager.AppSettings["AllowToRunOccupancy"];
            string AllowToRunCashflow = ConfigurationManager.AppSettings["AllowToRunCashflow"];
            string AllowToRunProfitabilty = ConfigurationManager.AppSettings["AllowToRunProfitabilty"];
            string AllowToRunAccReceivable = ConfigurationManager.AppSettings["AllowToRunAccReceivable"];
            string AllowToRunAccPayable = ConfigurationManager.AppSettings["AllowToRunAccPayable"];
            string AllowToFinanceSource2 = ConfigurationManager.AppSettings["AllowToRunFS2"];

            oErrorLog.WriteErrorLog(" ");
            oErrorLog.WriteErrorLog("----------------------------------------");
            oErrorLog.WriteErrorLog("Import task starting...");
            oErrorLog.WriteErrorLog("Open CSV file to read data");
            Console.WriteLine("Data Extraction has been started.");

            if (AllowToRunAura == "true")
            {
                Console.WriteLine("Assets Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(Aurafilepath, oErrorLog);
                Console.WriteLine("Assets CSV file has been completed.");
            }
            if (AllowToRunAG == "true")
            {
                Console.WriteLine("Aura group Data Extraction has been started.");
                var dt2 = CsvReader.GetDataTabletFromCSVFile(Auragroupfilepath, oErrorLog);
                Console.WriteLine("Aura group CSV file has been completed.");
            }

            if (AllowToRunAccReceivable == "true")
            {
                Console.WriteLine("Accounts Receivable  Data Extraction has been started.");
                var dt1 = CsvReader.GetDTFromAccountsReceivableFile(AccReceivablefilepath, oErrorLog);
                Console.WriteLine("Accounts Receivable CSV file has been completed.");
            }
            if (AllowToRunAccPayable == "true")
            {
                Console.WriteLine("Accounts Payable Data Extraction has been started.");
                var dt2 = CsvReader.GetDTFromAccountsPayableFile(AccPayablefilepath, oErrorLog);
                Console.WriteLine("Accounts Payable CSV file has been completed.");
            }

            if (AllowToRunEP == "true")
            {
                Console.WriteLine("Elegenci Pal Data Extraction has been started.");
                var dt3 = CsvReader.GetDataTabletFromCSVFile(ElegenciPalfilepath, oErrorLog);
                Console.WriteLine("Elegenci Pal CSV file has been completed.");
            }

            if (AllowToRunYemek == "true")
            {
                Console.WriteLine("Yemek Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(Yemekfilepath, oErrorLog);
                Console.WriteLine("Yemek CSV file has been completed.");
            }

            if (AllowToRunRO_Master == "true")
            {
                Console.WriteLine("RO Master Data Extraction has been started.");
                var dt1 = CsvReader.GetDTFromROMasterFile(RO_MasterfilePath, oErrorLog);
                Console.WriteLine("RO Master CSV file has been completed.");
            }

            if (AllowToFFM == "true")
            {
                Console.WriteLine("Fixtures Fittings Master Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(FFMfilePath, oErrorLog);
                Console.WriteLine("Fixtures Fittings Master CSV file has been completed.");
            }

            if (AllowToRunMM == "true")
            {
                Console.WriteLine("Measurements Master Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(MMfilePath, oErrorLog);
                Console.WriteLine("Measurements Master CSV file has been completed.");
            }


            if (AllowToRunCM == "true")
            {
                Console.WriteLine("Conditions Master Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(CMfilePath, oErrorLog);
                Console.WriteLine("Conditions Master CSV file has been completed.");
            }

            if (AllowToRunOccupancy == "true")
            {
                Console.WriteLine("Occupancy Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(OccupancyfilePath, oErrorLog);
                Console.WriteLine("Occupancy CSV file has been completed.");
            }

            if (AllowToRunCashflow == "true")
            {
                Console.WriteLine("Cashflow Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(CashflowfilePath, oErrorLog);
                Console.WriteLine("Cashflow CSV file has been completed.");
            }

            if (AllowToRunProfitabilty == "true")
            {
                Console.WriteLine("Profitabilty Data Extraction has been started.");
                var dt1 = CsvReader.GetDataTabletFromCSVFile(ProfitabiltyfilePath, oErrorLog);
                Console.WriteLine("Profitabilty CSV file has been completed.");
            }

            if (AllowToRunPnL == "true")
            {
                Console.WriteLine("Profit and Lost Data Extraction has been started.");
                var dtProfitLost = CsvReader.GetDTFromPLCSVFile(PLfilepath, oErrorLog);
                Console.WriteLine("Profit And Lost CSV file has been completed.");
            }

            if (AllowToFinanceSource2 == "true")
            {
                Console.WriteLine("Finance Source2 Data Extraction has been started.");
                var dtFinanceSource2 = CsvReader.GetDTFromExcelFileSource2(FS2filepath, oErrorLog);
                Console.WriteLine("Finance Source2 Excel file has been completed.");
            }
        }
    }
}

