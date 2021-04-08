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

            string PLfilepath = ConfigurationManager.AppSettings["ProfitLossfilePath"];
            string AllowToRunAura = ConfigurationManager.AppSettings["AllowToRunAura"];
            string AllowToRunPnL = ConfigurationManager.AppSettings["AllowToRunProfitLost"];
            string AllowToRunAG = ConfigurationManager.AppSettings["AllowToRunAuraGroup"];
            string AllowToRunEP = ConfigurationManager.AppSettings["AllowToRunElegenciPal"];
            string AllowToRunYemek = ConfigurationManager.AppSettings["AllowToRunYemek"];


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

            if (AllowToRunPnL == "true")
            {
                Console.WriteLine("Profit and Lost Data Extraction has been started.");
                var dtProfitLost = CsvReader.GetDTFromPLCSVFile(PLfilepath, oErrorLog);
                Console.WriteLine("Profit And Lost CSV file has been completed.");

            }
        }
    }
}

