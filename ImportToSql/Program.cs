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
            string path = ConfigurationManager.AppSettings["filePath"];

            oErrorLog.WriteErrorLog(" ");
            oErrorLog.WriteErrorLog("----------------------------------------");
            oErrorLog.WriteErrorLog("Import task starting...");
            oErrorLog.WriteErrorLog("Open CSV file to read data");

            var dt = CsvReader.GetDataTabletFromCSVFile(path, oErrorLog);
            Console.ReadLine();
        }
    }
}

