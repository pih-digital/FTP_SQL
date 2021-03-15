using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;

namespace ImportToSql
{
    public class ErrorLog
    {
        public bool WriteErrorLog(string LogMessage)
        {
            bool Status = false;
            string LogDirectory = ConfigurationManager.AppSettings["LogDirectory"].ToString();

            DateTime CurrentDateTime = DateTime.Now;
            CheckCreateLogDirectory(LogDirectory);
            string logLine = BuildLogLine(CurrentDateTime, LogMessage);
            LogDirectory = LogDirectory + "Log_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";

            lock (typeof(ErrorLog))
            {
                StreamWriter oStreamWriter = null;
                try
                {
                    if (!System.IO.File.Exists(LogDirectory))
                    {
                        System.IO.File.Create(LogDirectory);
                    }
                    oStreamWriter = new StreamWriter(LogDirectory, true);
                    oStreamWriter.WriteLine(logLine);
                    Status = true;
                }
                catch (Exception ex)
                {
                    // Logging failure
                }
                finally
                {
                    if (oStreamWriter != null)
                    {
                        oStreamWriter.Close();
                    }
                }
            }
            return Status;
        }


        private bool CheckCreateLogDirectory(string LogPath)
        {
            bool loggingDirectoryExists = false;
            DirectoryInfo oDirectoryInfo = new DirectoryInfo(LogPath);
            if (oDirectoryInfo.Exists)
            {
                loggingDirectoryExists = true;
            }
            else
            {
                try
                {
                    Directory.CreateDirectory(LogPath);
                    loggingDirectoryExists = true;
                }
                catch (Exception ex)
                {
                    // Logging failure
                }
            }
            return loggingDirectoryExists;
        }

        private string BuildLogLine(DateTime CurrentDateTime, string LogMessage)
        {
            StringBuilder loglineStringBuilder = new StringBuilder();
            loglineStringBuilder.Append(CurrentDateTime.ToString("dd-MM-yyyy HH:mm:ss"));
            loglineStringBuilder.Append(" \t");
            loglineStringBuilder.Append(LogMessage);
            return loglineStringBuilder.ToString();
        }
    }
}