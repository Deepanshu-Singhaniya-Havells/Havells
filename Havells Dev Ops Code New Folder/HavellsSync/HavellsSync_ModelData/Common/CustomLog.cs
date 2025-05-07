using HavellsSync_ModelData.ICommon;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace HavellsSync_ModelData.Common
{
    public class CustomLog : ICustomLog
    {
        public CustomLog(IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(configuration), configuration);
        }
        public void LogToFile(Exception ex)
        {
            DateTime dateTime = DateTime.Now;
            DateTime utcDateTime = dateTime.ToUniversalTime();
            DateTime istDateTime = utcDateTime.AddHours(5).AddMinutes(30);
            string logDir = Path.Combine(Environment.CurrentDirectory, "LogDetails");
            string logFile = $"{dateTime.Date.ToString("dd-MM-yyyy")}_APILogs.txt";
            string logPath = Path.Combine(logDir, logFile);
            try
            {
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                using (var fs = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.WriteLine($"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff")}|{ex.Message}{(ex.InnerException != null ? $"~{ex.InnerException.Message}" : string.Empty)}");
                    }
                }
            }
            catch (Exception excp)
            {
                using (var fs = new FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.WriteLine($"{istDateTime.ToString("dd/MM/yyyy hh:mm:ss.fff")}|{excp.Message}");
                    }
                }
            }
        }
    }
}
