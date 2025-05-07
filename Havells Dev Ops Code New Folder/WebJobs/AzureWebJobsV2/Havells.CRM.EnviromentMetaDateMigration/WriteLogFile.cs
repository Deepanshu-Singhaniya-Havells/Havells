using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.EnviromentMetaDateMigration
{
    public class WriteLogFile
    {
        public static bool WriteLog(string strMessage)
        {
            try
            {
                string strFileName = "_InstanceMigration_Logs_"+DateTime.Today.ToString("yyyy/MM/dd")+".txt";
                FileStream objFilestream = new FileStream(string.Format("{0}\\{1}", "C:\\Logs\\", strFileName), FileMode.Append, FileAccess.Write);
                StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
                objStreamWriter.WriteLine(strMessage);
                objStreamWriter.Close();
                objFilestream.Close();
                return true;
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog(ex.Message);
                return false;
            }
        }
    }
}
