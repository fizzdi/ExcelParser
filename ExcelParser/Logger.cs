using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    static class Logger
    {
        private static string GetFileName()
        {
            if (!Directory.Exists("Logs"))
                Directory.CreateDirectory("Logs");
            return "Logs\\log " + DateTime.Now.ToString("dd.MM.yyyy") + ".txt";
        }
        private static string GetData()
        {
            return DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        }
        public static void LogMessage(String TypeMessage, String Message)
        {
            using (StreamWriter writer = new StreamWriter(GetFileName(), true))
            {
                writer.WriteLine(GetData() + " " + TypeMessage + ": " + Message);
            }
        }
        public static void LogMessage(Exception ex)
        {
            using (StreamWriter writer = new StreamWriter(GetFileName(), true))
            {
                writer.WriteLine(GetData() + " " + ex.GetType().Name + ": " + ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }
    }
}
