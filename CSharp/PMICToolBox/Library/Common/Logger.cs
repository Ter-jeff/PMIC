using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library.Common
{
    public class Logger
    {
        public static void PrintLog(string logFile, string msg)
        {
            if (logFile == null || !Directory.Exists(new FileInfo(logFile).DirectoryName))
                return;
            string currentTime = DateTime.Now.ToLocalTime().ToString();
            StreamWriter sw = File.AppendText(logFile);
            sw.WriteLine(currentTime + ": " + msg);
            sw.Flush();
            sw.Close();
        }
    }
}
