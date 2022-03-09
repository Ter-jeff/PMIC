using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CLBistDataConverter
{
    public class ComLib
    {
        public static string LogTimeStemp()
        {
            return DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        }

        public static string TimeStemp()
        {
            return DateTime.Now.ToString("yyyyMMdd_hhmmss");
        }

        public static void WriteLog(string log)
        {
            StreamWriter sw = File.AppendText(GlobalSpecs.LogFile);
            sw.WriteLine(log);
            sw.Flush();
            sw.Close();
        }
    }
}
