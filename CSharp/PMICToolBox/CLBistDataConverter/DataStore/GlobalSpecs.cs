using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CLBistDataConverter
{
    static class GlobalSpecs
    {
        public static string LogFile = "Log.txt";

        public static void initialize()
        {
            LogFile = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, LogFile);
        }
    }
}
