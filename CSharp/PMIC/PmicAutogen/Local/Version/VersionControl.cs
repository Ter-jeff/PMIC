using CommonLib.Utility;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace PmicAutogen.Local.Version
{
    public class VersionControl
    {
        private static readonly string ToolVersion = "V" + Assembly.GetExecutingAssembly().GetName().Version;
        public static List<SrcInfoRow> SrcInfoRows = new List<SrcInfoRow>();
        private static string TimeStamp
        {
            get { return ToolVersion + "_" + TimeProvider.Current.Now.ToString("yyyy-MM-dd HHmmss"); }
        }

        public static void Initialize()
        {
            SrcInfoRows = new List<SrcInfoRow>();
        }

        public static string AddTimeStamp(string file)
        {
            return Path.GetFileNameWithoutExtension(file) + "_Real_" + TimeStamp + Path.GetExtension(file);
        }
    }
}