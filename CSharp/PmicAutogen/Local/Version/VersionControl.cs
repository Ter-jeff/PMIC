using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

namespace PmicAutogen.Local.Version
{
    public class VersionControl
    {
        #region Field

        public static string ToolVersion = "V" + Assembly.GetExecutingAssembly().GetName().Version;

        #endregion

        public static void Initialize()
        {
            SrcInfoRows = new List<SrcInfoRow>();
        }

        #region Property

        public static string TimeStamp => "_" + ToolVersion + "_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss");

        public static string AddTimeStamp(string file)
        {
            return Path.GetFileNameWithoutExtension(file) + "_Real" + TimeStamp + Path.GetExtension(file);
        }

        public static List<SrcInfoRow> SrcInfoRows = new List<SrcInfoRow>();

        #endregion
    }

    public class ClassUtility
    {
        public static string GetFileMd5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var txt = "";
                    foreach (var hash in md5.ComputeHash(stream))
                        txt = string.Format("{0}{1:x2}", txt, hash);
                    return txt;
                }
            }
        }
    }
}