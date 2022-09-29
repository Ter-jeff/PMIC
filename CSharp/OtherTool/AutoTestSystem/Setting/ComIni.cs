using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoTestSystem.Setting
{
    public class ComIni
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string Default,
            StringBuilder retVal, int size, string filePath);

        public string IniRead(string path, string key, string section = null)
        {
            var retVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", retVal, 255, path);
            return retVal.ToString();
        }

        public void IniWrite(string path, string key, string value, string section)
        {
            WritePrivateProfileString(section, key, value, path);
        }

        public void IniDeleteKey(string path, string key, string section)
        {
            IniWrite(path, key, null, section);
        }

        public void DeleteSection(string path, string section)
        {
            IniWrite(path, null, null, section);
        }

        public bool IniKeyExists(string path, string key, string section)
        {
            return IniRead(path, key, section).Length > 0;
        }

        public Dictionary<string, List<IniRow>> Read(string file)
        {
            var dic = new Dictionary<string, List<IniRow>>();
            var lines = File.ReadAllLines(file).ToList();
            var iniRows = new List<IniRow>();
            var groupName = "";
            foreach (var line in lines)
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    if (!string.IsNullOrEmpty(groupName))
                        dic.Add(groupName, iniRows);
                    iniRows = new List<IniRow>();
                    groupName = line.TrimStart('[').TrimEnd(']');
                }
                else
                {
                    var arr = line.Split('=');
                    iniRows.Add(new IniRow { Name = arr.First(), Value = arr.Last() });
                }

            if (!string.IsNullOrEmpty(groupName) && iniRows.Count() != 0)
                dic.Add(groupName, iniRows);
            return dic;
        }
    }

    public class IniRow
    {
        public string Name;
        public string Value;
    }
}