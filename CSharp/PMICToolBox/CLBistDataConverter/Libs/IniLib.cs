using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace CLBistDataConverter.Libs
{
    public class IniLib
    {
        public string _IniFilePath = "";

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        public string IniRead(string key, string section)
        {
            var retVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", retVal, 255, _IniFilePath);
            return retVal.ToString();
        }

        public void IniWrite(string key, string value, string section)
        {
            WritePrivateProfileString(section, key, value, _IniFilePath);
        }

        public void IniDeleteKey(string key, string section)
        {
            IniWrite(key, null, section);
        }

        public void DeleteSection(string section)
        {
            IniWrite(null, null, section);
        }

        public bool IniKeyExists(string key, string section)
        {
            return IniRead(key, section).Length > 0 ? true : false;
        }
    }
}
