using CommonLib.Enum;
using NLog;
using System;
using System.IO;
using System.Linq;

namespace AutoProgram.Reader
{
    internal class DlexChecker
    {
        public bool Check(string dlex)
        {
            var dlexUserName = "";
            var lines = File.ReadAllLines(dlex);
            if (lines.Any())
            {
                var arr = lines.First().Split('|');
                if (arr.Count() > 2)
                    dlexUserName = arr[2];
            }

            var user = Environment.UserName;
            if (!user.Equals(dlexUserName, StringComparison.CurrentCultureIgnoreCase))
            {
                var logger = LogManager.GetCurrentClassLogger();
                logger.Error("[" + EnumNLogMessage.Input + "] " + dlexUserName +
                             @" user name in dlex is not the same with current user name " + user + " !!!");
                return true;
            }

            return false;
        }

        public void Modify(string dlex, string log)
        {
            var user = Environment.UserName;
            var lines = File.ReadAllLines(dlex);
            if (lines.Any())
            {
                var arr = lines.First().Split('|');
                if (arr.Count() > 2)
                    arr[2] = user;
                if (arr.Count() > 19)
                    arr[19] = log;
                lines[0] = string.Join("|", arr);
            }

            File.WriteAllLines(dlex, lines);
        }
    }
}