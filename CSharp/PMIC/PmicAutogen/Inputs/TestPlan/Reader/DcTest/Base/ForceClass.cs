using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class ForceClass
    {
        //Dc setting eg: Level:XXX
        private const string RegLevelSetting = @"(Level:|Levels:)(?<level>[\w]+)";

        //Ac setting eg: AC:XXX:XXX
        private const string RegAcSetting = @"AC:[\w|&]+:[\w]+";

        //Ac Category eg: AC:XXX
        private const string RegAcCategory = @"AC:[\w]+";

        //Ac selector eg: ACSelector:NV:XXX
        private const string RegAcSelector = @"ACSelector:[\w|&]+:[\w]+";

        //Ac Category eg: DC:XXX
        private const string RegDcCategory = @"DC:[\w]+";

        //Dc selector eg: DCSelector:NV:XXX
        private const string RegDcSelector = @"DCSelector:[\w|&]+:[\w]+";

        public ForceClass()
        {
            IsShmooInForce = false;
            IsShmooInProdInst = true;
            IsShmooInProdFlow = true;
            IsShmooInCharInst = false;
            IsShmooInCharFlow = false;
            IsCz2InstName = false;
            ForceCondition = "";
        }

        public bool IsShmooInForce { get; set; }
        public bool IsShmooInProdInst { get; set; }
        public bool IsShmooInProdFlow { get; set; }
        public bool IsShmooInCharInst { get; set; }
        public bool IsShmooInCharFlow { get; set; }
        public bool IsCz2InstName { get; set; }
        public string ForceCondition { get; set; }

        public string GetLevelSetting()
        {
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase))
                    return force.Split(':')[1];
            return "";
        }

        public string GetAcSetting()
        {
            var acSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                    if (IsAcSpecPin(Regex.Replace(force, "::", "&").Split(':')[1]))
                        acSettings += force + ";";
            return acSettings.Trim(';');
        }

        private bool IsAcSpecPin(string pinName)
        {
            //return Regex.IsMatch(pinName, @"^(TCK|ShiftIn)$") ||
            //       NwireSingleton.Instance().SettingInfo.NwirePins.Find(s => s.OutClk.Equals(pinName, StringComparison.OrdinalIgnoreCase))
            //        != null;

            return Regex.IsMatch(pinName, @"^(TCK|ShiftIn)$");
        }

        public string GetAcSelector()
        {
            var acAcSelector = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase))
                    acAcSelector += force + ";";
            return acAcSelector.Trim(';');
        }

        public string GetDcCategory()
        {
            var dcSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase))
                    dcSettings += force + ";";
            return dcSettings.Trim(';');
        }

        public string GetAcCategory()
        {
            var acSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcCategory, RegexOptions.IgnoreCase))
                    acSettings += force + ";";
            return acSettings.Trim(';');
        }

        public string GetDcSelector()
        {
            var dcAcSelector = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase))
                    dcAcSelector += force + ";";
            return dcAcSelector.Trim(';');
        }

        public string GetMcgSetting()
        {
            var mcgSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                    if (!IsAcSpecPin(Regex.Replace(force, "::", "&").Split(':')[1]))
                        mcgSettings += force + ";";
            return mcgSettings.Trim(';');
        }

        public string GetPrePatForceCondition()
        {
            // Remove Dc setting, Ac setting, and Mcg setting
            var forceList = ForceCondition.Split(';').ToList();
            forceList.RemoveAll(string.IsNullOrEmpty);
            foreach (var force in forceList.ToArray())
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcCategory, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase))
                    forceList.Remove(force);
            return string.Join(";", forceList);
        }
    }
}