using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace IgxlData.Others
{
    public class TesterConfig
    {
        public static UflexConfig GetCongXml(string configName = "")
        {
            var cfgFile = configName != "" ? configName : Directory.GetCurrentDirectory() + "\\Config\\Tester\\" + "TesterConfig_Default.xml";
            if (!File.Exists(cfgFile))
                return null;

            var cfx = new XmlSerializer(typeof(UflexConfig));
            var cfgXml = new StreamReader(cfgFile);
            try
            {
                var testerCfg = (UflexConfig)cfx.Deserialize(cfgXml);
                testerCfg.Vsm = testerCfg.Vsm==null?"":Regex.Replace(testerCfg.Vsm, @"\s+", "");
                testerCfg.Io = testerCfg.Io == null ? "" : Regex.Replace(testerCfg.Io, @"\s+", "");
                testerCfg.HexVs = testerCfg.HexVs == null ? "" : Regex.Replace(testerCfg.HexVs, @"\s+", "");
                testerCfg.Uvs256 = testerCfg.Uvs256 == null ? "" : Regex.Replace(testerCfg.Uvs256, @"\s+", "");
                testerCfg.Uvs64 = testerCfg.Uvs64 == null ? "" : Regex.Replace(testerCfg.Uvs64, @"\s+", "");
                testerCfg.Uvi80 = testerCfg.Uvi80 == null ? "" : Regex.Replace(testerCfg.Uvi80, @"\s+", "");
                testerCfg.Support = testerCfg.Support == null ? "" : Regex.Replace(testerCfg.Support, @"\s+", "");
                testerCfg.UltraPac = testerCfg.UltraPac == null ? "" : Regex.Replace(testerCfg.UltraPac, @"\s+", "");
                testerCfg.Dc30 = testerCfg.Dc30 == null ? "" : Regex.Replace(testerCfg.Dc30, @"\s+", "");
                testerCfg.Us10G = testerCfg.Us10G == null ? "" : Regex.Replace(testerCfg.Us10G, @"\s+", "");
                return testerCfg;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            return null;
        }
    }

    public class UflexConfig
    {
        public string Vsm;
        public string Io;
        public string HexVs;
        public string Uvs64;
        public string Uvs256;
        public string Uvi80;
        public string UltraPac;
        public string Dc30;
        public string Us10G;
        public string Support;

        public string GetToolTypeByChannelAssignment(string channelAssignment)
        {
            //if (channelAssignment != null && Regex.IsMatch(channelAssignment, @"\.ch|\.sense|\.util|.SrcPos|.SrcNeg|.cappos|.capneg", RegexOptions.IgnoreCase))
            if (channelAssignment != null)
                return GetToolType(channelAssignment.Split('.')[0]);
            return "";
        }

        public string GetToolTypeByChannelAssignment(List<string> channelAssignments)
        {
            List<string> toolType = channelAssignments.Where(x => !x.Equals("SITE0", StringComparison.OrdinalIgnoreCase))
                .Select(GetToolTypeByChannelAssignment).ToList();
            if (toolType.Distinct().Count() == 1 && toolType.Count > 0) return toolType.First();
            return "";
        }

        public string GetToolType(string ch)
        {

            if (Vsm.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "VSM";
            if (Io.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "I/O";
            if (HexVs.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "HexVS";
            if (Uvs256.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVS256";
            if (Uvs64.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVS64";
            if (Uvi80.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVI80";
            if (UltraPac.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UltraPAC";
            if (Dc30.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "DC30";
            if (Us10G.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "US10G";
            if (Support.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "Support";
            return "";
        }
    }
}
