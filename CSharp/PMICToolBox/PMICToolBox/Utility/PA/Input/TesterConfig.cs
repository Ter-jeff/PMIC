using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace PmicAutomation.Utility.PA.Input
{
    public static class UflexConfigReader
    {
        public static UflexConfig GetXml(string configName)
        {
            XmlSerializer cfx = new XmlSerializer(typeof(UflexConfig));
            try
            {
                using (StreamReader cfgXml = new StreamReader(configName))
                {
                    UflexConfig testerCfg = (UflexConfig)cfx.Deserialize(cfgXml);
                    testerCfg.IO = Regex.Replace(testerCfg.IO, @"\s+", "");
                    testerCfg.HexVS = Regex.Replace(testerCfg.HexVS, @"\s+", "");
                    testerCfg.UVS256 = Regex.Replace(testerCfg.UVS256, @"\s+", "");
                    testerCfg.UVI80 = Regex.Replace(testerCfg.UVI80, @"\s+", "");
                    testerCfg.Support = Regex.Replace(testerCfg.Support, @"\s+", "");
                    testerCfg.UltraPAC = Regex.Replace(testerCfg.UltraPAC, @"\s+", "");
                    testerCfg.DC30 = Regex.Replace(testerCfg.DC30, @"\s+", "");
                    testerCfg.US10G = Regex.Replace(testerCfg.US10G, @"\s+", "");
                    return testerCfg;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Read config file error: " + ex.ToString());
            }
        }
    }

    public class UflexConfig
    {
        public string DC30 = "";
        public string HexVS = "";
        public string IO = "";
        public string Support = "";
        public string UltraPAC = "";
        public string US10G = "";
        public string UVI80 = "";
        public string UVS256 = "";

        public string GetToolTypeByChannelAssignment(string channelAssignment)
        {
            if (channelAssignment != null && Regex.IsMatch(channelAssignment,
                    @"\.ch|\.sense|\.util|.SrcPos|.SrcNeg|.cappos|.capneg", RegexOptions.IgnoreCase))
            {
                return GetToolType(channelAssignment.Split('.')[0]);
            }

            return "";
        }

        public string GetToolTypeByChannelAssignment(List<string> channelAssignments)
        {
            List<string> toolType = channelAssignments
                .Where(x => !x.Equals("SITE0", StringComparison.OrdinalIgnoreCase))
                .Select(GetToolTypeByChannelAssignment).ToList();
            if (toolType.Distinct().Count() == 1 && toolType.Count > 0)
            {
                return toolType.First();
            }

            return "";
        }

        public string GetToolType(string channel)
        {
            if (IO.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "I/O";
            }

            if (HexVS.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "HexVS";
            }

            if (UVS256.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "UVS256";
            }

            if (UVI80.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "UVI80";
            }

            if (UltraPAC.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "UltraPAC";
            }

            if (DC30.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "DC30";
            }

            if (US10G.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "US10G";
            }

            if (Support.Split(',').Any(x => x.Equals(channel, StringComparison.OrdinalIgnoreCase)))
            {
                return "Support";
            }

            return "";
        }
    }
}