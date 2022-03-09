using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace IgxlData.Others
{
    public class TesterConfigManager
    {
        public Dictionary<string, TesterConfig> TesterConfigs { get; set; }

        public string GetToolTypeByChannelAssignment(string channelAssignment, string sheetName)
        {
            if (channelAssignment != null)
            {
                var testerConfig = GetTesterConfigs(sheetName);
                if (testerConfig == null)
                    return "";
                return testerConfig.GetToolType(channelAssignment.Split('.')[0]);
            }
            return "";
        }

        private TesterConfig GetTesterConfigs(string sheetName)
        {
            var name = sheetName.ToUpper();
            if (name.Contains("_CP") && TesterConfigs.ContainsKey("CP"))
                return TesterConfigs["CP"];
            if (name.Contains("_FT") && TesterConfigs.ContainsKey("FT"))
                return TesterConfigs["FT"];
            if (TesterConfigs.ContainsKey("All"))
                return TesterConfigs["All"];
            if (TesterConfigs.ContainsKey("CP"))
                return TesterConfigs["CP"];
            return null;
        }

        public string GetToolTypeByChannelAssignment(List<string> channelAssignments, string sheetName)
        {
            var toolType = channelAssignments.Where(x => !x.Equals("SITE0", StringComparison.OrdinalIgnoreCase))
                .Select(x => GetToolTypeByChannelAssignment(x, sheetName)).ToList();
            if (toolType.Distinct().Count() == 1 && toolType.Count > 0) return toolType.First();
            return "";
        }

        public List<string> GetHex(string sheetName)
        {
            var testerConfig = GetTesterConfigs(sheetName);
            if (testerConfig == null)
                return new List<string>();
            return testerConfig.HexVS.Split(',').ToList();
        }
    }


    public class TesterConfigReader
    {
        public static TesterConfigManager GetTesterConfigs(string configName = "")
        {
            var cfgFile = configName != "" ? configName : Directory.GetCurrentDirectory() + "\\Config\\Tester\\" + "TesterConfig_Default.xml";
            if (!File.Exists(cfgFile))
                return null;

            var doc = new XmlDocument();
            var settings = new XmlReaderSettings { IgnoreComments = true };
            var reader = XmlReader.Create(configName, settings);
            doc.Load(reader);
            var dic = new Dictionary<string, TesterConfig>(StringComparer.CurrentCultureIgnoreCase);
            var configNode = doc.SelectSingleNode("UflexConfig");
            if (configNode != null)
            {
                var nodeCp = configNode.SelectSingleNode("CP");
                if (nodeCp != null)
                {
                    var testerCfg = GetUflexConfig(nodeCp);
                    dic.Add("CP", testerCfg);
                }

                var nodeFt = configNode.SelectSingleNode("FT");
                if (nodeFt != null)
                {
                    var testerCfg = GetUflexConfig(nodeFt);
                    dic.Add("FT", testerCfg);
                }

                if (nodeCp != null && nodeFt != null)
                {
                    var testerCfgAll = GetUflexConfig(configNode);
                    dic.Add("All", testerCfgAll);
                }
            }
            var testerConfigManager = new TesterConfigManager();
            testerConfigManager.TesterConfigs = dic;
            return testerConfigManager;
        }

        private static TesterConfig GetUflexConfig(XmlNode configNode)
        {
            var testerCfg = new TesterConfig();
            var node = configNode.SelectSingleNode("VSM");
            if (node != null)
                testerCfg.VSM = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("IO");
            if (node != null)
                testerCfg.IO = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("HexVS");
            if (node != null)
                testerCfg.HexVS = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("UVS256");
            if (node != null)
                testerCfg.UVS256 = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("UVS64");
            if (node != null)
                testerCfg.UVS64 = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("UVI80");
            if (node != null)
                testerCfg.UVI80 = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("Support");
            if (node != null)
                testerCfg.Support = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("UltraPAC");
            if (node != null)
                testerCfg.UltraPAC = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("DC30");
            if (node != null)
                testerCfg.DC30 = Regex.Replace(node.InnerText, @"\s+", "");
            node = configNode.SelectSingleNode("US10G");
            if (node != null)
                testerCfg.US10G = Regex.Replace(node.InnerText, @"\s+", "");
            return testerCfg;
        }
    }

    public class TesterConfig
    {
        public string VSM = "";
        public string IO = "";
        public string HexVS = "";
        public string UVS64 = "";
        public string UVS256 = "";
        public string UVI80 = "";
        public string UltraPAC = "";
        public string DC30 = "";
        public string US10G = "";
        public string Support = "";

        public string GetToolType(string ch)
        {

            if (VSM.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "VSM";
            if (IO.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "I/O";
            if (HexVS.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "HexVS";
            if (UVS256.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVS256";
            if (UVS64.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVS64";
            if (UVI80.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UVI80";
            if (UltraPAC.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "UltraPAC";
            if (DC30.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "DC30";
            if (US10G.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "US10G";
            if (Support.Split(',').Any(x => x.Equals(ch, StringComparison.OrdinalIgnoreCase)))
                return "Support";
            return "";
        }
    }
}
