//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path.
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using Library.DataStruct;

namespace Library.Common
{
    public class SettingLib
    {
        public static Settings ReadSetting()
        {
            var appSettings = new Settings();

            var xmlDocument = new XmlDocument();
            // 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path. chg start
            //xmlDocument.Load(Environment.CurrentDirectory + @"\Settings\ParserSettings.xml");
            var settingPath = Environment.CurrentDirectory + @"\Settings\ParserSettings.xml";
            if (!System.IO.File.Exists(settingPath))
                settingPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Settings\ParserSettings.xml";

            xmlDocument.Load(settingPath);
            // 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path. chg end

            var rootNode = xmlDocument.SelectSingleNode("Settings");

            if (rootNode == null) return appSettings;
            foreach (var settingNode in rootNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() != typeof(XmlNode)).Cast<XmlNode>())
            {
                switch (settingNode.Name)
                {
                    case "ItemDefine":
                        appSettings.ItemDefine.AddRange(LoadItemDefine(settingNode));
                        break;
                    case "HeaderPatterns":
                        appSettings.HeaderPatterns.AddRange(LoadHeaderPattern(settingNode));
                        break;
                    case "LogRowType":
                        appSettings.LogRowTypePatterns.AddRange(LoadLogRowTypePattern(settingNode));
                        break;
                    case "IgnoredItem":
                        appSettings.IgnoredItemPatterns.AddRange(LoadIgnoredItemPattern(settingNode));
                        break;
                }
            }
            return appSettings;
        }

        private static IEnumerable<ItemInfo> LoadItemDefine(XmlNode itemDefineNode)
        {
            var itemDefine = new List<ItemInfo>();
            foreach (XmlElement itemNode in itemDefineNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
            {
                var itemInfo = new ItemInfo {Name = itemNode.GetAttribute("name")};
                foreach (XmlElement itemPat in itemNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
                {
                    itemInfo.Patterns.Add(new Regex(itemPat.InnerText));
                }
                itemDefine.Add(itemInfo);
            }

            return itemDefine;
        }

        private static IEnumerable<HeaderPattern> LoadHeaderPattern(XmlNode headerPatternNode)
        {
            var headerPatterns = new List<HeaderPattern>();
            foreach (XmlElement headerNode in headerPatternNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
            {
                var headerPat = new HeaderPattern() { Name = headerNode.GetAttribute("name"), Pattern = headerNode.GetAttribute("pattern"), HeaderReg = new Regex(headerNode.GetAttribute("pattern"), RegexOptions.IgnoreCase) };

                foreach (XmlElement itemPat in headerNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
                {
                    var headItem = new HeaderItem() { Name = itemPat.GetAttribute("name"), Missingpossible = itemPat.GetAttribute("missingpossible") == "true", Pattern = itemPat.InnerText };
                    headerPat.Items.Add(headItem);
                }
                headerPatterns.Add(headerPat);
            }
            return headerPatterns;
        }

        public static void CreateHeaderDataPattern(Settings appSettings)
        {
            foreach (var headerpat in appSettings.HeaderPatterns)
            {
                var dataPat = new StringBuilder();
                dataPat.Append("^");
                foreach (var itemPat in headerpat.Items)
                {
                    string oneColumnReg = string.Format(@"(?<{0}>{1})\s+", itemPat.Name, itemPat.Pattern);
                    if (itemPat.Missingpossible) oneColumnReg = string.Format("({0})?", oneColumnReg);
                    dataPat.Append(oneColumnReg);
                }
                string pat = dataPat.ToString();
                pat = pat.Substring(0, pat.Length - 3) + "$";
                //dataPat.Append("$");
                headerpat.DataRegex = new Regex(pat, RegexOptions.IgnoreCase);
            }
        }

        private static IEnumerable<LogRowTypePattern> LoadLogRowTypePattern(XmlNode logRowTypePatternNode)
        {
            var logRowTypePatterns = new List<LogRowTypePattern>();
            foreach (XmlElement itemNode in logRowTypePatternNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
            {
                var itemInfo = new LogRowTypePattern { Name = itemNode.GetAttribute("name") };
                foreach (XmlElement itemPat in itemNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
                {
                    itemInfo.Pattern = new Regex(itemPat.InnerText,RegexOptions.IgnoreCase);
                }
                logRowTypePatterns.Add(itemInfo);
            }
            return logRowTypePatterns;
        }

        private static IEnumerable<IgnoredItemPattern> LoadIgnoredItemPattern(XmlNode ignoredItemPatternNode)
        {
            var IgnoredItemPatterns = new List<IgnoredItemPattern>();
            foreach (XmlElement itemNode in ignoredItemPatternNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
            {
                var itemInfo = new IgnoredItemPattern { Name = itemNode.GetAttribute("name") };
                foreach (XmlElement itemPat in itemNode.ChildNodes.Cast<object>().Where(childNode => childNode.GetType() == typeof(XmlElement)))
                {
                    itemInfo.Pattern = new Regex(itemPat.InnerText, RegexOptions.IgnoreCase);
                }
                IgnoredItemPatterns.Add(itemInfo);
            }
            return IgnoredItemPatterns;
        }
    }
}
