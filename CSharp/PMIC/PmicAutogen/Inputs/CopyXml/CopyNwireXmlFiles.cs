using CommonLib.Enum;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
    public class CopyXmlFiles
    {
        public void Work()
        {
            var targetDir = FolderStructure.DirXmlFiles;
            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
                if (resourceName.Contains(".Protocol."))
                {
                    var arr = resourceName.Split('.').ToList();
                    var fileName = targetDir + "\\" + string.Join(".", arr.GetRange(arr.Count - 2, 2));
                    using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        var resourceStream = assembly.GetManifestResourceStream(resourceName);
                        if (resourceStream != null) resourceStream.CopyTo(file);
                    }
                }
        }
    }

    public class CopyIgxlConfigFiles
    {
        public void Work()
        {
            var targetDir = FolderStructure.DirIgLink;
            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            var fileName = string.Empty;
            var lStrResource = string.Empty;


            var lStrXmlConfig = string.Empty;
            var lStrConfigTemplate = string.Empty;

            foreach (var resourceName in resourceNames)
                if (resourceName.Contains(".IGXLConfig."))
                {
                    if (resourceName.Contains("SimulatedConfig.txt"))
                    {
                        lStrResource = resourceName;
                        var arr = resourceName.Split('.').ToList();
                        fileName = targetDir + "\\" + string.Join(".", arr.GetRange(arr.Count - 2, 2));
                    }
                    else if (resourceName.Contains("SimulatedConfigTemplate.tmp"))
                    {
                        lStrConfigTemplate = resourceName;
                    }
                    else if (resourceName.Contains("SimulatedConfigTypeMapping.xml"))
                    {
                        lStrXmlConfig = resourceName;
                    }
                }

            TesterConfigTypeItem lConfigTypeItem = null;
            try
            {
                var lStreamConfigTemplate = assembly.GetManifestResourceStream(lStrXmlConfig);
                var xs = new XmlSerializer(typeof(TesterConfigTypeItem));
                lConfigTypeItem = (TesterConfigTypeItem)xs.Deserialize(lStreamConfigTemplate);
                lStreamConfigTemplate.Close();

                Dictionary<int, string> lDic = null;
                var lCheckPinName = CheckPinNameType(lConfigTypeItem);
                var lCheckSlotType = CheckSlotType(lConfigTypeItem, out lDic);

                if (lCheckPinName == false || lCheckSlotType == false)
                {
                    Response.Report(
                        "The tool will generate a default tester config because there is something incorrect definition in the channel map.",
                        EnumMessageLevel.Error, 0);
                    using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        var resourceStream = assembly.GetManifestResourceStream(lStrResource);
                        if (resourceStream != null) resourceStream.CopyTo(file);
                    }
                }
                else
                {
                    var lStrDate = TimeProvider.Current.Now.ToString("yyyy-MM-dd");
                    var lStrHour = TimeProvider.Current.Now.ToLongTimeString();
                    var lStrContent = string.Empty;
                    var dic1Asc = lDic.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);
                    foreach (var k in dic1Asc)
                        lStrContent = lStrContent + k.Key + lConfigTypeItem.GetContentByPinType(k.Value);

                    var lTemplateStream = assembly.GetManifestResourceStream(lStrConfigTemplate);
                    var lTemplate = new byte[lTemplateStream.Length];
                    lTemplateStream.Read(lTemplate, 0, lTemplate.Length);
                    var lStrTemplate = Encoding.UTF8.GetString(lTemplate);
                    var lStrOutput = string.Format(lStrTemplate, lStrDate, lStrHour, lStrDate, lStrHour, lStrContent);

                    File.WriteAllText(fileName, lStrOutput);
                }
            }
            catch (Exception e)
            {
                Response.Report(e.ToString(), EnumMessageLevel.Error, 0);
                Response.Report("Generate tester config file failed!", EnumMessageLevel.Error, 0);
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="p_ConfigTypeItem"></param>
        /// <returns></returns>
        private bool CheckPinNameType(TesterConfigTypeItem pConfigTypeItem)
        {
            var lRtn = true;

            var lChannelMapSheet = StaticTestPlan.ChannelMapSheets[0];

            foreach (var lRow in lChannelMapSheet.ChannelMapRows)
            {
                var lIsValidRow = pConfigTypeItem.IsValidPinNameType(lRow);
                if (lIsValidRow == false)
                {
                    Response.Report(
                        "The pin name: " + lRow.DeviceUnderTestPinName + " can not match the type: " + lRow.Type,
                        EnumMessageLevel.Error, 0);
                    lRtn = false;
                }
            }

            return lRtn;
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        private bool CheckSlotType(TesterConfigTypeItem pConfigTypeItem, out Dictionary<int, string> pDic)
        {
            var lRtn = true;
            pDic = new Dictionary<int, string>(); //key slot number,value is type

            var lChannelMapSheet = StaticTestPlan.ChannelMapSheets[0];

            foreach (var lRow in lChannelMapSheet.ChannelMapRows)
            {
                var lStrSlotType = pConfigTypeItem.GetTesterConfigTypeByPinAndPinType(lRow);

                //if slot type is null. it shold not be dc30, uvi80 and up1600,igore
                if (!string.IsNullOrEmpty(lStrSlotType))
                {
                    var lLstSlotNumber = new List<string>();
                    foreach (var lStrSite in lRow.Sites)
                    {
                        var lIntSiteNumber = GetSiteNumber(lStrSite);
                        if (lIntSiteNumber != -1)
                        {
                            if (!pDic.ContainsKey(lIntSiteNumber))
                            {
                                pDic.Add(lIntSiteNumber, lStrSlotType);
                            }
                            else if (!pDic[lIntSiteNumber].Equals(lStrSlotType))
                            {
                                Response.Report("Slot Number: " + lIntSiteNumber +
                                                " should not be defined with two different Pin Type[" +
                                                pDic[lIntSiteNumber] + "|" + lStrSlotType + "]", EnumMessageLevel.Error, 0);
                                lRtn = false;
                            }
                        }
                    }
                }
            }

            return lRtn;
        }

        /// <summary>
        /// </summary>
        /// <param name="p_strSite"></param>
        /// <returns></returns>
        private int GetSiteNumber(string pStrSite)
        {
            var lRtn = -1;

            var lResult = 0;


            var lArySite = pStrSite.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries);

            if (lArySite.Length > 0)
                if (int.TryParse(lArySite[0], out lResult))
                {
                    lRtn = lResult;
                    return lRtn;
                }

            return lRtn;
        }
    }
}