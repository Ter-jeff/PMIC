using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using PmicAutogen.Local;
using IgxlData.IgxlSheets;
using IgxlData.IgxlBase;
using PmicAutogen.InputPackages;
using AutomationCommon.DataStructure;
using System.Collections.Generic;

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

    public class CopyIGXLConfigFiles
    {
        public void Work()
        {
            var targetDir = FolderStructure.DirIgLink;
            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();

            var fileName = string.Empty;
            var l_strResource = string.Empty;


            string l_strXMLConfig = string.Empty;
            string l_strConfigTemplate = string.Empty;

            foreach (var resourceName in resourceNames)
                if (resourceName.Contains(".IGXLConfig."))
                {
                    if (resourceName.Contains("SimulatedConfig.txt"))
                    {
                        l_strResource = resourceName;
                        var arr = resourceName.Split('.').ToList();
                        fileName = targetDir + "\\" + string.Join(".", arr.GetRange(arr.Count - 2, 2));
                    }
                    else if (resourceName.Contains("SimulatedConfigTemplate.tmp"))
                    {
                        l_strConfigTemplate = resourceName;
                    }
                    else if (resourceName.Contains("SimulatedConfigTypeMapping.xml"))
                    {
                        l_strXMLConfig = resourceName;
                    }
                    else
                    {
                        //do nothing
                    }
                }

            TesterConfigTypeItem l_ConfigTypeItem = null;
            try
            {
                Stream l_StreamConfigTemplate = assembly.GetManifestResourceStream(l_strXMLConfig);
                //FileStream l_FileStream = new FileStream(@"C:\C#\PmicAutogen\PmicAutogen\IGXLConfig\SimulatedConfigTypeMapping.xml", FileMode.Open);
                XmlSerializer xs = new XmlSerializer(typeof(TesterConfigTypeItem));
                l_ConfigTypeItem = (TesterConfigTypeItem)xs.Deserialize(l_StreamConfigTemplate);
                l_StreamConfigTemplate.Close();

                Dictionary<int, string> l_Dic = null;
                bool l_CheckPinName = CheckPinNameType(l_ConfigTypeItem);
                bool l_CheckSlotType = CheckSlotType(l_ConfigTypeItem, out l_Dic);

                if (l_CheckPinName == false || l_CheckSlotType == false)
                {
                    Response.Report("The tool will generate a default tester config because there is something incorrect definition in the channel map.", MessageLevel.Error, 0);
                    using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                    {
                        var resourceStream = assembly.GetManifestResourceStream(l_strResource);
                        if (resourceStream != null) resourceStream.CopyTo(file);
                    }
                }
                else
                {
                    string l_strDate = DateTime.Now.ToString("yyyy-MM-dd");
                    string l_strHour = DateTime.Now.ToLongTimeString().ToString();
                    string l_strContent = string.Empty;
                    Dictionary<int, string> dic1Asc = l_Dic.OrderBy(o => o.Key).ToDictionary(o => o.Key, p => p.Value);
                    foreach (KeyValuePair<int, string> k in dic1Asc)
                    {
                        l_strContent = l_strContent + k.Key + l_ConfigTypeItem.GetContentByPinType(k.Value);
                    }

                    Stream l_TemplateStream = assembly.GetManifestResourceStream(l_strConfigTemplate);
                    byte[] l_Template = new byte[l_TemplateStream.Length];
                    l_TemplateStream.Read(l_Template, 0, l_Template.Length);
                    string l_strTemplate = System.Text.Encoding.UTF8.GetString(l_Template);
                    //string l_strTemplate = File.ReadAllText(@"C:\C#\PmicAutogen\PmicAutogen\IGXLConfig\SimulatedConfigTemplate.tmp");
                    string l_strOutput = string.Format(l_strTemplate, l_strDate, l_strHour, l_strDate, l_strHour, l_strContent);

                    File.WriteAllText(fileName, l_strOutput);

                }
            }
            catch (Exception e)
            {
                Response.Report(e.ToString(), MessageLevel.Error, 0);
                Response.Report("Generate tester config file failed!", MessageLevel.Error, 0);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_ConfigTypeItem"></param>
        /// <returns></returns>
        private bool CheckPinNameType(TesterConfigTypeItem p_ConfigTypeItem)
        {
            bool l_Rtn = true;

            ChannelMapSheet l_ChannelMapSheet = StaticTestPlan.ChannelMapSheets[0];

            foreach (ChannelMapRow l_Row in l_ChannelMapSheet.ChannelMapRows)
            {
                bool l_IsValidRow = p_ConfigTypeItem.IsValidPinNameType(l_Row);
                if (l_IsValidRow == false)
                {
                    Response.Report("The pin name: "+l_Row.DeviceUnderTestPinName+" can not match the type: "+l_Row.Type, MessageLevel.Error, 0);
                    l_Rtn = false;
                }
                else
                {
                     //do nothing
                }
            }

            return l_Rtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool CheckSlotType(TesterConfigTypeItem p_ConfigTypeItem,out Dictionary<int,string> p_Dic)
        {
            bool l_Rtn = true;
            p_Dic = new Dictionary<int, string>();//key slot number,value is type

            ChannelMapSheet l_ChannelMapSheet = StaticTestPlan.ChannelMapSheets[0];

            foreach (ChannelMapRow l_Row in l_ChannelMapSheet.ChannelMapRows)
            {
                string l_strSlotType = p_ConfigTypeItem.GetTesterConfigTypeByPinAndPinType(l_Row);

                //if slot type is null. it shold not be dc30, uvi80 and up1600,igore
                if (!string.IsNullOrEmpty(l_strSlotType))
                {
                    List<string> l_lstSlotNumber = new List<string>();
                    foreach (string l_strSite in l_Row.Sites)
                    {
                        int l_intSiteNumber = getSiteNumber(l_strSite);
                        if (l_intSiteNumber!=-1)
                        {
                            if (!p_Dic.ContainsKey(l_intSiteNumber))
                            {
                                p_Dic.Add(l_intSiteNumber, l_strSlotType);
                            }
                            else if (!p_Dic[l_intSiteNumber].Equals(l_strSlotType))
                            {
                                Response.Report("Slot Number: " + l_intSiteNumber + " should not be defined with two different Pin Type[" +
                                    p_Dic[l_intSiteNumber] + "|" + l_strSlotType + "]", MessageLevel.Error, 0);
                                l_Rtn = false;
                            }
                            else
                            {
                                //do nothing
                            }
                        }
                        else
                        {
                            //do nothing--->no slot number
                        }
                    }
                }
                else
                {
                    //do nothing
                }
            }

            return l_Rtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_strSite"></param>
        /// <returns></returns>
        private int getSiteNumber(string p_strSite)
        {
            int l_Rtn = -1;

            int l_Result = 0;


            string[] l_ArySite = p_strSite.Split(new string[] { "." },StringSplitOptions.RemoveEmptyEntries);

            if (l_ArySite.Length > 0)
            {
                if (int.TryParse(l_ArySite[0], out l_Result))
                {
                    l_Rtn = l_Result;
                    return l_Rtn;
                }
                else
                {
                        //do nothing
                }
            }
            else
            {
                    //do nothing
            }

            return l_Rtn;
        }
    }

}