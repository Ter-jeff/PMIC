using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
    [Serializable]
    public class TesterConfigTypeItem
    {
        [XmlElement("TesterConfigPinType")] public List<TesterConfigPinType> LstTesterConfigPinType { get; set; }

        /// <summary>
        /// </summary>
        /// <param name="p_Row"></param>
        /// <returns></returns>
        public bool IsValidPinNameType(ChannelMapRow pRow)
        {
            var lRtn = false;
            var lHasMatched = false;

            foreach (var lType in LstTesterConfigPinType)
            {
                if (lType.NeedCheck.Equals("0")) continue;

                if (lType.TypeValue.Equals(pRow.Type, StringComparison.CurrentCultureIgnoreCase))
                {
                    lHasMatched = true;
                    if (Regex.IsMatch(pRow.DeviceUnderTestPinName, lType.Value))
                    {
                        lRtn = true;
                        return lRtn;
                    }
                }
            }

            lRtn = lHasMatched == false ? true : false;

            return lRtn;
        }

        /// <summary>
        /// </summary>
        /// <param name="p_Row"></param>
        /// <returns></returns>
        public string GetTesterConfigTypeByPinAndPinType(ChannelMapRow pRow)
        {
            var lStrRtn = string.Empty;

            foreach (var lType in LstTesterConfigPinType)
                if (lType.TypeValue.Equals(pRow.Type, StringComparison.CurrentCultureIgnoreCase))
                    if (lType.NeedCheck.Equals("0") ||
                        Regex.IsMatch(pRow.DeviceUnderTestPinName, lType.Value))
                    {
                        lStrRtn = lType.Value;
                        return lStrRtn;
                    }


            return lStrRtn;
        }

        /// <summary>
        /// </summary>
        /// <param name="p_strType"></param>
        /// <returns></returns>
        public string GetContentByPinType(string pStrType)
        {
            var lStrRtn = string.Empty;

            foreach (var lType in LstTesterConfigPinType)
                if (lType.Value.Equals(pStrType, StringComparison.CurrentCultureIgnoreCase))
                {
                    lStrRtn = lType.ToString();
                    return lStrRtn;
                }


            return lStrRtn;
        }
    }
}