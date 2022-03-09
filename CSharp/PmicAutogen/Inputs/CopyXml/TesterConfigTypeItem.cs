using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
	[Serializable]
	public class TesterConfigTypeItem
	{
		[XmlElementAttribute("TesterConfigPinType")]
		public List<TesterConfigPinType> lstTesterConfigPinType { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_Row"></param>
        /// <returns></returns>
        public bool IsValidPinNameType(ChannelMapRow p_Row)
        {
            bool l_Rtn = false;
            bool l_HasMatched = false;

            foreach (TesterConfigPinType l_Type in lstTesterConfigPinType)
            {
                if (l_Type.NeedCheck.Equals("0"))
                {
                    continue;
                }

                if (l_Type.TypeValue.Equals(p_Row.Type, StringComparison.CurrentCultureIgnoreCase))
                {
                    l_HasMatched = true;
                    if (Regex.IsMatch(p_Row.DeviceUnderTestPinName, l_Type.Value))
                    {
                        l_Rtn = true;
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

            }

            l_Rtn = l_HasMatched==false?true:false;

            return l_Rtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_Row"></param>
        /// <returns></returns>
        public string GetTesterConfigTypeByPinAndPinType(ChannelMapRow p_Row)
        {
            string l_strRtn = string.Empty;

            foreach (TesterConfigPinType l_Type in lstTesterConfigPinType)
            {
                if (l_Type.TypeValue.Equals(p_Row.Type, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (l_Type.NeedCheck.Equals("0")|| 
                        Regex.IsMatch(p_Row.DeviceUnderTestPinName, l_Type.Value))
                    {
                        l_strRtn = l_Type.Value;
                        return l_strRtn;
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

            }


            return l_strRtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_strType"></param>
        /// <returns></returns>
        public string GetContentByPinType(string p_strType)
        {
            string l_strRtn = string.Empty;

            foreach (TesterConfigPinType l_Type in lstTesterConfigPinType)
            {
                if (l_Type.Value.Equals(p_strType, StringComparison.CurrentCultureIgnoreCase))
                {
                    l_strRtn = l_Type.ToString();
                    return l_strRtn;
                }
                else
                {
                    //do nothing
                }

            }


            return l_strRtn;
        }
	}
}
