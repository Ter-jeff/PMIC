using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
	[Serializable]
	public class TesterConfigPinType
	{
		[XmlAttribute]
		public string Value { get; set; }

		[XmlAttribute]
		public string TypeValue { get; set; }

        [XmlAttribute]
        public string NeedCheck { get; set; }

        [XmlElementAttribute("TesterConfigRowItem")]
		public List<TesterConfigRowItem> lstTesterConfigRowItems { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string l_strRtn = string.Empty;

            foreach (TesterConfigRowItem l_RowItem in lstTesterConfigRowItems)
            {
                l_strRtn = l_strRtn + "\t\t" + l_RowItem.Column3 + "\t\t" + l_RowItem.Column5 + "\r\n";
            }

            return l_strRtn;
        }
    }
}
