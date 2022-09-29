using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
    [Serializable]
    public class TesterConfigPinType
    {
        [XmlAttribute] public string Value { get; set; }

        [XmlAttribute] public string TypeValue { get; set; }

        [XmlAttribute] public string NeedCheck { get; set; }

        [XmlElement("TesterConfigRowItem")] public List<TesterConfigRowItem> LstTesterConfigRowItems { get; set; }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var lStrRtn = string.Empty;

            foreach (var lRowItem in LstTesterConfigRowItems)
                lStrRtn = lStrRtn + "\t\t" + lRowItem.Column3 + "\t\t" + lRowItem.Column5 + "\r\n";

            return lStrRtn;
        }
    }
}