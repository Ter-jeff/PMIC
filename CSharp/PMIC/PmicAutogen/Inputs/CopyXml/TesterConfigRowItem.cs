using System;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
    [Serializable]
    public class TesterConfigRowItem
    {
        [XmlAttribute] public string Id { get; set; }

        [XmlAttribute] public string Column3 { get; set; }

        [XmlAttribute] public string Column5 { get; set; }
    }
}