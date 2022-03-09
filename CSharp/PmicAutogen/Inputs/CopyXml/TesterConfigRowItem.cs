using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PmicAutogen.Inputs.CopyXml
{
	[Serializable]
	public class TesterConfigRowItem
	{
		[XmlAttribute]
		public string ID { get; set; }

		[XmlAttribute]
		public string Column3 { get; set; }

		[XmlAttribute]
		public string Column5 { get; set; }
	}
}
