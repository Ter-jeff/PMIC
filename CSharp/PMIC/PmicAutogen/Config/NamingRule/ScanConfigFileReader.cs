using System;
using System.Data;
using System.IO;
using System.Xml;

namespace PmicAutogen.Config.NamingRule
{
    public class ScanConfigFileReader
    {
        private const string KeyPosition = "KeyPosition";
        private const string Key = "Key";

        public DataTable ReadConfig(Stream configFilePath)
        {
            try
            {
                var doc = new XmlDocument();
                var settings = new XmlReaderSettings();
                settings.IgnoreComments = true;
                var reader = XmlReader.Create(configFilePath, settings);
                doc.Load(reader);
                var configNode = doc.SelectSingleNode("ScanConfig");
                if (configNode != null && configNode.SelectSingleNode("PayloadType") != null)
                    return ReadPayloadType(configNode.SelectSingleNode("PayloadType"));
                return null;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message + "Error occurs when Reading Scan Config, please Check Config file!");
            }
        }

        private DataTable ReadPayloadType(XmlNode xmlNode)
        {
            var payloadTable = new DataTable();
            var positionNode = xmlNode.SelectSingleNode(KeyPosition);
            if (positionNode != null)
            {
                var innerTexts = positionNode.InnerText.Split(',');
                payloadTable.Columns.Add(Key);
                foreach (var innerText in innerTexts) payloadTable.Columns.Add(innerText.Trim());
                var keyNode = xmlNode.SelectSingleNode(Key);
                if (keyNode != null)
                {
                    var nodes = keyNode.ChildNodes;
                    foreach (XmlNode node in nodes)
                    {
                        var row = payloadTable.NewRow();
                        row[0] = node.Name;
                        var values = node.InnerText.Split(',');
                        for (var i = 0; i < values.Length; i++)
                            row[i + 1] = values[i].Trim();
                        payloadTable.Rows.Add(row);
                    }
                }
            }

            return payloadTable;
        }
    }
}