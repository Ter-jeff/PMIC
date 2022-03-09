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

        private DataTable ReadPayloadType(XmlNode pNode)
        {
            var payloadTable = new DataTable();
            var positionNode = pNode.SelectSingleNode(KeyPosition);
            if (positionNode != null)
            {
                var innerTexts = positionNode.InnerText.Split(',');
                payloadTable.Columns.Add(Key);
                foreach (var innerText in innerTexts)
                    payloadTable.Columns.Add(innerText.Trim());

                var keyNode = pNode.SelectSingleNode(Key);
                if (keyNode != null)
                {
                    var nodeList = keyNode.ChildNodes;
                    foreach (XmlNode node in nodeList)
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