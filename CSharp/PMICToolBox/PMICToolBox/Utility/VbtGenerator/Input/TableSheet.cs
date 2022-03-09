using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace PmicAutomation.Utility.VbtGenerator.Input
{
    [Serializable]
    public class TableSheet
    {
        public Dictionary<string, string> AllPinSettingDic = new Dictionary<string, string>();
        public string Block;
        public string Name;
        public List<Dictionary<string, string>> Table = new List<Dictionary<string, string>>();

        public TableSheet DeepClone()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, this);
                stream.Seek(0, SeekOrigin.Begin);
                return (TableSheet)formatter.Deserialize(stream);
            }
        }
    }
}