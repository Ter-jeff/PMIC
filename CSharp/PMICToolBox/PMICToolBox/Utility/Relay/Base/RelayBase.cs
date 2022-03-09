using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.Relay.Base
{
    public class RelayItem
    {
        public string DevicePin;
        public string ResourcePin;
        public List<List<string>> Paths;
        public List<string> Relays;
        public List<string> Adgs;
        private const string Delimiter = "_To_";

        public List<string> GetNames()
        {
            List<string> names = new List<string>();
            if (Relays.Distinct().Count() == 1)
            {
                names.Add(GetResourcePin() + Delimiter + GetDevicePin());
                return names;
            }

            List<int> diff = new List<int>();
            for (var i = 0; i < Paths[0].Count; i++)
            {
                string context1 = Paths[0][i];
                for (var j = 1; j < Paths.Count; j++)
                {
                    string context2 = Paths[j][i];
                    if (context1 != context2)
                    {
                        if (Regex.IsMatch(context1, @"S\d_R", RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(context1, @"S\d_K", RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(context2, @"S\d_R", RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(context2, @"S\d_K", RegexOptions.IgnoreCase))
                        {
                            diff.Add(i);
                            break;
                        }
                    }
                }
            }

            for (var j = 0; j < Paths.Count; j++)
            {
                List<string> texts = new List<string>();
                for (int i = 0; i < Paths[j].Count; i++)
                {
                    if (diff.Contains(i))
                        texts.Add(Paths[j][i]);
                }
                texts.Reverse();
                names.Add(string.Join(Delimiter, texts));
            }

            return names;
        }

        public string GetDevicePin()
        {
            return GetPinName(DevicePin);
        }

        public string GetResourcePin()
        {
            return GetPinName(ResourcePin);
        }


        public static string GetPinName(string PinName)
        {
            if (Regex.IsMatch(PinName, @"^S\d_", RegexOptions.IgnoreCase))
                return PinName.Substring(3);
            return PinName;
        }
    }

    public class NodeName
    {
        public string NetName = "";
        public string PinName = "";
        public int PinNumber;
        public string Refdes = "";

        public int OppositePinName
        {
            get
            {
                if (PinNumber % 2 == 0)
                    return PinNumber - 1;
                return PinNumber + 1;
            }
        }

        public string GetName()
        {
            if (!string.IsNullOrEmpty(NetName)) return NetName;

            return Refdes + "_" + PinNumber;
        }
    }

    [Serializable]
    public class RelayPathRecord
    {
        public string BOutputEmpty = "";
        public string Input = "";
        public string InType = "";
        public string Output = "";
        public string OutType = "";
        public string SourceIndex = "";

        public RelayPathRecord DeepClone()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, this);
                stream.Seek(0, SeekOrigin.Begin);
                RelayPathRecord clonedSource = (RelayPathRecord)formatter.Deserialize(stream);
                return clonedSource;
            }
        }
    }

    public class Node
    {
        public int Index = -1;
        public int SourceIndex = -1;
        public NodeName NodeName = new NodeName();
        public List<Node> RelationNodes = new List<Node>();

        public NodeName GetName()
        {
            return NodeName;
        }

        public string GetNodeName()
        {
            return string.IsNullOrEmpty(NodeName.NetName)
                ? NodeName.Refdes + "_" + NodeName.PinNumber
                : NodeName.NetName;
        }

        public string GetOppositeNodeName()
        {
            return string.IsNullOrEmpty(NodeName.NetName)
                ? NodeName.Refdes + "_" + NodeName.OppositePinName
                : NodeName.NetName;
        }

        public void SetName(NodeName nodeName)
        {
            NodeName = nodeName;
        }

        public Node GetRelationNodes(int i)
        {
            if (i < RelationNodes.Count)
            {
                return RelationNodes[i];
            }

            return null;
        }

        public void SetRelationNodes(List<Node> nodes)
        {
            RelationNodes = nodes;
        }
    }
}
