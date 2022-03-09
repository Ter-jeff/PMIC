using PmicAutomation.MyControls;
using PmicAutomation.Utility.Relay.Base;
using PmicAutomation.Utility.Relay.Input;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.Relay.Function
{
    public class SearchRelay
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly List<NodeName> _nodeNames = new List<NodeName>();
        private List<Tuple<Node, Node>> _paths = new List<Tuple<Node, Node>>();
        private List<List<Node>> paths = new List<List<Node>>();
        //private List<Node> onePath = new List<Node>();
        private readonly Queue<Node> _queueNodes = new Queue<Node>();
        private readonly List<RelayPathRecord> _relayPathRecord = new List<RelayPathRecord>();
        private readonly Stack<Node> _stackNodes = new Stack<Node>();

        private int _currentQueueIndex;
        private Node[] _nodes;

        private int _sourceIndex;

        public SearchRelay(MyForm.RichTextBoxAppend appendText)
        {
            _appendText = appendText;
        }

        private void Initialize()
        {
            _sourceIndex = 0;
            _relayPathRecord.Clear();
            _paths.Clear();
            paths.Clear();
            _stackNodes.Clear();
            _queueNodes.Clear();
            foreach (var node in _nodes)
            {
                node.SourceIndex = -1;
                node.Index = -1;
            }
        }

        private Node GetCurrentQueueNode()
        {
            if (_currentQueueIndex >= _queueNodes.Peek().RelationNodes.Count)
            {
                return null;
            }
            return _queueNodes.Peek().GetRelationNodes(_currentQueueIndex++);
        }

        private void QueueNodesDequeue()
        {
            _currentQueueIndex = 0;
            _queueNodes.Dequeue();
        }

        private bool IsNodeInStack(Node node)
        {
            if (!string.IsNullOrEmpty(node.NodeName.NetName) && _stackNodes.Any(x =>
                    x.NodeName.NetName.Equals(node.NodeName.NetName, StringComparison.CurrentCultureIgnoreCase)))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                _stackNodes.Any(x =>
                    x.NodeName.Refdes.Equals(node.NodeName.Refdes, StringComparison.CurrentCultureIgnoreCase)) &&
                _stackNodes.Any(x =>
                    x.NodeName.PinNumber == node.NodeName.PinNumber))
            {
                return true;
            }

            int matchPinNumber = GetMatchPinNumber(node);
            if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                matchPinNumber != -1 &&
                _stackNodes.Any(x =>
                    x.NodeName.Refdes.Equals(node.NodeName.Refdes, StringComparison.CurrentCultureIgnoreCase)) &&
                _stackNodes.Any(x =>
                    x.NodeName.PinNumber == matchPinNumber))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                _stackNodes.Any(x =>
                    x.NodeName.Refdes.Equals(node.NodeName.Refdes, StringComparison.CurrentCultureIgnoreCase)))
            {
                return true;
            }

            return false;
        }

        private int GetMatchPinNumber(Node node)
        {
            return GetMatchPinNumber(node.NodeName.PinNumber);
        }

        private int GetMatchPinNumber(int pinNumber)
        {
            if (pinNumber % 2 == 0)
                return pinNumber - 1;

            return pinNumber + 1;
        }

        private bool IsNodeInQueue(Node node)
        {
            if (!string.IsNullOrEmpty(node.NodeName.NetName) && _queueNodes.Any(x =>
                    x.NodeName.NetName.Equals(node.NodeName.NetName, StringComparison.CurrentCultureIgnoreCase)))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                _queueNodes.Any(x =>
                    x.NodeName.Refdes.Equals(node.NodeName.Refdes, StringComparison.CurrentCultureIgnoreCase)) &&
                _queueNodes.Any(x =>
                    x.NodeName.PinNumber == node.NodeName.PinNumber))
            {
                return true;
            }

            int matchPinNumber = GetMatchPinNumber(node);
            if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                matchPinNumber != -1 &&
                _queueNodes.Any(x =>
                    x.NodeName.Refdes.Equals(node.NodeName.Refdes, StringComparison.CurrentCultureIgnoreCase)) &&
                _queueNodes.Any(x =>
                    x.NodeName.PinNumber == matchPinNumber))
            {
                return true;
            }

            return false;
        }

        private bool GetPaths(Node currentNode, Node parentNode, Node startNode, Node endNode)
        {
            if (currentNode != null && parentNode != null && currentNode == parentNode)
            {
                return false;
            }

            if (currentNode != null)
            {
                int i = 0;
                _stackNodes.Push(currentNode);
                if (currentNode == endNode)
                {
                    return true;
                }

                Node nextNode = currentNode.GetRelationNodes(i);
                while (nextNode != null)
                {
                    if (parentNode != null &&
                        (nextNode == startNode || nextNode == parentNode || IsNodeInStack(nextNode)))
                    {
                        i++;
                        nextNode = i >= currentNode.RelationNodes.Count ? null : currentNode.GetRelationNodes(i);
                        continue;
                    }

                    if (GetPaths(nextNode, currentNode, startNode, endNode))
                    {
                        _stackNodes.Pop();
                    }

                    i++;
                    nextNode = i >= currentNode.RelationNodes.Count ? null : currentNode.GetRelationNodes(i);
                }

                _stackNodes.Pop();
                return false;
            }

            return false;
        }

        public List<List<Node>> GetAllBreadthFirstSearch(Node currentNode, Node parentNode, Node startNode, Node endNode)
        {
            Initialize();
            GetBreadthFirstSearch(currentNode, parentNode, startNode, endNode);
            if (endNode != null)
            {
                var current = endNode;
                List<Node> onePath = new List<Node>();
                GetPaths(onePath, current, startNode, endNode);
            }
            return paths;
        }


        public List<RelayPathRecord> GetRelayPathRecord(Node currentNode, Node parentNode, Node startNode, Node endNode)
        {
            Initialize();
            GetRelayPathRecordByBreadthFirstSearch(currentNode, parentNode, startNode, endNode);
            return _relayPathRecord;
        }

        private bool GetPaths(List<Node> onePath, Node currentNode, Node startNode, Node endNode)
        {
            if (endNode != null)
            {
                onePath.Add(currentNode);
                if (currentNode == startNode)
                    paths.Add(onePath);

                var nodes = _paths.Where(x => x.Item1 == currentNode).ToList();
                if (nodes.Any())
                {
                    foreach (var node in nodes)
                    {
                        List<Node> nodeList = new List<Node>();
                        foreach (var item in onePath)
                            nodeList.Add(item);
                        GetPaths(nodeList, node.Item2, startNode, endNode);
                    }
                    return true;
                }
            }
            return false;
        }

        private bool GetRelayPathRecordByBreadthFirstSearch(Node currentNode, Node parentNode, Node startNode, Node endNode)
        {
            //Set Index & sourceIndex
            currentNode.Index = _relayPathRecord.Count;
            if (_relayPathRecord.Count == 0) currentNode.SourceIndex = 0;
            foreach (var node in currentNode.RelationNodes)
            {
                if (node.SourceIndex == -1)
                    node.SourceIndex = currentNode.Index;
            }

            _relayPathRecord.Add(GetRelayPath(currentNode, parentNode));

            if (currentNode != null && parentNode != null && currentNode == parentNode)
            {
                return false;
            }

            if (currentNode != null)
            {
                _queueNodes.Enqueue(currentNode);
                _stackNodes.Push(currentNode);
                if (currentNode == endNode)
                {
                    return true;
                }

                Node nextNode = GetCurrentQueueNode();
                while (nextNode != null)
                {
                    if (nextNode == startNode || nextNode == parentNode || IsNodeInStack(nextNode))
                    {
                        nextNode = GetCurrentQueueNode();
                        continue;
                    }

                    if (GetRelayPathRecordByBreadthFirstSearch(nextNode, _queueNodes.Peek(), startNode, endNode))
                    {
                        QueueNodesDequeue();
                    }
                }
                QueueNodesDequeue();
                return false;
            }

            return false;
        }

        private bool GetBreadthFirstSearch(Node currentNode, Node parentNode, Node startNode, Node endNode)
        {
            foreach (var relationNode in currentNode.RelationNodes)
            {
                if (relationNode == startNode || relationNode == parentNode || IsNodeInStack(relationNode))
                    continue;
                _paths.Add(new Tuple<Node, Node>(relationNode, currentNode));
            }

            if (currentNode != null && parentNode != null && currentNode == parentNode)
            {
                return false;
            }

            if (currentNode != null)
            {
                _queueNodes.Enqueue(currentNode);
                _stackNodes.Push(currentNode);
                if (currentNode == endNode)
                {
                    return true;
                }

                Node nextNode = GetCurrentQueueNode();
                while (nextNode != null)
                {
                    if (nextNode == startNode || nextNode == parentNode || IsNodeInStack(nextNode))
                    {
                        nextNode = GetCurrentQueueNode();
                        continue;
                    }

                    if (GetBreadthFirstSearch(nextNode, _queueNodes.Peek(), startNode, endNode))
                    {
                        QueueNodesDequeue();
                    }
                }
                QueueNodesDequeue();
                return false;
            }

            return false;
        }

        private RelayPathRecord GetRelayPath(Node currentNode, Node parentNode)
        {
            string output = string.Join(",",
                currentNode.RelationNodes.Where(x => x != parentNode).Select(x => x.GetNodeName()));
            //.Where(x => x.StartsWith("S0", StringComparison.CurrentCultureIgnoreCase)));
            RelayPathRecord relayPathRecord = new RelayPathRecord
            {
                BOutputEmpty = string.IsNullOrEmpty(output) ? "True" : "False",
                Input = currentNode.GetOppositeNodeName(),
                InType = !string.IsNullOrEmpty(currentNode.GetName().NetName) ? "1" : "0",
                Output = output,
                OutType = !string.IsNullOrEmpty(currentNode.GetName().NetName) ? "0" : "1",
                SourceIndex = parentNode == null ? "0" : parentNode.Index.ToString()
            };
            return relayPathRecord;
        }

        private Node FindNode(string name)
        {
            int index = _nodeNames.FindIndex(x => x.NetName == name);
            return index == -1 ? null : _nodes[index];
        }

        private int FindNodeIndex(string name)
        {
            return _nodeNames.FindIndex(x => x.NetName == name);
        }

        public void SetNodes(List<ComPinRow> comPinRows, LinkedNodeRuleSheet linkedNodeRule)
        {
            Dictionary<string, string> linkedNodeDictionary = new Dictionary<string, string>();
            foreach (LinkedNodeRuleRow row in linkedNodeRule.Rows)
            {
                if (!linkedNodeDictionary.ContainsKey(row.Node))
                    linkedNodeDictionary.Add(row.Node, row.LinkedNode);
            }

            List<IGrouping<string, ComPinRow>> netNameGroup = comPinRows.GroupBy(x => x.NetName).Distinct().ToList();
            foreach (IGrouping<string, ComPinRow> row in netNameGroup)
            {
                NodeName nodeName = new NodeName { NetName = row.First().NetName };
                _nodeNames.Add(nodeName);
            }

            List<IGrouping<string, ComPinRow>> refdesGroup =
                comPinRows.GroupBy(x => x.Refdes + x.PinNumber).Distinct().ToList();
            foreach (IGrouping<string, ComPinRow> row in refdesGroup)
            {
                NodeName nodeName = new NodeName { Refdes = row.First().Refdes, PinNumber = row.First().PinNumber, PinName = row.First().PinName };
                _nodeNames.Add(nodeName);
            }

            List<List<int>> nodeRelation = new List<List<int>>();
            foreach (NodeName nodeName in _nodeNames)
            {
                if (nodeName.NetName == "S0_VBAT_UVI80_F")
                {
                }

                if (nodeName.Refdes == "S0_C6804" && nodeName.PinNumber == 1)
                {
                }

                List<int> list = new List<int>();
                if (!string.IsNullOrEmpty(nodeName.NetName))
                {
                    //NetName => REFDES
                    List<ComPinRow> netNameRows = comPinRows
                        .Where(x => x.NetName.Equals(nodeName.NetName, StringComparison.OrdinalIgnoreCase))
                            .Where(x => x.Refdes.StartsWith("S0", StringComparison.CurrentCultureIgnoreCase)).ToList();
                    foreach (ComPinRow netNameRow in netNameRows)
                    {
                        list.Add(_nodeNames.FindIndex(x =>
                            x.Refdes == netNameRow.Refdes && x.PinNumber == netNameRow.PinNumber));
                    }
                }

                if (!string.IsNullOrEmpty(nodeName.Refdes))
                {
                    //REFDES => NetName
                    List<ComPinRow> refdesRows1 = comPinRows.Where(x =>
                        x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                        x.PinNumber == nodeName.PinNumber).ToList();
                    foreach (ComPinRow refdesRow in refdesRows1)
                    {
                        list.Add(FindNodeIndex(refdesRow.NetName));
                    }

                    //By linkedNodeDictionary
                    var PinName = nodeName.PinName;
                    if (linkedNodeDictionary.ContainsKey(PinName) && comPinRows.Exists(x =>
                        x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                        x.PinName == linkedNodeDictionary[PinName]))
                    {

                        List<ComPinRow> refdesRows2 = comPinRows.Where(x =>
                            x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                             x.PinName == linkedNodeDictionary[PinName]).ToList();
                        foreach (ComPinRow refdesRow in refdesRows2)
                        {
                            list.Add(FindNodeIndex(refdesRow.NetName));
                        }
                    }
                    else
                    {
                        // pair for 1,2 & 3,4
                        var matchPinNumber = GetMatchPinNumber(nodeName.PinNumber);
                        List<ComPinRow> refdesRows2 = comPinRows.Where(x =>
                            x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                            x.PinNumber == matchPinNumber).ToList();
                        foreach (ComPinRow refdesRow in refdesRows2)
                        {
                            list.Add(FindNodeIndex(refdesRow.NetName));
                        }
                    }
                }

                nodeRelation.Add(list.Distinct().ToList());
            }

            _nodes = new Node[nodeRelation.Count];
            for (int i = 0; i < nodeRelation.Count; i++)
            {
                _nodes[i] = new Node();
                _nodes[i].SetName(_nodeNames[i]);
            }

            for (int i = 0; i < nodeRelation.Count; i++)
            {
                List<Node> list = new List<Node>();
                for (int j = 0; j < nodeRelation[i].Count; j++)
                {
                    list.Add(_nodes[nodeRelation[i][j]]);
                }
                _nodes[i].SetRelationNodes(list);
            }

            //Set single end node
            foreach (var node in _nodes)
            {
                if (node.NodeName.NetName.Equals("GND", StringComparison.CurrentCultureIgnoreCase))
                    node.RelationNodes.Clear();
                //if (!string.IsNullOrEmpty(node.NodeName.Refdes) &&
                //    !node.NodeName.Refdes.StartsWith("S0", StringComparison.CurrentCultureIgnoreCase))
                //    node.RelationNodes.Clear();
            }

        }

        public void SetNodesOld(List<ComPinRow> comPinRows, LinkedNodeRuleSheet linkedNodeRule)
        {
            Dictionary<string, string> linkedNodeDictionary = new Dictionary<string, string>();
            foreach (LinkedNodeRuleRow row in linkedNodeRule.Rows)
            {
                if (!linkedNodeDictionary.ContainsKey(row.Node))
                {
                    linkedNodeDictionary.Add(row.Node, row.LinkedNode);
                }

                if (!linkedNodeDictionary.ContainsKey(row.LinkedNode))
                {
                    linkedNodeDictionary.Add(row.LinkedNode, row.Node);
                }
            }

            List<IGrouping<string, ComPinRow>> netNameGroup = comPinRows.GroupBy(x => x.NetName).Distinct().ToList();
            foreach (IGrouping<string, ComPinRow> row in netNameGroup)
            {
                NodeName nodeName = new NodeName { NetName = row.First().NetName };
                _nodeNames.Add(nodeName);
            }

            List<IGrouping<string, ComPinRow>> refdesGroup =
                comPinRows.GroupBy(x => x.Refdes + x.PinNumber).Distinct().ToList();
            foreach (IGrouping<string, ComPinRow> row in refdesGroup)
            {
                NodeName nodeName = new NodeName { Refdes = row.First().Refdes, PinNumber = row.First().PinNumber };
                _nodeNames.Add(nodeName);
            }

            List<List<int>> nodeRelation = new List<List<int>>();
            foreach (NodeName nodeName in _nodeNames)
            {
                if (nodeName.NetName == "S0_VBAT_UVI80_F")
                {
                }

                //NetName => REFDES
                List<int> list = new List<int>();
                List<ComPinRow> netNameRows = comPinRows
                    .Where(x => x.NetName.Equals(nodeName.NetName, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (ComPinRow netNameRow in netNameRows)
                {
                    if (linkedNodeRule.Rows.Exists(x =>
                        x.Node.Equals(netNameRow.PinName, StringComparison.OrdinalIgnoreCase)))
                    {
                        list.Add(_nodeNames.FindIndex(x =>
                            x.Refdes == netNameRow.Refdes && x.PinNumber == netNameRow.PinNumber));
                    }
                }

                //REFDES => NetName
                List<ComPinRow> refdesRows = comPinRows.Where(x =>
                    x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                    x.PinNumber == nodeName.PinNumber).ToList();
                foreach (ComPinRow refdesRow in refdesRows)
                {
                    if (linkedNodeRule.Rows.Exists(x =>
                        x.Node.Equals(refdesRow.PinName, StringComparison.OrdinalIgnoreCase)))
                    {
                        string pinName = linkedNodeRule.Rows
                            .Find(x => x.Node.Equals(refdesRow.PinName, StringComparison.OrdinalIgnoreCase)).LinkedNode;
                        List<ComPinRow> meetRuleRows = comPinRows.Where(x =>
                            x.Refdes.Equals(nodeName.Refdes, StringComparison.CurrentCulture) &&
                            x.PinName.Equals(pinName, StringComparison.CurrentCulture)).ToList();
                        foreach (ComPinRow row in meetRuleRows)
                        {
                            list.Add(FindNodeIndex(row.NetName));
                        }
                    }
                }

                nodeRelation.Add(list.Distinct().ToList());
            }

            _nodes = new Node[nodeRelation.Count];
            for (int i = 0; i < nodeRelation.Count; i++)
            {
                _nodes[i] = new Node();
                _nodes[i].SetName(_nodeNames[i]);
            }

            for (int i = 0; i < nodeRelation.Count; i++)
            {
                List<Node> list = new List<Node>();
                for (int j = 0; j < nodeRelation[i].Count; j++)
                {
                    list.Add(_nodes[nodeRelation[i][j]]);
                }
                _nodes[i].SetRelationNodes(list);
            }
        }
        public List<RelayItem> GenRelayList(Dictionary<string, List<string>> filterPins, List<AdgMatrix> adgMatrixList)
        {
            if (!filterPins.ContainsKey("Resource Pin"))
            {
                return null;
            }

            if (!filterPins.ContainsKey("Device Pin"))
            {
                return null;
            }

            List<string> resourcePins = filterPins["Resource Pin"].Distinct().ToList();
            List<string> devicePins = filterPins["Device Pin"].Distinct().ToList();

            List<RelayItem> relayItems = new List<RelayItem>();
            int cnt = 0;
            foreach (string resourcePin in resourcePins)
            {
                cnt++;
                _appendText.Invoke("Starting to analyze pin " + resourcePin + " " + cnt + "/" + resourcePins.Count + " !!!", Color.Blue);
                foreach (string devicePin in devicePins)
                {
                    if (resourcePin.Equals(devicePin, StringComparison.CurrentCultureIgnoreCase))
                        continue;

                    if (resourcePin == "S0TO7_MAIN_BUFFER1_OUTPUT" && devicePin == "S0_CLK32K_UP1600")
                    {
                    }

                    Node startIndex = FindNode(resourcePin);
                    Node endIndex = FindNode(devicePin);
                    if (startIndex != null)
                    {
                        var paths = GetAllBreadthFirstSearch(startIndex, null, startIndex, endIndex);

                        if (paths.Count > 0)
                        {
                            List<string> relayPaths = new List<string>();
                            List<string> adgs = new List<string>();
                            List<string> recordPaths = new List<string>();
                            for (var i = 0; i < paths.Count(); i++)
                            {
                                var path = paths[i];
                                List<string> relays = path.Select(x => x.NodeName.Refdes)
                                    .Where(x => !string.IsNullOrEmpty(x)).Where(x => x.Length > 4)
                                    .Where(x => x.Substring(2, 2).Equals("_K", StringComparison.OrdinalIgnoreCase))
                                    .Select(x => RelayItem.GetPinName(x)).Reverse().ToList();

                                relayPaths.Add(string.Join(",", relays));

                                foreach (var node in path)
                                {
                                    if (adgMatrixList.Any(x => x.Name.Equals(node.NodeName.Refdes,
                                        StringComparison.CurrentCultureIgnoreCase)))
                                    {
                                        var adgMatrix = adgMatrixList.Find(x => x.Name.Equals(node.NodeName.Refdes,
                                            StringComparison.CurrentCultureIgnoreCase));
                                        if (adgMatrix.ResourcePins.Any(x =>
                                            x.NetName.Equals(resourcePin, StringComparison.CurrentCultureIgnoreCase)))
                                        {
                                            var pinName = adgMatrix.ResourcePins.Find(x =>
                                                x.NetName.Equals(resourcePin,
                                                    StringComparison.CurrentCultureIgnoreCase)).PinName;
                                            if (node.NodeName.Refdes.Contains("_"))
                                                adgs.Add(node.NodeName.Refdes.Substring(node.NodeName.Refdes.IndexOf("_") + 1) + "_" + pinName);
                                            else
                                                adgs.Add(node.NodeName.Refdes + "_" + pinName);
                                        }
                                    }
                                }

                                recordPaths.Add(string.Join(",",
                                    path.Select(
                                        x => x.NodeName.NetName + x.NodeName.Refdes + "_" + x.NodeName.PinNumber)));
                            }

                            if (relayPaths.Count() != 0)
                            {
                                RelayItem relay = new RelayItem
                                {
                                    ResourcePin = resourcePin,
                                    DevicePin = devicePin,
                                    Paths = paths.Select(x => x.Select(y => y.NodeName.GetName()).ToList()).ToList(),
                                    Relays = relayPaths,
                                    Adgs = adgs
                                };
                                relayItems.Add(relay);
                            }
                        }
                    }
                }
            }

            return relayItems;
        }

        public Dictionary<string, List<RelayPathRecord>> GenPinToPinFiles(Dictionary<string, List<string>> filterPins)
        {
            Dictionary<string, List<RelayPathRecord>> relayPathsDic = new Dictionary<string, List<RelayPathRecord>>();
            List<string> resourcePins = filterPins.ContainsKey("Resource Pin") ? filterPins["Resource Pin"] : null;
            if (resourcePins != null)
            {
                foreach (string resourcePin in resourcePins)
                {
                    string pinName = resourcePin;
                    Node startIndex = FindNode(pinName);
                    if (startIndex != null)
                    {
                        var relayPathRecord = GetRelayPathRecord(startIndex, null, startIndex, null).Select(x => x.DeepClone()).ToList();
                        relayPathsDic.Add(pinName, relayPathRecord);
                    }
                    else
                    {
                        _appendText.Invoke("The index of pin " + pinName + " can not be found !!!", Color.Red);
                    }
                }
            }

            return relayPathsDic;
        }
    }
}