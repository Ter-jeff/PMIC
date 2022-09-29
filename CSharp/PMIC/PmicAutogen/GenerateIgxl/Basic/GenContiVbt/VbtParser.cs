using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace PmicAutogen.GenerateIgxl.Basic.GenContiVbt
{
    public class VbtParser
    {
        private readonly string _file;
        private string _functionName = "";
        private int _left;
        private int _totalCnt;

        public VbtParser(string file)
        {
            _file = file;
            _left = 0;
        }

        #region Gen table

        public void GenTable(ExcelWorksheet sheet, List<Comment> comment)
        {
            var xList = comment.Select(x => x.X).Distinct().ToList();
            var yList = comment.Select(y => y.Y).Distinct().ToList();
            var xCnt = xList.Count;
            var yCnt = yList.Count;
            var arr = new object[yCnt, xCnt];
            sheet.Cells[1, 2].PrintExcelCol(xList.ToArray());
            sheet.Cells[2, 1].PrintExcelRow(yList.ToArray());
            foreach (var row in comment)
            {
                var xIndex = xList.IndexOf(row.X);
                var yIndex = yList.IndexOf(row.Y);
                arr[yIndex, xIndex] = row.Value;
            }

            sheet.Cells[2, 2].PrintExcelRange(arr);
        }

        #endregion

        #region Gen Vbt

        public void GenVbt(List<Dictionary<string, string>> dic, string functionName, string outputFileName)
        {
            _functionName = functionName;
            _totalCnt = dic.Count;
            var module = Path.GetFileNameWithoutExtension(outputFileName);
            //string output = Path.Combine(Path.GetDirectoryName(_file), module) + ".Bas";
            if (outputFileName != null)
                using (var sw = new StreamWriter(outputFileName, true))
                {
                    AddModuleName(module, sw);
                    using (var sr = new StreamReader(_file))
                    {
                        while (!sr.EndOfStream)
                        {
                            var line = sr.ReadLine();
                            if (line != null && line.Contains("<"))
                            {
                                var first = line.IndexOf("<", StringComparison.Ordinal);
                                var last = line.LastIndexOf(">", StringComparison.Ordinal);
                                var flag = false;
                                var isEnd = false;
                                var lastString = "";
                                if (first > 0 && last != -1)
                                {
                                    if (last + 1 != line.Length)
                                    {
                                        flag = true;
                                        lastString = line.Substring(last + 1);
                                    }

                                    if (last == line.Length - 1)
                                        isEnd = true;
                                    sw.Write(line.Substring(0, first));
                                    line = line.Substring(first, last - first + 1);
                                }

                                var matches = Regex.Matches(line, @"<.*?>");
                                var tag = matches[0].Value.Split(' ')[0].Replace(@"<", @"</");
                                var xml = GetXml(line, tag, sr);
                                var doc = new XmlDocument();
                                doc.LoadXml(xml);
                                XmlNode newNode = doc.DocumentElement;
                                SearchNode(newNode, dic, sw, 0, isEnd);
                                if (flag)
                                    sw.WriteLine(lastString);
                                //else
                                //    sw.Write(Environment.NewLine);
                            }
                            else
                            {
                                sw.WriteLine(line);
                                _left = 0;
                            }
                        }
                    }
                }
        }

        private string GetXml(string line, string tag, StreamReader sr)
        {
            var flag = false;
            do
            {
                if (line.Contains(tag))
                    flag = true;
                else
                    line += sr.ReadLine() + Environment.NewLine;
            } while (!flag);

            return line;
        }

        private int WriteStream(StreamWriter sw, string content, bool isEnd)
        {
            if (isEnd)
                sw.WriteLine(content);
            else
                sw.Write(content);
            return content.Length;
        }

        private void SearchNode(XmlNode nodes, List<Dictionary<string, string>> dic, StreamWriter sw, int currentCnt,
            bool isEnd)
        {
            if (dic.ElementAt(currentCnt).ContainsKey(nodes.Name))
            {
                var content = string.IsNullOrEmpty(nodes.InnerText)
                    ? dic.ElementAt(currentCnt)[nodes.Name]
                    : dic.ElementAt(currentCnt)[nodes.Name] + "|" + nodes.InnerText;
                if (nodes.Attributes != null)
                    foreach (XmlAttribute attribute in nodes.Attributes)
                    {
                        if (attribute.Name.Equals("Left", StringComparison.OrdinalIgnoreCase))
                        {
                            int leftValue;
                            int.TryParse(attribute.Value, out leftValue);
                            var pad = leftValue - _left - 1;
                            _left += WriteStream(sw, new string(' ', pad > 0 ? pad : 0), isEnd);
                        }

                        if (attribute.Name.Equals("Comment", StringComparison.OrdinalIgnoreCase) &&
                            attribute.Value.Equals("Y", StringComparison.OrdinalIgnoreCase))
                            content = @"'<Comment>" + content + "</Comment>";
                    }

                _left += WriteStream(sw, content, isEnd);
                return;
            }

            if (nodes.Name.Equals("SheetName", StringComparison.OrdinalIgnoreCase))
            {
                _left += WriteStream(sw, _functionName, isEnd);
                return;
            }

            if (nodes.Name.Equals("Seq", StringComparison.OrdinalIgnoreCase))
            {
                int value;
                int.TryParse(nodes.InnerText, out value);
                _left += WriteStream(sw, (value + currentCnt).ToString(), isEnd);
                return;
            }

            if (nodes.Name.Equals("Count", StringComparison.OrdinalIgnoreCase))
            {
                _left += WriteStream(sw, _totalCnt.ToString(), isEnd);
                return;
            }

            if (nodes.Name.Equals("#text", StringComparison.OrdinalIgnoreCase))
            {
                _left += WriteStream(sw, nodes.InnerText, isEnd);
                if (nodes.InnerText.LastIndexOf(Environment.NewLine, StringComparison.Ordinal) != -1)
                    _left = nodes.InnerText.Length -
                            nodes.InnerText.IndexOf(Environment.NewLine, StringComparison.Ordinal) - 2;
                return;
            }

            if (nodes.Name.Equals("Loop", StringComparison.OrdinalIgnoreCase))
            {
                var flag = false;
                var groupByName = "";
                if (nodes.Attributes != null)
                    foreach (XmlAttribute attribute in nodes.Attributes)
                    {
                        if (attribute.Name.Equals("IsInsertNewLine", StringComparison.OrdinalIgnoreCase) &&
                            attribute.Value.Equals("Y", StringComparison.OrdinalIgnoreCase))
                            flag = true;

                        if (attribute.Name.Equals("GroupBy", StringComparison.OrdinalIgnoreCase))
                            groupByName = attribute.Value;
                    }

                if (string.IsNullOrEmpty(groupByName))
                {
                    currentCnt = 0;
                    for (var i = 0; i < _totalCnt; i++)
                    {
                        if (nodes.HasChildNodes)
                        {
                            foreach (XmlNode node in nodes) SearchNode(node, dic, sw, currentCnt, isEnd);
                            var xmlNode = nodes.ChildNodes.Item(nodes.ChildNodes.Count - 1);
                            if (xmlNode != null && xmlNode.NodeType == XmlNodeType.Element)
                                sw.Write(Environment.NewLine);
                        }

                        _left = 0;
                        currentCnt++;
                    }
                }
                else
                {
                    if (dic.ElementAt(currentCnt).ContainsKey(groupByName))
                    {
                        var categoryList = dic.Select(x =>
                                x.First(y => y.Key.Equals(groupByName, StringComparison.OrdinalIgnoreCase)).Value)
                            .Distinct();
                        foreach (var category in categoryList)
                        {
                            var newDic = dic.Where(row =>
                                row.ContainsKey(groupByName) &&
                                row[groupByName].Equals(category, StringComparison.OrdinalIgnoreCase)).ToList();
                            currentCnt = 0;
                            if (nodes.HasChildNodes)
                            {
                                foreach (XmlNode node in nodes)
                                    if (node.Name.Equals("GroupBy", StringComparison.OrdinalIgnoreCase))
                                        for (var i = 0; i < newDic.Count; i++)
                                        {
                                            if (node.HasChildNodes)
                                            {
                                                var childNodes = node.ChildNodes;
                                                foreach (XmlNode childNode in childNodes)
                                                    SearchNode(childNode, newDic, sw, currentCnt, isEnd);
                                                sw.Write(Environment.NewLine);
                                            }

                                            _left = 0;
                                            currentCnt++;
                                        }
                                    else
                                        SearchNode(node, dic, sw, currentCnt, isEnd);

                                var xmlNode1 = nodes.ChildNodes.Item(nodes.ChildNodes.Count - 1);
                                if (xmlNode1 != null && xmlNode1.NodeType == XmlNodeType.Element)
                                    sw.Write(Environment.NewLine);
                                _left = 0;
                            }
                        }
                    }
                }

                if (flag) sw.Write(Environment.NewLine);
            }
            else
            {
                if (nodes.HasChildNodes)
                    foreach (XmlNode item in nodes)
                        SearchNode(item, dic, sw, currentCnt, isEnd);
            }
        }

        private void AddModuleName(string module, StreamWriter sw)
        {
            sw.WriteLine("Attribute VB_Name = \"" + module + "\"");
        }

        #endregion
    }
}