using PmicAutomation.MyControls;
using PmicAutomation.Utility.VbtGenerator.Input;
using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace PmicAutomation.Utility.VbtGenerator.Function
{
    public class Comment
    {
        public string Value;
        public string X;
        public string Y;
    }

    public class BasParser
    {
        private readonly MyForm.RichTextBoxAppend _append;
        private readonly string _file;
        private string _functionName = "";
        private int _totalCnt;
        private int _distinctCnt;

        public BasParser(string file, MyForm.RichTextBoxAppend append)
        {
            _file = file;
            _append = append;
        }

        #region Gen Bas

        public List<string> GenBas(List<TableSheet> tableSheets)
        {
            List<string> lines = new List<string>();
            foreach (TableSheet tableSheet in tableSheets)
            {
                _functionName = tableSheet.Name;
                _totalCnt = tableSheet.Table.Count;
                _distinctCnt = tableSheet.Table[0].ContainsKey("MeasurePins") ? tableSheet.Table.Select(p => p["MeasurePins"]).Distinct().Count() : 0;
                string[] inputLines = File.ReadAllLines(_file);
                List<MyLine> myLines = inputLines.Select((t, i) => new MyLine { Line = t, LineNum = i + 1 }).ToList();

                int currentLineNum = 0;
                for (int i = 0; i < myLines.Count; i++)
                {
                    if (myLines[i].LineNum <= currentLineNum)
                    {
                        continue;
                    }

                    if (myLines[i].Line != null &&
                        Regex.IsMatch(myLines[i].Line, MyXml.StartSymbol + ".+" + MyXml.EndSymbol))
                    {
                        lines.Add("");
                        MyXml myXml = new MyXml(_append, myLines[i]); // starting node is "<#Loop Groupby=\"MeasurePins\"#>"
                        string before;
                        List<MyLine> xmlLines;
                        string after;
                        myXml.GetXml(myLines.GetRange(i + 1, myLines.Count - i - 1), out before,
                            out xmlLines, out after); // xmlLines will fetch full block of <#Loop#><#/Loop#>
                        currentLineNum = xmlLines.Last().LineNum;
                        MyNode myNode = myXml.LoadXml(xmlLines); // fetch InnerXml & OuterXml & ChildNodes( e, <#MeasurePins#><#/MeasurePins#>,=, <#SeqNum#><#/SeqNum#>)

                        if (!string.IsNullOrEmpty(before))
                        {
                            lines[lines.Count - 1] = lines.Last() + before;
                        }

                        SearchNode(myNode, tableSheet, lines, 0, 0);

                        if (!string.IsNullOrEmpty(after))
                        {
                            lines[lines.Count - 1] = lines.Last() + after;
                        }
                    }
                    else
                    {
                        lines.Add(myLines[i].Line);
                    }
                }
            }

            return lines;
        }

        private void WriteLine(List<string> lines, string content)
        {
            lines[lines.Count - 1] = lines.Last() + content;
        }

        private void SearchNode(MyNode myNode, TableSheet tableSheet, List<string> lines, int currentCnt, int seqCnt)
        {
            List<Dictionary<string, string>> dics = tableSheet.Table;
            if (dics.ElementAt(currentCnt).ContainsKey(myNode.Name))
            {
                string content = !myNode.InnerText.Any()
                    ? dics.ElementAt(currentCnt)[myNode.Name]
                    : dics.ElementAt(currentCnt)[myNode.Name] + "|" + string.Join(Environment.NewLine, myNode.InnerText.Select(x => x.Line));

                if (myNode.ExistAttribute("Left"))
                {
                    int leftValue;
                    int.TryParse(myNode.GetAttribute("Left"), out leftValue);
                    int start = lines.Last().LastIndexOf(Environment.NewLine, StringComparison.Ordinal);
                    int pad = start == -1
                        ? leftValue - lines.Last().Length - 1
                        : leftValue - (lines.Last().Length - start - 2) - 1;
                    WriteLine(lines, new string(' ', pad > 0 ? pad : 0));
                }

                if (myNode.ExistAttribute("Comment") &&
                    myNode.GetAttribute("Comment").Equals("TRUE", StringComparison.CurrentCultureIgnoreCase))
                {
                    content = @"'<Comment>" + content + "</Comment>";
                }

                if (myNode.ExistAttribute("Join"))
                {
                    List<string> contents = new List<string>();
                    foreach (Dictionary<string, string> dic in dics)
                    {
                        contents.AddRange(dic
                            .Where(x => x.Key.Equals(myNode.Name, StringComparison.CurrentCultureIgnoreCase))
                            .Select(x => x.Value));
                    }

                    content = string.Join(myNode.GetAttribute("Join"), contents);
                }

                if (myNode.ExistAttribute("ReplaceOld") && myNode.ExistAttribute("ReplaceNew"))
                {
                    content = content.Replace(myNode.GetAttribute("ReplaceOld"), myNode.GetAttribute("ReplaceNew"));
                }

                WriteLine(lines, content);
                return;
            }

            if (myNode.Name.Equals("LIST_MeasurePins", StringComparison.OrdinalIgnoreCase))
            {
                List<string> measPinList;
                if (myNode.ExistAttribute("Type"))
                {
                    string type = myNode.GetAttribute("Type");
                    measPinList = dics.Where(p => p["Instance"].Equals(type, StringComparison.CurrentCultureIgnoreCase))
                        .Select(q => q["MeasurePins"]).Distinct().ToList();
                }
                else
                {
                    measPinList = dics.Select(p => p["MeasurePins"]).Distinct().ToList();
                }
                WriteLine(lines, string.Join(",", measPinList));
                return;
            }

            if (myNode.Name.Equals("AllPinSetting", StringComparison.OrdinalIgnoreCase))
            {
                foreach (MyAttribute attribute in myNode.Attributes)
                {
                    if (tableSheet.AllPinSettingDic.ContainsKey(attribute.Value))
                    {
                        WriteLine(lines, tableSheet.AllPinSettingDic[attribute.Value]);
                    }
                }
                return;
            }

            if (myNode.Name.Equals("SheetName", StringComparison.OrdinalIgnoreCase))
            {
                WriteLine(lines, _functionName);
                return;
            }

            if (myNode.Name.Equals("SeqNum", StringComparison.OrdinalIgnoreCase))
            {
                int value = 0;
                if (myNode.ExistAttribute("Start"))
                {
                    int.TryParse(myNode.GetAttribute("Start"), out value);
                }

                WriteLine(lines, (value + seqCnt).ToString(CultureInfo.InvariantCulture));
                return;
            }

            if (myNode.Name.Equals("Count", StringComparison.OrdinalIgnoreCase))
            {   
                int value;
                int.TryParse(string.Join(Environment.NewLine, myNode.InnerXml.Select(x => x.Line)), out value);
                if (myNode.ExistAttribute("Type")  && myNode.GetAttribute("Type").Equals("Distinct"))
                    WriteLine(lines, (_distinctCnt + value).ToString(CultureInfo.InvariantCulture));
                else
                    WriteLine(lines, (_totalCnt + value).ToString(CultureInfo.InvariantCulture));
                return;
            }

            if (myNode.Name.Equals("#text", StringComparison.OrdinalIgnoreCase))
            {
                WriteLine(lines, string.Join(Environment.NewLine, myNode.InnerText.Select(x => x.Line)));
                return;
            }

            if (myNode.Name.Equals("Loop", StringComparison.OrdinalIgnoreCase))
            {
                if (!myNode.ExistAttribute("GroupBy"))
                {
                    currentCnt = 0;
                    seqCnt = 0;
                    for (int i = 0; i < _totalCnt; i++)
                    {
                        if (myNode.ExistAttribute("Type"))
                        {
                            string type = myNode.GetAttribute("Type");
                            if (!dics.ElementAt(currentCnt)["Instance"].Equals(type, StringComparison.OrdinalIgnoreCase))
                            {
                                ++currentCnt;
                                //++seqCnt;
                                continue;
                            }
                        }

                        if (myNode.ExistAttribute("IsInsertNewLine") &&
                            myNode.GetAttribute("IsInsertNewLine")
                                .Equals("TRUE", StringComparison.CurrentCultureIgnoreCase))
                        {
                            lines.Add("");
                        }

                        // split each row and using every token to do recursive call
                        if (myNode.HasChildNodes)
                        {
                            int currentLineNum = 0;
                            for (int j = 0; j < myNode.ChildNodes.Count; j++)
                            {
                                MyNode node = myNode.ChildNodes[j];
                                if (j != 0 && currentLineNum != myNode.ChildNodes[j].OuterXml.First().LineNum)
                                {
                                    lines.Add("");
                                }

                                currentLineNum = myNode.ChildNodes[j].OuterXml.Last().LineNum;
                                SearchNode(node, tableSheet, lines, currentCnt, seqCnt);
                            }
                        }

                        if (i != _totalCnt - 1)
                        {
                            lines.Add("");
                        }

                        currentCnt++;
                        seqCnt++;
                    }
                }
                else
                {
                    string groupByName = myNode.GetAttribute("GroupBy");
                    if (dics.ElementAt(currentCnt).ContainsKey(groupByName))
                    {
                        List<string> categories = dics.Select(x =>
                                x.First(y => y.Key.Equals(groupByName, StringComparison.OrdinalIgnoreCase)).Value)
                            .Distinct().ToList();
                        seqCnt = 0;
                        for (int index = 0; index < categories.Count; index++)
                        {
                            if (myNode.ExistAttribute("IsInsertNewLine") &&
                                myNode.GetAttribute("IsInsertNewLine")
                                    .Equals("TRUE", StringComparison.CurrentCultureIgnoreCase))
                            {
                                lines.Add("");
                            }

                            TableSheet newTableSheet = new TableSheet();
                            List<Dictionary<string, string>> newDic = dics.Where(row =>
                                    row.ContainsKey(groupByName) &&
                                    row[groupByName].Equals(categories[index], StringComparison.OrdinalIgnoreCase))
                                .ToList();
                            newTableSheet.Table = newDic;
                            if (myNode.HasChildNodes)
                            {
                                int currentLineNum = 0;
                                for (int i = 0; i < myNode.ChildNodes.Count; i++)
                                {
                                    if (i != 0 && currentLineNum != myNode.ChildNodes[i].OuterXml.First().LineNum)
                                    {
                                        lines.Add("");
                                    }

                                    currentLineNum = myNode.ChildNodes[i].OuterXml.Last().LineNum;
                                    MyNode node = myNode.ChildNodes[i];
                                    currentCnt = 0;
                                    if (node.Name.Equals("GroupBy", StringComparison.OrdinalIgnoreCase))
                                    {
                                        for (int j = 0; j < newDic.Count; j++)
                                        {
                                            SearchNode(node, newTableSheet, lines, currentCnt, seqCnt);
                                            if (j != newDic.Count - 1)
                                            {
                                                lines[lines.Count - 1] += Environment.NewLine;
                                            }

                                            currentCnt++;
                                        }
                                    }
                                    else
                                    {
                                        SearchNode(node, newTableSheet, lines, currentCnt, seqCnt);
                                    }
                                }
                            }

                            if (index != categories.Count - 1)
                            {
                                lines.Add("");
                            }

                            seqCnt++;
                        }
                    }
                }
            }
            else
            {
                if (myNode.HasChildNodes)
                {
                    foreach (MyNode item in myNode.ChildNodes)
                    {
                        SearchNode(item, tableSheet, lines, currentCnt, seqCnt);
                    }
                }
            }
        }

        #endregion

        #region Gen table

        public void GenTable(ExcelWorksheet workSheet, List<Comment> comments)
        {
            List<string> xList = comments.Select(x => x.X).Distinct().ToList();
            List<string> yList = comments.Select(y => y.Y).Distinct().ToList();
            int xCnt = xList.Count;
            int yCnt = yList.Count;
            object[,] arr = new object[yCnt, xCnt];
            workSheet.Cells[1, 2].PrintExcelRow(xList.ToArray());
            workSheet.Cells[2, 1].PrintExcelCol(yList.ToArray());
            foreach (Comment row in comments)
            {
                int xIndex = xList.IndexOf(row.X);
                int yIndex = yList.IndexOf(row.Y);
                arr[yIndex, xIndex] = row.Value;
            }

            workSheet.Cells[2, 2].PrintExcelRange(arr);
        }

        public List<Comment> GetComment(string file)
        {
            List<Comment> comments = new List<Comment>();
            using (StreamReader sr = new StreamReader(file))
            {
                const string regexPattern = "\"(?<value>.*)\"";
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (line == null)
                    {
                        continue;
                    }

                    if (line.Contains(@"'<Comment>"))
                    {
                        string before = line.Substring(0, line.IndexOf(@"'<Comment>", StringComparison.Ordinal));
                        string after = line.Substring(line.IndexOf(@"'<Comment>", StringComparison.Ordinal) + 1);
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(after);
                        XmlElement node = doc.DocumentElement;
                        if (node == null)
                        {
                            continue;
                        }

                        if (string.IsNullOrEmpty(node.InnerText))
                        {
                            continue;
                        }

                        if (node.Name.Equals("Comment", StringComparison.OrdinalIgnoreCase))
                        {
                            string[] arr = node.InnerText.Split('|');
                            if (arr.Length != 3)
                            {
                                continue;
                            }

                            Comment comment = new Comment { Y = arr[0], X = arr[1] };
                            int start = before.IndexOf(arr[2], StringComparison.Ordinal);
                            string value = start != -1 ? before.Substring(start + arr[2].Length) : "";
                            if (Regex.IsMatch(value, regexPattern, RegexOptions.IgnoreCase))
                            {
                                comment.Value = Regex.Match(value, regexPattern, RegexOptions.IgnoreCase)
                                    .Groups["value"].Value;
                            }

                            comments.Add(comment);
                        }
                    }
                }
            }

            return comments;
        }

        #endregion
    }
}