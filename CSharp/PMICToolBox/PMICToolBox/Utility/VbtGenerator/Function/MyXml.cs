using PmicAutomation.MyControls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.VbtGenerator.Function
{
    public enum MyNodeType { Element, Text }

    public class MyLine
    {
        public string Line;
        public int LineNum;
    }

    public class MyInner
    {
        public string Name;
        public MyNodeType NodeType;
        public List<MyLine> Text;
    }

    public class MyTag
    {
        public string After;
        public List<MyAttribute> Attributes = new List<MyAttribute>();
        public string Before;
        public string EndTag;
        public string FullStartTag;
        public List<MyLine> Inner;
        public string Name;
        public MyNodeType NodeType;
        public List<MyLine> Outer;
        public string StartTag;
    }

    public class MyAttribute
    {
        public string Name;
        public string Value;
    }

    public class MyNode
    {
        public readonly List<MyNode> ChildNodes = new List<MyNode>();
        public List<MyAttribute> Attributes = new List<MyAttribute>();
        public bool HasChildNodes;
        public List<MyLine> InnerText = new List<MyLine>();
        public List<MyLine> InnerXml;
        public string Name;
        public MyNodeType NodeType;
        public List<MyLine> OuterXml;
        public List<MyLine> Value;

        public bool ExistAttribute(string name)
        {
            foreach (MyAttribute attribute in Attributes)
            {
                if (attribute.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public string GetAttribute(string name)
        {
            foreach (MyAttribute attribute in Attributes)
            {
                if (attribute.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return attribute.Value;
                }
            }

            return "";
        }
    }

    public class MyXml
    {
        public const string StartSymbol = "<#";
        public const string EndSymbol = "#>";
        private readonly MyForm.RichTextBoxAppend _append;
        private readonly MyLine _myLine;

        public MyXml(MyForm.RichTextBoxAppend append, MyLine myLine)
        {
            _append = append;
            _myLine = myLine;
        }

        public MyNode LoadXml(List<MyLine> myLines)
        {
            if (!myLines.First().Line.StartsWith(StartSymbol, StringComparison.CurrentCulture) &&
                !myLines.Last().Line.EndsWith(EndSymbol, StringComparison.CurrentCulture))
            {
                AppendFormatError(myLines.First());
            }

            MyTag tag = GetTag(myLines);

            MyNode myNode = new MyNode
            {
                NodeType = MyNodeType.Element,
                Value = null,
                Name = tag.Name,
                Attributes = tag.Attributes,
                InnerText = new List<MyLine>(),
                InnerXml = tag.Inner,
                OuterXml = myLines,
                HasChildNodes = tag.Inner.Count != 0
            };

            GetNode(tag.Inner, myNode);
            return myNode;
        }


        private void GetNode(List<MyLine> myLines, MyNode myNode)
        {
            List<MyInner> inners = new List<MyInner>();
            SplitInner(myLines, inners);
            foreach (MyInner inner in inners)
            {
                myNode.HasChildNodes = true;
                if (inner.NodeType == MyNodeType.Text)
                {
                    MyNode childNode = new MyNode
                    {
                        NodeType = MyNodeType.Text,
                        Name = "#text",
                        Value = inner.Text,
                        InnerText = inner.Text,
                        InnerXml = new List<MyLine>(),
                        OuterXml = inner.Text,
                        HasChildNodes = false
                    };
                    myNode.ChildNodes.Add(childNode);
                }
                else
                {
                    MyTag tag = GetTag(inner.Text);
                    MyNode childNode = new MyNode
                    {
                        NodeType = MyNodeType.Element,
                        Value = null,
                        Name = tag.Name,
                        Attributes = tag.Attributes,
                        InnerText = tag.Inner,
                        InnerXml = tag.Inner,
                        OuterXml = tag.Outer
                    };
                    myNode.ChildNodes.Add(childNode);
                    GetNode(tag.Inner, childNode);
                }
            }
        }

        private void GetIndexof(List<MyLine> myLines, string search, out int lineNum, out int index, int startLine = -1,
            int start = -1)
        {
            foreach (MyLine myLine in myLines)
            {
                if (myLine.LineNum < startLine)
                {
                    continue;
                }

                int idx = start == -1
                    ? myLine.Line.IndexOf(search, StringComparison.OrdinalIgnoreCase)
                    : myLine.Line.IndexOf(search, start, StringComparison.OrdinalIgnoreCase);
                if (idx != -1)
                {
                    index = idx;
                    lineNum = myLine.LineNum;
                    return;
                }
            }

            index = -1;
            lineNum = -1;
        }

        private List<MyLine> GetSubString(List<MyLine> myLines, int startLine, int start, int endLine = -1,
            int end = -1)
        {
            List<MyLine> xmlLines = new List<MyLine>();
            foreach (MyLine myLine in myLines)
            {
                if (myLine.LineNum == startLine && myLine.LineNum == endLine)
                {
                    string text = myLine.Line.Substring(start, end - start);
                    if (!string.IsNullOrEmpty(text))
                    {
                        xmlLines.Add(new MyLine { Line = text, LineNum = myLine.LineNum });
                    }

                    break;
                }

                if (myLine.LineNum == startLine)
                {
                    string text = myLine.Line.Substring(start);
                    if (!string.IsNullOrEmpty(text))
                    {
                        xmlLines.Add(new MyLine { Line = text, LineNum = myLine.LineNum });
                    }
                }
                else if (myLine.LineNum == endLine)
                {
                    string text = myLine.Line.Substring(0, end);
                    if (!string.IsNullOrEmpty(text))
                    {
                        xmlLines.Add(new MyLine { Line = text, LineNum = myLine.LineNum });
                    }

                    break;
                }
                else if (startLine < myLine.LineNum && myLine.LineNum < endLine)
                {
                    xmlLines.Add(myLine);
                }
                else if (startLine < myLine.LineNum && endLine == -1)
                {
                    xmlLines.Add(myLine);
                }
            }

            return xmlLines;
        }

        public void GetXml(List<MyLine> myLines, out string before, out List<MyLine> xmlLines, out string after)
        {
            xmlLines = new List<MyLine>();
            string endTag = GetEndTag(_myLine.Line);
            int index1 = _myLine.Line.IndexOf(StartSymbol, StringComparison.OrdinalIgnoreCase);
            before = index1 == -1 ? "" : _myLine.Line.Substring(0, index1);
            xmlLines = index1 == -1
                ? new List<MyLine> { new MyLine { Line = _myLine.Line.Substring(index1), LineNum = _myLine.LineNum } }
                : new List<MyLine> { _myLine };
            after = "";
            foreach (MyLine myLine in myLines)
            {
                int lineNum2; int index2;
                GetIndexof(xmlLines, endTag, out lineNum2, out index2);
                if (index2 != -1)
                {
                    int lineNum3; int index3;
                    GetIndexof(xmlLines, EndSymbol, out lineNum3, out index3, lineNum2, index2);
                    after = string.Join(Environment.NewLine,
                        GetSubString(xmlLines, lineNum3, index3 + EndSymbol.Length).Select(x => x.Line));
                    xmlLines[xmlLines.Count - 1].Line = xmlLines.Last().Line.Substring(0, index3 + EndSymbol.Length);
                    return;
                }

                xmlLines.Add(myLine);
            }
        }

        private void AppendFormatError(MyLine myLine)
        {
            _append("Syntax: "+StartSymbol+"TagName"+EndSymbol+"  data "+StartSymbol+"/TagName"+EndSymbol, Color.Blue);
            _append("Format Error @ line "+ myLine.LineNum+" :", Color.Red);
            _append(myLine.Line, Color.Red);
            _append("", Color.Red);
            throw new FormatException();
        }

        public MyTag GetTag(List<MyLine> myLines)
        {
            MyTag myTag = new MyTag { Inner = myLines };
            int lineNum1; int index1;
            GetIndexof(myLines, StartSymbol, out lineNum1, out index1);
            if (index1 == -1)
            {
                AppendFormatError(myLines.First());
                return myTag;
            }
            int lineNum2; int index2;
            GetIndexof(myLines, " ", out lineNum2, out index2, lineNum1, index1);
            int lineNum3; int index3;
            GetIndexof(myLines, EndSymbol, out lineNum3, out index3, lineNum1, index1);
            int index4 = lineNum2 > lineNum3 || index2 > index3 || index2 == -1 ? index3 : index2;
            int lineNum4 =
                lineNum2 > lineNum3 || index2 > index3 || index2 == -1 ? lineNum3 : lineNum2;
            string before = string.Join(Environment.NewLine,
                GetSubString(myLines, 0, 0, lineNum1, index1).Select(x => x.Line));
            string startTag = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum1, index1, lineNum4, index4).Select(x => x.Line));
            string fullStartTag = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum1, index1, lineNum3, index3 + EndSymbol.Length).Select(x => x.Line));
            string name = startTag.Substring(StartSymbol.Length, startTag.Length - StartSymbol.Length);
            string attribute = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum4, index4, lineNum3, index3).Select(x => x.Line));
            List<MyAttribute> attributes = GetAttributes(attribute);
            string endTag = startTag.Replace(StartSymbol, StartSymbol + @"/");

            if (index3 == -1)
            {
                AppendFormatError(myLines.Find(x => x.LineNum == lineNum4));
                return myTag;
            }
            int lineNum5; int index5;
            GetIndexof(myLines, endTag, out lineNum5, out index5);
            if (index5 == -1)
            {
                AppendFormatError(myLines.Find(x => x.LineNum == lineNum3));
                return myTag;
            }
            int lineNum6; int index6;
            GetIndexof(myLines, EndSymbol, out lineNum6, out index6, lineNum5, index5);
            endTag = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum5, index5, lineNum6, index6 + EndSymbol.Length).Select(x => x.Line));
            List<MyLine> inner = GetSubString(myLines, lineNum3, index3 + EndSymbol.Length, lineNum5, index5);
            List<MyLine> outer = GetSubString(myLines, lineNum1, index1, lineNum6, index6 + EndSymbol.Length);
            string after = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum6, index6 + EndSymbol.Length).Select(x => x.Line));
            myTag.Name = name;
            myTag.Before = before;
            myTag.StartTag = startTag;
            myTag.FullStartTag = fullStartTag;
            myTag.Attributes = attributes;
            myTag.EndTag = endTag;
            myTag.Inner = inner;
            myTag.Outer = outer;
            myTag.After = after;

            return myTag;
        }

        private void SplitInner(List<MyLine> myLines, List<MyInner> myInners)
        {
            if (myLines.Count == 0)
            {
                return;
            }

            MyInner myTag = new MyInner { Text = myLines };
            int lineNum1; int index1;
            GetIndexof(myLines, StartSymbol, out lineNum1, out index1);
            if (index1 == -1)
            {
                myTag.NodeType = MyNodeType.Text;
                myTag.Name = "#text";
                myTag.Text = myLines;
                myInners.Add(myTag);
                return;
            }
            int lineNum2; int index2;
            GetIndexof(myLines, " ", out lineNum2, out index2, lineNum1, index1);
            int lineNum3; int index3;
            GetIndexof(myLines, EndSymbol, out lineNum3, out index3, lineNum1, index1);
            int index4 = lineNum2 > lineNum3 || index2 > index3 || index2 == -1 ? index3 : index2;
            int lineNum4 = lineNum2 > lineNum3 || index2 > index3 || index2 == -1 ? lineNum3 : lineNum2;

            List<MyLine> before = GetSubString(myLines, 0, 0, lineNum1, index1);
            string startTag = string.Join(Environment.NewLine,
                GetSubString(myLines, lineNum1, index1, lineNum4, index4).Select(x => x.Line));
            string name = startTag.Substring(StartSymbol.Length, startTag.Length - StartSymbol.Length);
            string endTag = startTag.Replace(StartSymbol, StartSymbol + @"/");
            if (index3 == -1)
            {
                AppendFormatError(myLines.Find(x => x.LineNum == lineNum4));
                return;
            }
            int lineNum5; int index5;
            GetIndexof(myLines, endTag, out lineNum5, out index5);
            if (index5 == -1)
            {
                AppendFormatError(myLines.Find(x => x.LineNum == lineNum3));
                return;
            }
            int lineNum6; int index6;
            GetIndexof(myLines, EndSymbol, out lineNum6, out index6, lineNum5, index5);
            List<MyLine> outer = GetSubString(myLines, lineNum1, index1, lineNum6, index6 + EndSymbol.Length);
            List<MyLine> after = GetSubString(myLines, lineNum6, index6 + EndSymbol.Length);
            myTag.Name = name;
            myTag.Text = outer;
            if (before.Any())
            {
                MyInner myTagBefore = new MyInner { NodeType = MyNodeType.Text, Name = "#text", Text = before };
                myInners.Add(myTagBefore);
            }

            myInners.Add(myTag);

            if (after.Any())
            {
                SplitInner(after, myInners);
            }
        }

        public static string GetEndTag(string text)
        {
            string startTag = GetStartTag(text);
            return startTag.Replace(StartSymbol, StartSymbol + @"/");
        }

        public static string GetStartTag(string text)
        {
            int index1 = text.IndexOf(StartSymbol, StringComparison.OrdinalIgnoreCase);
            if (index1 == -1)
            {
                return "";
            }

            int index2 = text.IndexOf(" ", index1, StringComparison.OrdinalIgnoreCase);
            int index3 = text.IndexOf(EndSymbol, index1, StringComparison.OrdinalIgnoreCase);
            int index = 0;
            if (index2 != -1 && index3 != -1)
            {
                index = index2 > index3 ? index3 : index2;
            }
            else if (index2 == -1 && index3 != -1)
            {
                index = index3;
            }
            else if (index2 != -1 && index3 == -1)
            {
                index = index2;
            }

            if (index == -1)
            {
                return "";
            }

            return text.Substring(index1, index - index1);
        }

        private static List<MyAttribute> GetAttributes(string text)
        {
            List<MyAttribute> attributes = new List<MyAttribute>();
            text = Regex.Replace(text, "= +", string.Empty);
            text = Regex.Replace(text, " +=", string.Empty);
            foreach (string data in text.Split(' '))
            {
                if (data.Contains("="))
                {
                    MyAttribute attribute = new MyAttribute
                    {
                        Name = data.Split('=')[0],
                        Value = data.Split('=')[1].Trim('"')
                    };
                    attributes.Add(attribute);
                }
            }

            return attributes;
        }
    }
}