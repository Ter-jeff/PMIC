using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProfileTool_PMIC.Reader
{
    public class ExecutionProfileRow
    {
        #region Field

        #endregion

        #region Properity
        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string Flowsheet { get; set; }
        public string Line { get; set; }
        public string Flowstep { get; set; }
        public double Total { get; set; }
        public string Prebody { get; set; }
        public string Body { get; set; }
        public string Postbody { get; set; }
        public string Tname { get; set; }
        public string Tnum { get; set; }
        #endregion

        #region Constructor
        public ExecutionProfileRow()
        {
            Flowsheet = "";
            Line = "";
            Flowstep = "";
            Total = 0;
            Prebody = "";
            Body = "";
            Postbody = "";
            Tname = "";
            Tnum = "";
        }

        public ExecutionProfileRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Flowsheet = "";
            Line = "";
            Flowstep = "";
            Total = 0;
            Prebody = "";
            Body = "";
            Postbody = "";
            Tname = "";
            Tnum = "";
        }
        #endregion
    }

    public class ExecutionProfileReader
    {
        private const string ConHeaderFlowsheet = "Flow Sheet";
        private const string ConHeaderLine = "Line";
        private const string ConHeaderFlowstep = "Flow Step";
        private const string ConHeaderTotal = "Total";
        private const string ConHeaderPrebody = "PreBody";
        private const string ConHeaderBody = "Body";
        private const string ConHeaderPostbody = "PostBody";
        private const string ConHeaderTname = "TName";
        private const string ConHeaderTnum = "TNum";

        private int _flowsheetIndex = -1;
        private int _lineIndex = -1;
        private int _flowstepIndex = -1;
        private int _totalIndex = -1;
        private int _prebodyIndex = -1;
        private int _bodyIndex = -1;
        private int _postbodyIndex = -1;
        private int _tnameIndex = -1;
        private int _tnumIndex = -1;
        private int _headerindex = -1;
        private Regex _rgexSepArray = new Regex(@"\t", RegexOptions.Compiled);
        private ProfileToolForm _profileToolForm ;

        public ExecutionProfileReader(ProfileToolForm profileToolForm)
        {
            _profileToolForm = profileToolForm;
        }

        public List<ExecutionProfileRow> ReadIgxl90(string fileName)
        {
            var lines = File.ReadAllLines(fileName).ToList();
            var executionProfileIgRows = new List<ExecutionProfileRow>();

            GetHeaderIndex90(lines);

            if (_headerindex != -1)
            {
                var count = _headerindex;
                foreach (var line in lines.GetRange(_headerindex, lines.Count - _headerindex))
                {
                    count++;
                    var array = _rgexSepArray.Split(line);
                    var executionProfileRow = new ExecutionProfileRow();
                    executionProfileRow.Flowsheet = _flowsheetIndex != -1 && _flowsheetIndex < array.Count() ? array[_flowsheetIndex].Trim() : "";
                    executionProfileRow.Line = _lineIndex != -1 && _lineIndex < array.Count() ? array[_lineIndex].Trim() : "";
                    executionProfileRow.Flowstep = _flowstepIndex != -1 && _flowstepIndex < array.Count() ? array[_flowstepIndex].Trim() : "";
                    if (_totalIndex != -1 && _totalIndex < array.Count())
                    {
                        if (string.IsNullOrEmpty(array[_totalIndex]))
                        {
                            _profileToolForm.AppendText(string.Format("Cells[{0},{1}] is empty", count, _totalIndex + 1), Color.Red);
                            executionProfileRow.Total = 0;
                        }    
                        else
                            executionProfileRow.Total = double.Parse(array[_totalIndex].Trim());  
                    }
                    else
                        executionProfileRow.Total = 0;
                    //executionProfileRow.Total = _totalIndex != -1 && _totalIndex < array.Count() ? double.Parse(array[_totalIndex].Trim()) : 0;
                    executionProfileRow.Prebody = _prebodyIndex != -1 && _prebodyIndex < array.Count() ? array[_prebodyIndex].Trim() : "";
                    executionProfileRow.Body = _bodyIndex != -1 && _bodyIndex < array.Count() ? array[_bodyIndex].Trim() : "";
                    executionProfileRow.Postbody = _postbodyIndex != -1 && _postbodyIndex < array.Count() ? array[_postbodyIndex].Trim() : "";
                    executionProfileRow.Tname = _tnameIndex != -1 && _tnameIndex < array.Count() ? array[_tnameIndex].Trim() : "";
                    executionProfileRow.Tnum = _tnumIndex != -1 && _tnumIndex < array.Count() ? array[_tnumIndex].Trim() : "";
                    executionProfileIgRows.Add(executionProfileRow);
                }
            }
            return executionProfileIgRows;
        }

        public List<ExecutionProfileRow> Read(string fileName)
        {
            var lines = File.ReadAllLines(fileName).ToList();
            var executionProfileIgRows = new List<ExecutionProfileRow>();

            if (IsMatchIgxl830(lines))
                GetHeaderIndex80(lines);
            else
                GetHeaderIndex90(lines);

            if (_headerindex != -1)
            {
                foreach (var line in lines.GetRange(_headerindex, lines.Count - _headerindex))
                {
                    var array = _rgexSepArray.Split(line);
                    var executionProfileRow = new ExecutionProfileRow();
                    executionProfileRow.Flowsheet = _flowsheetIndex != -1 && _flowsheetIndex < array.Count() ? array[_flowsheetIndex].Trim() : "";
                    executionProfileRow.Line = _lineIndex != -1 && _lineIndex < array.Count() ? array[_lineIndex].Trim() : "";
                    executionProfileRow.Flowstep = _flowstepIndex != -1 && _flowstepIndex < array.Count() ? array[_flowstepIndex].Trim() : "";
                    executionProfileRow.Total = _totalIndex != -1 && _totalIndex < array.Count() ? double.Parse(array[_totalIndex].Trim()) : 0;
                    executionProfileRow.Prebody = _prebodyIndex != -1 && _prebodyIndex < array.Count() ? array[_prebodyIndex].Trim() : "";
                    executionProfileRow.Body = _bodyIndex != -1 && _bodyIndex < array.Count() ? array[_bodyIndex].Trim() : "";
                    executionProfileRow.Postbody = _postbodyIndex != -1 && _postbodyIndex < array.Count() ? array[_postbodyIndex].Trim() : "";
                    executionProfileRow.Tname = _tnameIndex != -1 && _tnameIndex < array.Count() ? array[_tnameIndex].Trim() : "";
                    executionProfileRow.Tnum = _tnumIndex != -1 && _tnumIndex < array.Count() ? array[_tnumIndex].Trim() : "";
                    executionProfileIgRows.Add(executionProfileRow);
                }
            }
            return executionProfileIgRows;
        }

        private void GetHeaderIndex90(List<string> lines)
        {
            var rgexHeader1Start90 = new Regex(@"Flow Sheet\tLine\tFlow Step", RegexOptions.IgnoreCase);
            var rgexHeader2Start90 = new Regex(@"Total\tPreBody\tBody\tPostBody\tTName\tTNum", RegexOptions.IgnoreCase);

            for (var index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (line == null | line == string.Empty)
                    continue;

                if (rgexHeader1Start90.IsMatch(line) && rgexHeader2Start90.IsMatch(lines[index + 1]))
                {
                    var header1Array = _rgexSepArray.Split(line).ToList();
                    for (var i = 0; i < header1Array.Count(); i++)
                    {
                        var header = header1Array[i];
                        GetHeaderIndex(header, i);
                    }

                    var header2Array = _rgexSepArray.Split(lines[index + 1]).ToList();
                    for (var i = 0; i < header2Array.Count(); i++)
                    {
                        var header = header2Array[i];
                        GetHeaderIndex(header, i);
                    }
                    _headerindex = index + 2;
                    break;
                }
            }
        }

        private void GetHeaderIndex80(List<string> lines)
        {
            var rgexHeader1Start83 = new Regex(@"Flow Step\tLine\tTotal\tPreBody\tBody\tPostBody", RegexOptions.IgnoreCase);
            for (var index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (line == null | line == string.Empty)
                    continue;
                if (rgexHeader1Start83.IsMatch(line))
                {
                    var header1Array = _rgexSepArray.Split(line).ToList();
                    for (var i = 0; i < header1Array.Count(); i++)
                    {
                        var header = header1Array[i];
                        GetHeaderIndex(header, i);
                    }
                    _headerindex = index + 1;
                    break;
                }
            }
        }

        private bool IsMatchIgxl830(List<string> lines)
        {
            var regxIgxl83 = new Regex(@"IG-XL Version\s+8.3", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var regxIgxl90 = new Regex(@"IG-XL Version\s+9.1", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            for (var index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (regxIgxl83.IsMatch(line))
                    return true;
                if (regxIgxl90.IsMatch(line))
                    return false;
            }
            return false;
        }

        private void GetHeaderIndex(string header, int i)
        {
            if (header.Equals(ConHeaderFlowsheet, StringComparison.Ordinal))
                _flowsheetIndex = i;
            else if (header.Equals(ConHeaderLine, StringComparison.CurrentCultureIgnoreCase))
                _lineIndex = i;
            else if (header.Equals(ConHeaderFlowstep, StringComparison.CurrentCultureIgnoreCase))
                _flowstepIndex = i;
            else if (header.Equals(ConHeaderTotal, StringComparison.CurrentCultureIgnoreCase))
                _totalIndex = i;
            else if (header.Equals(ConHeaderPrebody, StringComparison.CurrentCultureIgnoreCase))
                _prebodyIndex = i;
            else if (header.Equals(ConHeaderBody, StringComparison.CurrentCultureIgnoreCase))
                _bodyIndex = i;
            else if (header.Equals(ConHeaderPostbody, StringComparison.CurrentCultureIgnoreCase))
                _postbodyIndex = i;
            else if (header.Equals(ConHeaderTname, StringComparison.CurrentCultureIgnoreCase))
                _tnameIndex = i;
            else if (header.Equals(ConHeaderTnum, StringComparison.CurrentCultureIgnoreCase))
                _tnumIndex = i;
        }
    }
}