using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Library.Common;

namespace Library.DataStruct
{
    public class TestInstanceLogMeasure : TestInstanceLogBase
    {
        public static string LogType = "MeasureTest";
        private string _pin;
        private string _channel;
        private string _low;
        private string _measured;
        private string _high;
        private string _force;
        private string _loc;

        public string Pin
        {
            get{return _pin;}
        }

        public string Channel
        {
            get { return _channel; }
        }

        public string Low
        {
            get { return _low; }
        }

        public string Measured
        {
            get { return _measured; }
        }

        public string High
        {
            get { return _high; }
        }

        public string Force
        {
            get { return _force; }
        }

        public string Loc
        {
            get { return _loc; }
        }

        public TestInstanceLogMeasure(int row, string logContent)
        {
            _row = row;
            HeaderPattern logPattern = CommonData.GetInstance().LogSettings.GetHeaderPatternByName(LogType);
            Match match = logPattern.DataRegex.Match(logContent);
            _testNumber = match.Groups["Number"].ToString();
            _site = match.Groups["Site"].ToString();
            _testName = match.Groups["TestName"].ToString();
            _pin = match.Groups["Pin"].ToString();
            _channel = match.Groups["Channel"].ToString();
            _low = match.Groups["Low"].ToString();
            _measured = match.Groups["Measured"].ToString();
            _high = match.Groups["High"].ToString();
            _force = match.Groups["Force"].ToString();
            _loc = match.Groups["Loc"].ToString();

            _keyWord = _testName + "&" + _site + "&" + _pin;
            //string[] contextArr = Regex.Split(logContent.Trim(), @"[\s]+");

            //if (contextArr.Length < 9)
            //{
            //    throw new Exception("Log context is not a Measre log: " + logContent);
            //}

            //_testNumber = contextArr[0];
            //_site = contextArr[1];
            //_testName = contextArr[2];

            //int columnIndex = 3;
            ////Pin(sometimes will not have pin in datalog)
            //if (IsPin(contextArr[columnIndex]))
            //{
            //    _pin = contextArr[columnIndex];
            //    columnIndex++;
            //}
            ////Channel
            //_channel = contextArr[columnIndex];
            //columnIndex++;
            ////low(may be contains unit eg: 10 mA)
            //_low = contextArr[columnIndex];
            //columnIndex++;
            //if (!IsNumber(contextArr[columnIndex]) && !IsNA(contextArr[columnIndex]))
            //{
            //    //Add unit
            //    _low = _low + " " + contextArr[columnIndex];
            //    columnIndex++;
            //}
            ////measured(may be contains unit AND result eg: 10 mA (F))
            //_measured = contextArr[columnIndex];
            //columnIndex++;
            //if (!IsNumber(contextArr[columnIndex]) && !IsNA(contextArr[columnIndex]))
            //{
            //    //Add unit
            //    _measured = _measured + " " + contextArr[columnIndex];
            //    columnIndex++;
            //}
            //if (!IsNumber(contextArr[columnIndex]) && !IsNA(contextArr[columnIndex]))
            //{
            //    //Add result
            //    _measured = _measured + " " + contextArr[columnIndex];
            //    columnIndex++;
            //}
            ////high(may be contains unit eg: 10 mA)
            //_high = contextArr[columnIndex];
            //columnIndex++;
            //if (!IsNumber(contextArr[columnIndex]) && !IsNA(contextArr[columnIndex]))
            //{
            //    //Add unit
            //    _high = _high + " " + contextArr[columnIndex];
            //    columnIndex++;
            //}
            ////force(may be contains unit eg: 10 mA)
            //_force = contextArr[columnIndex];
            //columnIndex++;
            //if (columnIndex >= contextArr.Length)
            //{
            //    string aa = "";
            //}
            //if (!IsNumber(contextArr[columnIndex]) && !IsNA(contextArr[columnIndex]))
            //{
            //    //Add unit
            //    _force = _force + " " + contextArr[columnIndex];
            //    columnIndex++;
            //}
            ////loc
            //_loc = contextArr[columnIndex];
        }

        private bool IsPin(string input)
        {
            if (!Regex.IsMatch(input, @"[\w]+", RegexOptions.IgnoreCase))
                return false;
            if (!Regex.IsMatch(input, @"[a-zA-Z]+", RegexOptions.IgnoreCase))
                return false;
            return true;
        }

        private bool IsNumber(string input)
        {
            return Regex.IsMatch(input, @"[\-]?([\d]+[\.]?[\d]*)|([\.]?[\d]+)", RegexOptions.IgnoreCase);
        }

        private bool IsNA(string input)
        {
            input = input.Trim().ToUpper();
            if(input == "NA" || input == "N/A" || input == "NULL")
                return true;
            return false;
        }

        public override bool Compare(TestInstanceLogBase testInstancelog, out DiffResultLogRow diffResultRow)
        {
            TestInstanceLogMeasure reftestInstancelog = (TestInstanceLogMeasure)testInstancelog;
            bool result = true;

            diffResultRow = new DiffResultLogRow();
            diffResultRow.Site = _site;
            diffResultRow.BasedInst = InstanceName;
            diffResultRow.ComparedInst = testInstancelog.InstanceName;
            diffResultRow.TestName = _testName;
            diffResultRow.Row = this.Row.ToString();
            diffResultRow.RefLogFileRow = reftestInstancelog.Row.ToString();
            //diffResultRow.MeasurePin = _pin;
            //if (!_pin.Equals(reftestInstancelog.Pin))
            //{
            //    result = false;
            //    diffResultRow.RefMeasurePin = reftestInstancelog.Pin;
            //}

            //diffResultRow.ForceCondition = Utility.ConvertListToString(_forceConditionlst, ";");
            //if (!Utility.CompareTwoListItem(_forceConditionlst, reftestInstancelog.ForceConditionlst))
            //{
            //    result = false;
            //    diffResultRow.RefForceCondition = Utility.ConvertListToString(reftestInstancelog.ForceConditionlst, ";");
            //}

            //diffResultRow.ForceValue = _force;
            //if (!_force.Equals(reftestInstancelog.Force))
            //{
            //    result = false;
            //    diffResultRow.RefForceValue = reftestInstancelog.Force;
            //}

            diffResultRow.LimitLow = _low;
            if (!_low.Equals(reftestInstancelog.Low))
            {
                result = false;
                diffResultRow.RefLimitLow = reftestInstancelog.Low;
            }

            diffResultRow.LimitHigh = _high;
            if (!_high.Equals(reftestInstancelog.High))
            {
                result = false;
                diffResultRow.RefLimitHigh = reftestInstancelog.High;
            }

            if (result == false)
            {
                diffResultRow.Result = DiffResultType.Diff;
            }
            return result;
        }

        public override DiffResultLogRow ConvertToReportRow(string row, string refRow)
        {
            DiffResultLogRow summaryRow = new DiffResultLogRow();
            summaryRow.Row = row;
            summaryRow.RefLogFileRow = refRow;
            summaryRow.Site = _site;
            summaryRow.TestName = _testName;
            //summaryRow.MeasurePin = _pin;
            //summaryRow.ForceCondition = Utility.ConvertListToString(_forceConditionlst, ";");
            //summaryRow.ForceValue = _force;
            summaryRow.LimitLow = _low;
            summaryRow.LimitHigh = _high;
            return summaryRow;
        }
    }
}
