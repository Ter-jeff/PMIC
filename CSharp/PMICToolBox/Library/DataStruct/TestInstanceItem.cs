using Library.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class TestInstanceItem
    {
        private string _deviceName;
        private string _instanceNumber;
        private int _duplicateIndex;
        private int _row;
        //Dig Source
        private string _srcBits;
        private List<string> _srcPinlst;
        private List<string> _dataSequencelst;
        private List<string> _assignmentlst;
        //Dig Capture
        private string _capBits;
        private List<string> _capPinlst;
        private List<string> _dsscOutlst;
        private List<TestInstanceLogBase> _testInstanceLoglst;

        public Regex RegRealyOnOff = Common.CommonData.GetInstance().LogSettings.GetIgnoredItemPatternByName("RelayOnOffInstance").Pattern;
        //public Regex RegRealyOnOff = new Regex(@"Relay[_]?(on|off)", RegexOptions.IgnoreCase);

        public string DeviceNumber
        {
            get { return _deviceName; }
        }

        public string InstanceName
        {
            get { return _instanceNumber; }
        }

        public int Row
        {
            get { return _row; }
        }

        public int DuplicateIndex
        {
            get { return _duplicateIndex; }
        }
        public string SrcBits
        {
            get { return _srcBits; }
            set { _srcBits = value; }
        }

        public List<string> SrcPinlst
        {
            get { return _srcPinlst; }
        }

        public List<string> DataSequencelst
        {
            get { return _dataSequencelst; }
        }

        public List<string> Assignmentlst
        {
            get { return _assignmentlst; }
        }

        public string CapBits
        {
            get { return _capBits; }
            set { _capBits = value; }
        }

        public List<string> CapPinlst
        {
            get { return _capPinlst; }
        }

        public List<string> DsscOutlst
        {
            get { return _dsscOutlst; }
        }

        public List<TestInstanceLogBase> TestInstanceLoglst
        {
            get { return _testInstanceLoglst; }
            set { _testInstanceLoglst = value; }
        }

        public TestInstanceItem(string instanceName, string deviceName, int row, int duplicateIndex, List<TestInstanceLogBase> testInstanceLoglst)
        {
            _deviceName = deviceName;
            _instanceNumber = instanceName;
            _row = row;
            _duplicateIndex = duplicateIndex;
            _testInstanceLoglst = testInstanceLoglst;
        }
        

        public void setSrcPins(string srcPins)
        {
            if (!string.IsNullOrEmpty(srcPins))
            {
                _srcPinlst = srcPins.Split(',').ToList();
            }
        }

        public void setDataSequences(string dataSequence)
        {
            if (!string.IsNullOrEmpty(dataSequence))
            {
                _dataSequencelst = dataSequence.Split('+').ToList();
            }
        }

        public void setAssignments(string assignments)
        {
            if (!string.IsNullOrEmpty(assignments))
            {
                _assignmentlst = assignments.Split(';').ToList();
            }
        }

        public void setCapPins(string capPins)
        {
            if (!string.IsNullOrEmpty(capPins))
            {
                _capPinlst = capPins.Split(',').ToList();
            }
        }

        public void setDsscout(string dsscOut)
        {
            if (!string.IsNullOrEmpty(dsscOut))
            {
                _dsscOutlst = dsscOut.Split(',').ToList();
            }
        }

        public bool Compare(TestInstanceItem refInstance, out InstanceCompareResult instanceDiffResultRow)
        {
            bool result = true;
            List<DiffResultLogRow> logDiffResultlst = new List<DiffResultLogRow>();
            instanceDiffResultRow = new InstanceCompareResult(_deviceName, _instanceNumber, DiffResultType.LimitChange, logDiffResultlst);
            instanceDiffResultRow.Row = this.Row.ToString();
            instanceDiffResultRow.RefLogFileRow = refInstance.Row.ToString();
            //Compare Instance item
            //instanceDiffResultRow.SrcBits = _srcBits;
            //if (!_srcBits.Equals(refInstance.SrcBits))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefSrcBits = refInstance.SrcBits;
            //}

            //instanceDiffResultRow.SrcPin = Utility.ConvertListToString(_srcPinlst, ",");
            //if (!Utility.CompareTwoListItem(_srcPinlst, refInstance.SrcPinlst))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefSrcPin = Utility.ConvertListToString(refInstance.SrcPinlst, ",");
            //}

            //instanceDiffResultRow.DataSequence = Utility.ConvertListToString(_dataSequencelst, "+");
            //if (!Utility.CompareTwoListItem(_dataSequencelst, refInstance.DataSequencelst))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefDataSequence = Utility.ConvertListToString(refInstance.DataSequencelst, "+");
            //}

            //instanceDiffResultRow.Assignments = Utility.ConvertListToString(_assignmentlst, ";");
            //if (!Utility.CompareTwoListItem(_assignmentlst, refInstance.Assignmentlst))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefAssignment = Utility.ConvertListToString(refInstance.Assignmentlst, ";");
            //}

            //instanceDiffResultRow.CapBits = _capBits;
            //if (!_capBits.Equals(refInstance.CapBits))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefCapBits = refInstance.CapBits;
            //}

            //instanceDiffResultRow.CapPin = Utility.ConvertListToString(_capPinlst, "+");
            //if (!Utility.CompareTwoListItem(_capPinlst, refInstance.CapPinlst))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefCapPin = Utility.ConvertListToString(refInstance.CapPinlst, "+");
            //}

            //instanceDiffResultRow.DsscOut = Utility.ConvertListToString(_dsscOutlst, "+");
            //if (!Utility.CompareTwoListItem(_dsscOutlst, refInstance.DsscOutlst))
            //{
            //    result = false;
            //    instanceDiffResultRow.RefDsscOut = Utility.ConvertListToString(refInstance.DsscOutlst, "+");
            //}

            //Compare logs
            foreach (TestInstanceLogBase realLogItem in this.TestInstanceLoglst)
            {
                TestInstanceLogBase referenceLogItem = refInstance.TestInstanceLoglst.Find(
                    s => s.KeyWord.Equals(realLogItem.KeyWord, StringComparison.OrdinalIgnoreCase) && s.DuplicateIndex == realLogItem.DuplicateIndex);             

                DiffResultLogRow logResultRow = null;
                //Log only in real datalog instance
                if (referenceLogItem == null)
                {
                    result = false;
                    logResultRow = realLogItem.ConvertToReportRow(realLogItem.Row.ToString(),"");
                    logResultRow.Result = DiffResultType.OnlyInBaseDatalog;
                    instanceDiffResultRow.Result = DiffResultType.TestItemMismatch;
                    logDiffResultlst.Add(logResultRow);
                    continue;
                }
                //Compare real datalog item and reference datalog item
                bool logEquals = realLogItem.Compare(referenceLogItem, out logResultRow);
                if (logEquals == false)
                {
                    result = false;
                    logDiffResultlst.Add(logResultRow);
                }
            }

            //Log only in reference datalog instance
            foreach (TestInstanceLogBase referenceLogItem in refInstance.TestInstanceLoglst)
            {
                TestInstanceLogBase realLogItem =
                    this.TestInstanceLoglst.Find(
                    s => s.KeyWord.Equals(referenceLogItem.KeyWord, StringComparison.OrdinalIgnoreCase) && s.DuplicateIndex == referenceLogItem.DuplicateIndex);             

                if (realLogItem == null)
                {
                    result = false;
                    DiffResultLogRow diffResultRow = referenceLogItem.ConvertToReportRow("", referenceLogItem.Row.ToString());
                    diffResultRow.Result = DiffResultType.OnlyInCompareDatalog;
                    instanceDiffResultRow.Result = DiffResultType.TestItemMismatch;
                    logDiffResultlst.Add(diffResultRow);
                    continue;
                }

            }

            return result;
        }

        public InstanceCompareResult CovertToReportRow(DiffResultType result, string row, string refRow)
        {
            InstanceCompareResult instanceReportRow = new InstanceCompareResult(_deviceName, _instanceNumber,result, null);
            instanceReportRow.Row = row;
            instanceReportRow.RefLogFileRow = refRow;
            //instanceReportRow.SrcBits = _srcBits;
            //instanceReportRow.SrcPin = Utility.ConvertListToString(_srcPinlst,",");
            //instanceReportRow.DataSequence = Utility.ConvertListToString( _dataSequencelst,"+");
            //instanceReportRow.Assignments = Utility.ConvertListToString( _assignmentlst,";");
            //instanceReportRow.CapBits = _capBits;
            //instanceReportRow.CapPin = Utility.ConvertListToString( _capPinlst,",");
            //instanceReportRow.DsscOut = Utility.ConvertListToString( _dsscOutlst,",");
            return instanceReportRow;
        }

        public bool IsValidInstance()
        {
            if (RegRealyOnOff.IsMatch(this.InstanceName))
                return false;

            if (this.TestInstanceLoglst != null && this.TestInstanceLoglst.Count > 0)
                return true;            
            //if (!string.IsNullOrEmpty(this.SrcBits))
            //    return true;
            //if (!string.IsNullOrEmpty(this.CapBits))
            //    return true;
            //if (this.SrcPinlst != null && this.SrcPinlst.Count > 0)
            //    return true;            
            //if (this.CapPinlst != null && this.CapPinlst.Count > 0)
            //    return true;
            //if (this.DsscOutlst != null && this.DsscOutlst.Count > 0)
            //    return true;
            //if (this.DataSequencelst != null && this.DataSequencelst.Count > 0)
            //    return true;
            //if (this.Assignmentlst != null && this.Assignmentlst.Count > 0)
            //    return true;

            return false;
        }
        
    }
}
