using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class InstanceCompareResult
    {
        private string _deviceNumber;
        private string _instanceName;
        public string Row = string.Empty;
        public string RefLogFileRow = string.Empty;
        //Dig Source
        public string SrcBits;
        public string SrcPin;
        public string DataSequence;
        public string Assignments;
        //Dig Capture
        public string CapBits;
        public string CapPin;
        public string DsscOut;

        //Reference datalog value(when Result is Diff, the reference value will be stored)
        public string RefSrcBits = null;
        public string RefSrcPin = null;
        public string RefDataSequence = null;
        public string RefAssignment = null;
        public string RefCapBits = null;
        public string RefCapPin = null;
        public string RefDsscOut = null;
    

        private DiffResultType _result;
        private List<DiffResultLogRow> _logDiffResultlst;

        public string DeviceNumber
        {
            get { return _deviceNumber; }
        }

        public string InstanceName
        {
            get { return _instanceName; }
        }

        public DiffResultType Result
        {
            get { return _result; }
            set { _result = value; }
        }

        public List<DiffResultLogRow> LogDiffResultlst
        {
            get { return _logDiffResultlst; }
        }

        public InstanceCompareResult(string deviceNumber, string instanceName, DiffResultType result,
            List<DiffResultLogRow> logDiffResultlst)
        {
            _deviceNumber = deviceNumber;
            _instanceName = instanceName;
            _result = result;
            _logDiffResultlst = logDiffResultlst;
        }
    }
}
