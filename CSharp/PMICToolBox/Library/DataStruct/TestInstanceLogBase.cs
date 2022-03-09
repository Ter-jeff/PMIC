using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public abstract class TestInstanceLogBase
    {
        protected int _row;
        protected string _instanceName;
        protected string _testNumber;
        protected string _testName;
        protected string _site;
        protected string _keyWord;
        protected int _duplicateIndex;
        //Force Consition
        protected List<string> _forceConditionlst;

        public int Row
        {
            get { return _row; }
        }

        public int DuplicateIndex
        {
            get { return _duplicateIndex; }
            set { _duplicateIndex = value; }
        }

        public string KeyWord
        {
            get { return _keyWord; }
        }

        public string InstanceName
        {
            get { return _instanceName; }
            set { _instanceName = value; }
        }

        public string TestNumber
        {
            get { return _testNumber; }
        }

        public string TestName
        {
            get{return _testName;}
        }

        public string Site
        {
            get { return _site; }
        }

        public List<string> ForceConditionlst
        {
            get { return _forceConditionlst; }
        }

        public void setForceConditions(List<string> forceCondition)
        {
            _forceConditionlst = forceCondition;
        }
        public abstract bool Compare(TestInstanceLogBase testInstancelog, out DiffResultLogRow summaryRow);
        public abstract DiffResultLogRow ConvertToReportRow(string row, string refRow);
    }
}
