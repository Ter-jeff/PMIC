using CLBistDataConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class CLBistDataLogRow
    {
        protected int _row;
        protected string _testNumber;
        protected string _testName;
        protected string _site;
        private string _pin;
        private string _low;
        private string _measured;
        private string _high;
        private string _force;
        private string _loc;

        private string _dacNumber;
        private string _phase;
        public string Site
        {
            get { return _site; }
        }

        public string TestName
        {
            get { return _testName; }
        }
        public string Pin
        {
            get{return _pin;}
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

        public string DacNumber
        {
            get { return _dacNumber; }
        }

        public string Phase
        {
            get { return _phase; }
        }

        public CLBistDataLogRow(int row, string logContent)
        {
            _row = row;
            Match match = RegStore.RegClBistDatalogRow.Match(logContent);
            _testNumber = match.Groups["Number"].ToString();
            _site = match.Groups["Site"].ToString();
            _testName = match.Groups["TestName"].ToString();
            _pin = match.Groups["Pin"].ToString();
            _low = match.Groups["Low"].ToString();
            _measured = match.Groups["Measured"].ToString();
            _high = match.Groups["High"].ToString();
            _force = match.Groups["Force"].ToString();
            _loc = match.Groups["Loc"].ToString();

            _dacNumber = RegStore.RegDacNumber.Match(_testName).Groups["dacNumber"].ToString();
            _phase = RegStore.RegPhase.Match(_pin).Groups["phase"].ToString();
        }
    }
}
