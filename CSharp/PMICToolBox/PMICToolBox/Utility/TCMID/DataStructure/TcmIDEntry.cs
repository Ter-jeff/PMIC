using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CompareStatus = PmicAutomation.Utility.TCMIDComparator.DataStructure.EnumStore.CompareStatus;

namespace PmicAutomation.Utility.TCMID.DataStructure
{
    public class TcmIdEntry
    {
        private string _flowtable;
        private string _testname;
        private string _tcmId;
        private string _originalTcmId;
        private string _scale;
        private string _units;
        private string _lowlim;
        private string _hilim;
        private bool _resettable;
        private CompareStatus _status;

        public TcmIdEntry(string flowtable, string testname, string tcmId, string scale, string units, string lowlim, string hilim)
        {
            _flowtable = flowtable;
            _testname = testname;
            _tcmId = tcmId;
            _originalTcmId = tcmId;
            _scale = scale;
            _units = units;
            _lowlim = lowlim;
            _hilim = hilim;
            _resettable = true;
            _status = CompareStatus.NA;
        }

        public string OriginalTcmId
        {
            get { return _originalTcmId; }
            set { _originalTcmId = value; }
        }

        public CompareStatus Status
        {
            get { return _status; }
            set { _status = value; }
        }

        public string Flowtable
        {
            get { return _flowtable; }
        }

        public string Testname
        {
            get { return _testname; }
        }

        public string TcmId
        {
            get { return _tcmId; }
            set { _tcmId = value; }
        }

        public string Scale
        {
            get { return _scale; }
        }

        public string Units
        {
            get { return _units; }
        }

        public string LowLim
        {
            get { return _lowlim; }
        }

        public string HiLim
        {
            get { return _hilim; }
        }

        public bool Resettable
        {
            get { return _resettable; }
            set { _resettable = value; }
        }
    }
}
