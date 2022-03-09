using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class DiffResultLogRow
    {
        public DiffResultType Result;
        public string ComparedInst = string.Empty;
        public string BasedInst = string.Empty;
        public string Row = string.Empty;
        public string RefLogFileRow = string.Empty;
        public string Site = string.Empty;
        public string TestName = string.Empty;
        public string MeasurePin = string.Empty;
        public string LimitHigh = string.Empty;
        public string LimitLow = string.Empty;
        public string ForceValue = string.Empty;
        //Force Consition
        public string ForceCondition;
        public string RefForceCondition = null;

        //Reference datalog value(when Result is Diff, the reference value will be stored)
        public string RefMeasurePin = null;
        public string RefLimitHigh = null;
        public string RefLimitLow = null;
        public string RefForceValue = null;
    }

    public enum DiffResultType
    {
        Diff,
        LimitChange,
        TestItemMismatch,
        OnlyInBaseDatalog,
        OnlyInCompareDatalog
    }
}
