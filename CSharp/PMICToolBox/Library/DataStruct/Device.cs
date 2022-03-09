using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class Device
    {
        private string _deviceNumber;
        private List<TestInstanceItem> _testInstanceItemlist;

        public string DeviceNumber
        {
            get { return _deviceNumber; }
        }

        public List<TestInstanceItem> TestInstanceItemlist
        {
            get { return _testInstanceItemlist; }
        }

        public Device(string deviceNumber)
        {
            _deviceNumber = deviceNumber;
            _testInstanceItemlist = new List<TestInstanceItem>();
        }

        public List<InstanceCompareResult> Compare(Device compareDevice)
        {

            List<InstanceCompareResult> compareResultlist = new List<InstanceCompareResult>();
            if (compareDevice == null || compareDevice.TestInstanceItemlist.Count == 0)
                return compareResultlist;
            foreach (TestInstanceItem instance in this.TestInstanceItemlist)
            {
                TestInstanceItem refInstance = compareDevice.TestInstanceItemlist.Find(
                        s => s.InstanceName.Equals(instance.InstanceName, StringComparison.OrdinalIgnoreCase) &&
                        s.DuplicateIndex == instance.DuplicateIndex);

                //Instance only in real datalog
                if (refInstance == null)
                {
                    compareResultlist.Add(instance.CovertToReportRow(DiffResultType.OnlyInBaseDatalog, instance.Row.ToString(), ""));
                    continue;
                }

                //Compare Instance 
                InstanceCompareResult instanceCompResult = null;
                bool result = instance.Compare(refInstance, out instanceCompResult);
                if (result == false)
                {
                    compareResultlist.Add(instanceCompResult);
                }
            }

            //Instance Only in Reference datalog
            foreach (TestInstanceItem instance in compareDevice.TestInstanceItemlist)
            {
                TestInstanceItem realInstance = this.TestInstanceItemlist.Find(
                    s => s.InstanceName.Equals(instance.InstanceName, StringComparison.OrdinalIgnoreCase) &&
                    s.DuplicateIndex == instance.DuplicateIndex);

                //Instance only in compare datalog
                if (realInstance == null)
                {
                    compareResultlist.Add(instance.CovertToReportRow(DiffResultType.OnlyInCompareDatalog, "", instance.Row.ToString()));
                    continue;
                }
            }

            return compareResultlist;
        }
    }
}
