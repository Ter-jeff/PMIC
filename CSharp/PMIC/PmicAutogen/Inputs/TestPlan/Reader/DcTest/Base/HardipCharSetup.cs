using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class HardipCharSetup : CharSetup
    {
        public static string GetShmooParameterName(string name)
        {
            var hardCodeDic = new Dictionary<string, string>
                {{"d0", "On"}, {"d1", "Data"}, {"d2", "Return"}, {"d3", "Off"}};
            if (CharStepConst.ParameterName.ContainsKey(name)) name = CharStepConst.ParameterName[name];

            if (hardCodeDic.ContainsKey(name.ToLower()))
                return hardCodeDic[name.ToLower()];

            return name;
        }

        public static string GetShmooTimeSets(string name)
        {
            if (name.Contains(","))
            {
                var arr = name.Split(',').ToList();
                arr.RemoveAt(0);
                return string.Join(",", arr);
            }

            return "";
        }

        #region Property

        public string TestNameInFlow { set; get; }
        public bool IsSplitByVoltage { set; get; }

        #endregion

        #region Constructor

        #endregion
    }
}