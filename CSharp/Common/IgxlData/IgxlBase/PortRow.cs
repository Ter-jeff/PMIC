using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class PortRow
    {
        #region Constructor

        public PortRow()
        {
            ProtocolSettingValues = new List<string>();
            FunctionPropertyValues = new List<string>();
        }

        #endregion

        #region Property

        public int RowNum;
        public string PortName { get; set; }
        public string ProtocolFamily { get; set; }
        public string ProtocolType { get; set; }
        public string ProtocolSettings { get; set; }
        public List<string> ProtocolSettingValues { get; set; }
        public string FunctionName { get; set; }
        public string FunctionPin { get; set; }
        public string FunctionProperties { get; set; }
        public List<string> FunctionPropertyValues { get; set; }
        public string Comment { get; set; }
        public const int ConSettingNumber = 10;
        public const int ConPropertyNumber = 10;

        #endregion

        #region Member Function

        public void AddProperty(string property)
        {
            if (FunctionPropertyValues.Count > ConPropertyNumber)
                throw new Exception(string.Format("PortMap Property number has exceed the Max number: {0}",
                    ConPropertyNumber));
            FunctionPropertyValues.Add(property);
        }

        public void AddSetting(string setting)
        {
            if (FunctionPropertyValues.Count > ConPropertyNumber)
                throw new Exception(string.Format("PortMap setting number has exceed the Max number: {0}",
                    ConSettingNumber));
            ProtocolSettingValues.Add(setting);
        }

        #endregion
    }
}