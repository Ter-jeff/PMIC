using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;

namespace IgxlData.VBT
{
    public class VbtFunctionBase
    {
        public string FileName { get; set; }
        public string FunctionName { get; set; }
        public string Parameters { get; set; }
        public List<string> Args { get; set; }
        public string ParameterDefaults { get; set; }

        public VbtFunctionBase()
        {
            Args = Enumerable.Repeat("", 100).ToList();
            Parameters = "";
            FileName = "";            
        }

        public VbtFunctionBase(string functionName)
        {
            FunctionName = functionName;
            Args = Enumerable.Repeat("", 100).ToList();
            Parameters = "";
            FileName = "";
        }

        public void SetParamValue(string paramName, string paramValue)
        {
            if (string.IsNullOrEmpty(paramValue)) return;
            int index = (Parameters.Split(',').ToList()).FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            if (index != -1)
                Args[index] = paramValue;
        }

        public string GetParamValue(InstanceRow instanceRow, string paramName)
        {
            string paramValue = "";
            int index = (instanceRow.ArgList.Split(',').ToList()).FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            if (index != -1)
                paramValue = instanceRow.Args[index];
            return paramValue;
        }

        public void SetParamDefault()
        {
            var paramterList = Parameters.Split(',').ToList();
            var paramterDefaultList = ParameterDefaults.Split(',').ToList();
            for (int i = 0; i < paramterList.Count; i++)
            {
                if (Args[i] == "" && paramterDefaultList[i] != "")
                {
                    Args[i] = paramterDefaultList[i];
                }
            }

        }

        public void ReplaceVbt(InstanceRow instanceRow, Dictionary<string, string> dictionary)
        {
            instanceRow.Name = FunctionName;
            var args = instanceRow.ArgList.Split(',').ToList();
            Dictionary<string, string> dic = new Dictionary<string, string>();
            for (int i = 0; i < args.Count; i++)
            {
                dic.Add(args[i], instanceRow.Args[i]);
            }
            foreach (var item in dic)
            {
                SetParamValue(item.Key, item.Value);
                if (dictionary.ContainsKey(item.Key.ToUpper()))
                    SetParamValue(dictionary[item.Key.ToUpper()], item.Value);
            }

            instanceRow.ArgList = Parameters;
            instanceRow.Args = Args;
        }
    }
}
