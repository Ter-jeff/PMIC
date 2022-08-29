using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.NonIgxlSheets;

namespace IgxlData.VBT
{
    public class VbtFunctionBase
    {
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

        public string FileName { get; set; }
        public string FunctionName { get; set; }
        public string Parameters { get; set; }
        public List<string> Args { get; set; }
        public string ParameterDefaults { get; set; }
        public bool CheckParam { get; set; }

        public void SetParamValue(string paramName, string paramValue)
        {
            if (string.IsNullOrEmpty(paramValue)) return;
            var index = Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            if (index != -1)
                Args[index] = paramValue;
        }

        public void SetParamValue(int index, string paramValue)
        {
            if (index > 0 && index < Args.Count)
                Args[index] = paramValue;
        }

        public string GetParamValue(InstanceRow instanceRow, string paramName)
        {
            var paramValue = "";
            var index = instanceRow.ArgList.Split(',').ToList()
                .FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            if (index != -1)
                paramValue = instanceRow.Args[index];
            return paramValue;
        }

        public void SetParamDefault()
        {
            var parameterList = Parameters.Split(',').ToList();
            var parameterDefaultList = ParameterDefaults.Split(',').ToList();
            for (var i = 0; i < parameterList.Count; i++)
                if (Args[i] == "" && parameterDefaultList[i] != "")
                    Args[i] = parameterDefaultList[i];
        }

        public void ReplaceVbt(InstanceRow instanceRow, Dictionary<string, string> dictionary)
        {
            instanceRow.Name = FunctionName;
            var args = instanceRow.ArgList.Split(',').ToList();
            var dic = new Dictionary<string, string>();
            for (var i = 0; i < args.Count; i++) dic.Add(args[i], instanceRow.Args[i]);
            foreach (var item in dic)
            {
                SetParamValue(item.Key, item.Value);
                if (dictionary.ContainsKey(item.Key.ToUpper()))
                    SetParamValue(dictionary[item.Key.ToUpper()], item.Value);
            }

            instanceRow.ArgList = Parameters;
            instanceRow.Args = Args;
        }

        public string GetParamValue(string paramName)
        {
            var paramValue = "";
            var index = Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals(paramName, StringComparison.OrdinalIgnoreCase));
            var aliasList = VbtFunctionLib.ParamMappingList.FirstOrDefault(s =>
                s.Find(a => a.Equals(paramName, StringComparison.OrdinalIgnoreCase)) != null);
            if (index != -1)
                paramValue = Args[index];
            else if (aliasList != null)
                foreach (var aliasName in aliasList)
                {
                    index = Parameters.Split(',').ToList()
                        .FindIndex(s => s.Equals(aliasName, StringComparison.OrdinalIgnoreCase));
                    if (index != -1)
                    {
                        paramValue = Args[index];
                        break;
                    }
                }

            return paramValue;
        }
    }
}