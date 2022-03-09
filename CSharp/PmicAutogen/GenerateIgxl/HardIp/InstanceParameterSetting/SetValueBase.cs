using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.EpplusErrorReport;
using IgxlData.NonIgxlSheets;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting
{
    public abstract class SetValueBase
    {
        public abstract void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function, string voltage);

        public void SetValueByParamMapping(VbtFunctionBase function, HardIpPattern pattern)
        {
            #region ParamMapping in test plan from "Misc Info" column

            var miscStr = pattern.MiscInfo;
            var miscInfoDic = new Dictionary<string, List<string>>();
            foreach (var parameter in miscStr.Split(';'))
            {
                var parameters = parameter.Split(':').ToList();
                var paramName = parameters[0];
                if (parameters.Count == 1 || Regex.IsMatch(paramName, HardIpConstData.SweepVoltage + @"\s*\(",
                    RegexOptions.IgnoreCase))
                    continue;
                if (!miscInfoDic.ContainsKey(paramName))
                    miscInfoDic.Add(paramName, new List<string>());
                parameters.RemoveAt(0);
                var paramValue = string.Join(":", parameters);
                miscInfoDic[paramName].Add(paramValue);
            }

            foreach (var param in miscInfoDic)
            {
                var paramFound = false;

                if (!HardIpConstData.MiscKeyList.Exists(s =>
                    Regex.IsMatch(param.Key, @"^(" + s + ")$", RegexOptions.IgnoreCase)))
                {
                    //function.SetParamValue(paramName, paramValue);
                    var index = function.Parameters.Split(',').ToList()
                        .FindIndex(s => s.Equals(param.Key, StringComparison.OrdinalIgnoreCase));
                    if (index != -1)
                    {
                        function.Args[index] = string.Join(";", param.Value);
                        paramFound = true;
                    }
                    else
                    {
                        var aliasList = VbtFunctionLib.ParamMappingList.Where(s =>
                            s.Find(a => a.Equals(param.Key, StringComparison.OrdinalIgnoreCase)) != null);
                        foreach (var aliasName in aliasList)
                        foreach (var item in aliasName)
                        {
                            var newIndex = function.Parameters.Split(',').ToList()
                                .FindIndex(s => s.Equals(item, StringComparison.OrdinalIgnoreCase));
                            if (newIndex != -1)
                            {
                                function.Args[newIndex] = string.Join(";", param.Value);
                                paramFound = true;
                                break;
                            }
                        }
                    }

                    if (!paramFound)
                    {
                        var miscInfoIndex =
                            HardIpDataMain.TestPlanData.PlanHeaderIdx[pattern.SheetName]["miscInfoIndex"];
                        var errorMessage = "Missing Parameter in " + function.FunctionName + " : " + param.Key +
                                           " Or wrong key word in misc-info";
                        EpplusErrorManager.AddError(HardIpErrorType.MissingParameter, ErrorLevel.Error,
                            pattern.SheetName, pattern.RowNum, miscInfoIndex, errorMessage, param.Key);
                    }
                }
            }

            #endregion
        }

        public void CheckInstArgument(VbtFunctionBase function, HardIpPattern pattern)
        {
            for (var i = 0; i < function.Args.Count; i++)
            {
                if (Regex.IsMatch(function.Args[i], @"^\d+[,]+[,|\d]*\d+$"))
                {
                    const string errorMessage = "IG-XL 9.0 suffered comma disappear issue";
                    if (!EpplusErrorManager.GetErrors().Any(x =>
                        x.SheetName == pattern.SheetName && x.RowNum == pattern.RowNum && x.Message == errorMessage))
                        EpplusErrorManager.AddError(HardIpErrorType.IgxlVersion.ToString(), ErrorLevel.Error,
                            pattern.SheetName, pattern.RowNum, errorMessage);
                }

                if (function.Args[i].Length > 8000)
                {
                    const string errorMessage = "IG-XL 9.0 need to keep the String length < 8000 characters ";
                    if (!EpplusErrorManager.GetErrors().Any(x =>
                        x.SheetName == pattern.SheetName && x.RowNum == pattern.RowNum && x.Message == errorMessage))
                        EpplusErrorManager.AddError(HardIpErrorType.IgxlVersion.ToString(), ErrorLevel.Error,
                            pattern.SheetName, pattern.RowNum, errorMessage);
                    //todo need check whether 8000 is correct
                    if (function.Args[i].Length > 30000)
                        function.Args[i] = function.Args[i].Substring(0, 30000);
                }
            }
        }
    }
}