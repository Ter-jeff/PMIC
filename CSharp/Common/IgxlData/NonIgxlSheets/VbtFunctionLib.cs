using CommonLib.Enum;
using CommonLib.ErrorReport;
using IgxlData.VBT;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.NonIgxlSheets
{
    public class VbtFunctionLib
    {
        public const string FunctionalCharName = "Functional_T_char";
        public const string FunctionalTUpdated = "Functional_T_updated";
        public const string DvdcTrim = "wi_trimuniversalfunc";
        public const string LcdMeas = "measuniversalfunc";
        public const string VifName = "meas_freqvoltcurr_universal_func";
        public const string VirName = "meas_vir_io_universal_func";
        public const string VdiffFunc = "meas_vdiff_func";
        public const string FreqSynMeasFreqCurr = "freqsyn_measfreqcurr_func";

        public static List<List<string>> ParamMappingList = new List<List<string>>();
        public static Dictionary<string, int> GeneratedVbtFunctionDic = new Dictionary<string, int>();
        private static readonly List<string> MissingParams = new List<string>();

        public VbtFunctionLib()
        {
            VbtLib = new List<VbtFunctionBase>();
            GeneratedVbtFunctionDic.Clear();
        }

        public List<VbtFunctionBase> VbtLib { get; set; }

        public void Read(string igxl)
        {
            VbtLib = new List<VbtFunctionBase>();

            using (var zip = new ZipFile(igxl))
            {
                var entries = zip.Entries.ToList();
                foreach (var entry in entries)
                    GetVbtLib(entry);
            }
        }

        private void GetVbtLib(ZipEntry zipEntry)
        {
            if (zipEntry.FileName.StartsWith("VBT_", StringComparison.CurrentCultureIgnoreCase) ||
                zipEntry.FileName.StartsWith("LIB_", StringComparison.CurrentCultureIgnoreCase))
            {
                var stream = zipEntry.OpenReader();
                using (var sr = new StreamReader(stream))
                {
                    var extension = Path.GetExtension(zipEntry.FileName);
                    if (extension.ToLower() == ".bas")
                        VbtLib.AddRange(ReadBasFile(sr, zipEntry.FileName));
                }
            }
        }

        private List<VbtFunctionBase> ReadBasFile(StreamReader sr, string fileName)
        {
            var vbtFunctionBaseList = new List<VbtFunctionBase>();
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (!Regex.IsMatch(line, @"^\s*(Public\s)?Function.*\("))
                    continue;
                var functionName = Regex.Match(line, @"(Public\s)?Function\s(?<func>\w+)\(").Groups["func"].ToString();
                string paramStr;
                line = line.TrimEnd('_');
                if (Regex.IsMatch(line, @"\(.*\)"))
                {
                    paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                }
                else
                {
                    paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                    while ((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)"))
                    {
                        line = line.TrimEnd('_');
                        if (!Regex.IsMatch(line, @"\s*\'"))
                            paramStr += line;
                    }

                    if (line != null)
                        paramStr += Regex.Match(line, @"(?<str>.*)\)").Groups["str"].ToString();
                }

                var parameters = GetParameters(paramStr);
                var newVbt = new VbtFunctionBase(functionName);
                newVbt.FileName = fileName;

                for (var a = 0; a < parameters.Count; a++)
                    if (parameters[a].Name.ToLower() == "step_")
                    {
                        parameters.RemoveAt(a);
                        break;
                    }

                newVbt.Parameters = string.Join(",", parameters.Select(x => x.Name));
                newVbt.ParameterDefaults = string.Join(",", parameters.Select(x => x.Default));
                vbtFunctionBaseList.Add(newVbt);
            }

            sr.Close();
            return vbtFunctionBaseList;
        }

        private List<Parameter> GetParameters(string paramStr)
        {
            if (string.IsNullOrEmpty(paramStr))
                return new List<Parameter>();

            var parameters = new List<Parameter>();
            foreach (var str in paramStr.Split(','))
            {
                var parameter = new Parameter();
                var parameterName = Regex.Match(str, @"(?<param>\w+)\sAs\s", RegexOptions.IgnoreCase).Groups["param"]
                    .ToString();
                var parameterType = Regex.Match(paramStr, @"(?<param>\w+)\sAs\s(?<type>[^,]*)", RegexOptions.IgnoreCase)
                    .Groups["type"].ToString();
                parameter.Name = parameterName;
                if (parameterType.Contains("="))
                {
                    parameter.Type = parameterType.Replace(" ", "").Split('=')[0].Replace("\"", "");
                    parameter.Default = parameterType.Replace(" ", "").Split('=')[1].Replace("\"", "");
                }
                else
                {
                    parameter.Type = parameterType;
                    parameter.Default = "";
                }

                if (!(string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(parameterType)))
                    parameters.Add(parameter);
            }

            return parameters;
        }

        public VbtFunctionBase GetFunctionByName(string functionName)
        {
            var newVbt = new VbtFunctionBase();
            var resultVbt = VbtLib.Find(a => a.FunctionName.ToLower() == functionName.ToLower());
            if (resultVbt == null)
            {
                ErrorManager.AddError(EnumErrorType.MisVbtModule, EnumErrorLevel.Error, "", 0,
                    string.Format("The VBT function: {0} can not find in VBT library!", functionName));
                resultVbt = new VbtFunctionBase();
                resultVbt.FunctionName = functionName;
            }

            newVbt.Parameters = resultVbt.Parameters;
            newVbt.ParameterDefaults = resultVbt.ParameterDefaults;
            newVbt.FileName = resultVbt.FileName;
            newVbt.FunctionName = resultVbt.FunctionName;
            newVbt.Args = resultVbt.Args.ToList();
            return newVbt;
        }

        public void AddVbtFunction(VbtFunctionBase vbtFunction)
        {
            VbtLib.Add(vbtFunction);
        }

        public void AddVbtFunctionRange(List<VbtFunctionBase> vbtFunction)
        {
            VbtLib.AddRange(vbtFunction);
        }

        public static void CheckMissingParameter(string functionName, string paramName)
        {
            if (!MissingParams.Contains(functionName + "&" + paramName))
            {
                var errorMessage = "Missing Parameter in " + functionName + " : " + paramName;
                ErrorManager.AddError(EnumErrorType.MissingParameter, EnumErrorLevel.Error, "", 0,
                    errorMessage, paramName, functionName);
                MissingParams.Add(functionName + "&" + paramName);
            }
        }

        #region RF

        public const string RfFunc = "rffunctional_genericpgm";
        public const string RfCustom = "wi_custom";

        #endregion

        #region PMIC

        public const string PmicIdsVbtName = "IDS_OFF_TEST";
        public const string PmicLeakageVbtName = "dc_io_leakage";
        public const string PmicLeakageDcviVbtName = "dc_dcvi_leakage";
        public const string PmicLeakageDcvsVbtName = "dc_dcvs_leakage";

        #endregion
    }
}