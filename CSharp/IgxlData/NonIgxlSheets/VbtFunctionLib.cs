using System.Collections.Generic;
using System.Linq;
using AutomationCommon.EpplusErrorReport;
using IgxlData.VBT;

namespace IgxlData.NonIgxlSheets
{
    public class VbtFunctionLib
    {
        public const string FunctionalTUpdated = "functional_t_updated";
        public const string DvdcTrim = "wi_trimuniversalfunc";
        public const string LcdMeas = "measuniversalfunc";
        public const string VifName = "meas_freqvoltcurr_universal_func";
        public const string VirName = "meas_vir_io_universal_func";
        public const string VdiffFunc = "meas_vdiff_func";
        public const string FreqSynMeasFreqCurr = "freqsyn_measfreqcurr_func";

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

        public static List<List<string>> ParamMappingList = new List<List<string>>();
        public static Dictionary<string, int> GeneratedVbtFunctionDic = new Dictionary<string, int>();
        private static readonly List<string> MissingParams = new List<string>();
        public List<VbtFunctionBase> VbtLib { get; set; }

        public VbtFunctionLib()
        {
            VbtLib = new List<VbtFunctionBase>();
            GeneratedVbtFunctionDic.Clear();
        }

        public VbtFunctionBase GetFunctionByName(string functionName)
        {
            var newVbt = new VbtFunctionBase();
            var resultVbt = VbtLib.Find(a => a.FunctionName.ToLower() == functionName.ToLower());
            if (resultVbt == null)
            {
                EpplusErrorManager.AddError(HardIpErrorType.MisVbtModule, ErrorLevel.Error, "", 0, string.Format("The VBT function: {0} can not find in VBT library!", functionName));
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
                string errorMessage = "Missing Parameter in " + functionName + " : " + paramName;
                EpplusErrorManager.AddError(HardIpErrorType.MissingParameter.ToString(), ErrorLevel.Error, "", 0, errorMessage, paramName, functionName);
                MissingParams.Add(functionName + "&" + paramName);
            }
        }
    }
}
