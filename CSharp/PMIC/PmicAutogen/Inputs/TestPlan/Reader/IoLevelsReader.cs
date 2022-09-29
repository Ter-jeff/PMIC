using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.Utility;
using CommonReaderLib;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic.GenDc.DcInitial;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class IoLevelsItem
    {
        public string CharAppliedPins;
        public string Domain;
        public string Ioh;
        public string Iol;
        public bool IsSameDomain;
        public string Level;
        public string Vdd;
        public string Vih;
        public string Vil;
        public string Voh;
        public string Vol;

        public IoLevelsItem()
        {
            IsSameDomain = true;
            Level = "";
            Domain = "";
            Vdd = "";
            Vih = "";
            Vil = "";
            Voh = "";
            Vol = "";
            Ioh = "";
            Iol = "";
            CharAppliedPins = "";
        }

        public IoLevelsItem(string level)
        {
            IsSameDomain = true;
            Level = level;
            Domain = "";
            Vdd = "";
            Vih = "";
            Vil = "";
            Voh = "";
            Vol = "";
            Ioh = "";
            Iol = "";
            CharAppliedPins = "";
        }

        public int DomainIndex = -1;
    }

    public class IoLevelsRow
    {
        public string Domain;
        public bool IsBscanApplyPins;
        public bool IsGroupPin;
        public bool IsTheSameRow;

        #region Constructor

        public IoLevelsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Type = "";
            PinName = "";
            Fsdd = "";
            IoLevelDate = new List<IoLevelsItem>();
        }

        #endregion

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string Type { get; set; }
        public string PinName { get; set; }
        public string Fsdd { get; set; }
        public List<IoLevelsItem> IoLevelDate { get; set; }

        public GlobalSpec GetGlobalSpec(string vdd, string value, string domain, string level, bool isTheSameDomain,
            string type = "", PinSelector pinSelector = PinSelector.Nv, bool is2Nd = false)
        {
            double doubleValue;
            if (string.IsNullOrEmpty(value))
                value = "0";
            else if (!double.TryParse(value, out doubleValue))
                value = "";
            var glbSymbol = GetGlobalSpecName(vdd, domain, level, isTheSameDomain, type, pinSelector, is2Nd);

            if (LocalSpecs.VddRefInfoList.ContainsKey(domain))
            {
                double nv;
                double lv;
                double hv;
                double ulv;
                double uhv;
                double.TryParse(LocalSpecs.VddRefInfoList[domain].Nv, out nv);
                double.TryParse(LocalSpecs.VddRefInfoList[domain].Lv, out lv);
                double.TryParse(LocalSpecs.VddRefInfoList[domain].Hv, out hv);
                double.TryParse(LocalSpecs.VddRefInfoList[domain].ULv, out ulv);
                double.TryParse(LocalSpecs.VddRefInfoList[domain].UHv, out uhv);
                switch (pinSelector)
                {
                    case PinSelector.Nv:
                        if (type == "")
                            value = string.Format("_{0}",
                                SpecFormat.GenGlbSpecSymbol(LocalSpecs.VddRefInfoList[domain].WsBumpName));
                        break;
                    case PinSelector.Hv:
                        if (type == "")
                            value = string.Format("_{0}*_{1}",
                                SpecFormat.GenGlbSpecSymbol(LocalSpecs.VddRefInfoList[domain].WsBumpName),
                                SpecFormat.GenGlbPlus(LocalSpecs.VddRefInfoList[domain].WsBumpName));
                        else
                            value = nv == 0 ? "0" : (hv / nv).ToString(CultureInfo.InvariantCulture);
                        break;
                    case PinSelector.Lv:
                        if (type == "")
                            value = string.Format("_{0}*_{1}",
                                SpecFormat.GenGlbSpecSymbol(LocalSpecs.VddRefInfoList[domain].WsBumpName),
                                SpecFormat.GenGlbMinus(LocalSpecs.VddRefInfoList[domain].WsBumpName));
                        else
                            value = nv == 0 ? "0" : (lv / nv).ToString(CultureInfo.InvariantCulture);
                        break;
                    case PinSelector.Uhv:
                        if (type == "")
                            value = string.Format("_{0}*_{1}",
                                SpecFormat.GenGlbSpecSymbol(LocalSpecs.VddRefInfoList[domain].WsBumpName),
                                SpecFormat.GenGlbPlusUHv(LocalSpecs.VddRefInfoList[domain].WsBumpName));
                        else
                            value = nv == 0 ? "0" : (uhv / nv).ToString(CultureInfo.InvariantCulture);
                        break;
                    case PinSelector.Ulv:
                        if (type == "")
                            value = string.Format("_{0}*_{1}",
                                SpecFormat.GenGlbSpecSymbol(LocalSpecs.VddRefInfoList[domain].WsBumpName),
                                SpecFormat.GenGlbMinusULv(LocalSpecs.VddRefInfoList[domain].WsBumpName));
                        else
                            value = nv == 0 ? "0" : (ulv / nv).ToString(CultureInfo.InvariantCulture);
                        break;
                }
            }

            var globalSpecnew = new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(value));
            globalSpecnew.Comment = type;
            return globalSpecnew;
        }

        public string GetGlobalSpecName(string vdd, string domain, string level, bool isTheSameDomain, string type = "",
            PinSelector pinSelector = PinSelector.Nv, bool is2Nd = false)
        {
            domain = Combine.CombineByUnderLine(domain, type);
            var prefix = string.IsNullOrEmpty(vdd) ? "VIN_0v_" : "VIN_" + vdd.Replace(".", "p") + "v_";
            var pinName = isTheSameDomain ? domain : domain + (is2Nd ? "_2ndCondition" : "") + "_" + level;
            var glbSymbol = "";
            switch (pinSelector)
            {
                case PinSelector.Nv:
                    glbSymbol = prefix + SpecFormat.GenGlbSpecSymbol(pinName);
                    break;
                case PinSelector.Hv:
                    glbSymbol = prefix + SpecFormat.GenGlbPlus(pinName);
                    break;
                case PinSelector.Lv:
                    glbSymbol = prefix + SpecFormat.GenGlbMinus(pinName);
                    break;
                case PinSelector.Uhv:
                    glbSymbol = prefix + SpecFormat.GenGlbPlusUHv(pinName);
                    break;
                case PinSelector.Ulv:
                    glbSymbol = prefix + SpecFormat.GenGlbMinusULv(pinName);
                    break;
            }

            return glbSymbol;
        }


        public string GetBscanDomainGlobalSpecName(string vdd, string domain, string level, bool isTheSameDomain,
            string type = "", PinSelector pinSelector = PinSelector.Nv, bool is2Nd = false)
        {
            domain = Combine.CombineByUnderLine(domain, type);
            var xRows = StaticTestPlan.VddLevelsSheet.XRows;
            var xRow = xRows.Find(o => o.WsBumpName.Equals(domain, StringComparison.CurrentCultureIgnoreCase));
            if (xRow != null)
                vdd = xRow.Nv;
            var prefix = string.IsNullOrEmpty(vdd) ? "VIN_0v_" : "VIN_" + vdd.Replace(".", "p") + "v_";
            var pinName = isTheSameDomain ? domain : domain + (is2Nd ? "_2ndCondition" : "") + "_" + level;
            var glbSymbol = "";
            switch (pinSelector)
            {
                case PinSelector.Nv:
                    glbSymbol = prefix + SpecFormat.GenGlbSpecSymbol(pinName);
                    break;
                case PinSelector.Hv:
                    glbSymbol = prefix + SpecFormat.GenGlbHv(pinName);
                    break;
                case PinSelector.Lv:
                    glbSymbol = prefix + SpecFormat.GenGlbLv(pinName);
                    break;
                case PinSelector.Uhv:
                    glbSymbol = prefix + SpecFormat.GenGlbUHv(pinName);
                    break;
                case PinSelector.Ulv:
                    glbSymbol = prefix + SpecFormat.GenGlbULv(pinName);
                    break;
            }

            return glbSymbol;
        }

        private List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }

        public DcSpec GetDcSpecs(List<string> categoryList, string type)
        {
            var categories = IoLevelDate.Select(x => x.Level).ToList();
            var domain = IsGroupPin ? Domain : PinName;
            var dcSpec = new DcSpec(domain + "_" + type + "_VAR", GetSelectorList());
            foreach (var category in categoryList)
            {
                var orgCategory = category.Replace("_UltraVoltage", "");
                //if (!Category.Contains("_UltraVoltage") && orgCategory.StartsWith("BSCAN", StringComparison.CurrentCultureIgnoreCase))
                //{
                //    orgCategory = "BSCAN";
                //}
                //else if (Category.Contains("_UltraVoltage") && orgCategory.StartsWith("BSCAN", StringComparison.CurrentCultureIgnoreCase))
                //{
                //    orgCategory = "Common";
                //}
                if (orgCategory.StartsWith("BSCAN", StringComparison.CurrentCultureIgnoreCase)) orgCategory = "BSCAN";

                var ioLevelsItem =
                    IoLevelDate.Find(x => x.Level.Equals(orgCategory, StringComparison.OrdinalIgnoreCase)) ??
                    IoLevelDate.First();
                var vdd = ioLevelsItem.Vdd;
                var result = "";
                if (type.Equals("VIH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vih;
                else if (type.Equals("VIL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vil;
                else if (type.Equals("VOH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Voh;
                else if (type.Equals("VOL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vol;
                else if (type.Equals("IOH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Ioh;
                else if (type.Equals("IOL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Iol;

                var syntaxNv = Getsyntax(result, ioLevelsItem, categories, vdd, domain, orgCategory, type,
                    PinSelector.Nv);
                var syntaxHv = "0";
                var syntaxLv = "0";

                if (LocalSpecs.VddRefInfoList.ContainsKey(Domain))
                {
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = Getsyntax(result, ioLevelsItem, categories, vdd, domain, orgCategory, type,
                            PinSelector.Uhv);
                    else if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                             !LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = "0";
                    else
                        syntaxHv = Getsyntax(result, ioLevelsItem, categories, vdd, domain, orgCategory, type,
                            PinSelector.Hv);

                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = Getsyntax(result, ioLevelsItem, categories, vdd, domain, orgCategory, type,
                            PinSelector.Ulv);
                    else if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                             !LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = "0";
                    else
                        syntaxLv = Getsyntax(result, ioLevelsItem, categories, vdd, domain, orgCategory, type,
                            PinSelector.Lv);
                }
                else
                {
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        !LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = "0";
                    else
                        syntaxHv = syntaxNv;
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        !LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = "0";
                    else
                        syntaxLv = syntaxNv;
                }

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase))
                    syntaxNv = "0";

                dcSpec.AddCategory(GetCategoryItem(orgCategory, syntaxNv, syntaxHv, syntaxLv));
            }

            return dcSpec;
        }

        public DcSpec GetBscanApplyPinsDcSpecs(List<string> categoryList,
            Dictionary<string, Tuple<string, string>> dcSpecCategoryMapping, string type,
            List<KeyValuePair<string, List<string>>> typeValues2NdCondition)
        {
            var categories = IoLevelDate.Select(x => x.Level).ToList();
            var domain = IsGroupPin ? Domain : PinName;
            var dcSpec = new DcSpec(domain + "_BSCAN_ApplyPins" + "_" + type + "_VAR", GetSelectorList());


            foreach (var category in categoryList)
            {
                var orgCategory = category.Replace("_UltraVoltage", "");
                var bscanCategory = orgCategory;
                var is2NdCondition = false;
                if (typeValues2NdCondition.FindAll(o => o.Key.Equals(type, StringComparison.CurrentCultureIgnoreCase))
                        .Count > 0) is2NdCondition = true;

                if (category.StartsWith("BScan", StringComparison.CurrentCultureIgnoreCase))
                    bscanCategory = "Bscan";
                else
                    is2NdCondition = false;

                var ioLevelsItem =
                    IoLevelDate.Find(x => x.Level.Equals(orgCategory, StringComparison.OrdinalIgnoreCase)) ??
                    IoLevelDate.First();
                var vdd = ioLevelsItem.Vdd;
                var result = "";
                if (type.Equals("VIH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vih;
                else if (type.Equals("VIL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vil;
                else if (type.Equals("VOH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Voh;
                else if (type.Equals("VOL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Vol;
                else if (type.Equals("IOH", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Ioh;
                else if (type.Equals("IOL", StringComparison.CurrentCulture))
                    result = ioLevelsItem.Iol;

                var syntaxNv = GetBscansyntax(result, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                    bscanCategory, orgCategory, type, PinSelector.Nv, is2NdCondition);
                var syntaxHv = "0";
                var syntaxLv = "0";

                if (LocalSpecs.VddRefInfoList.ContainsKey(Domain) || dcSpecCategoryMapping.ContainsKey(orgCategory))
                {
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = GetBscansyntax(result, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                            bscanCategory, orgCategory, type, PinSelector.Uhv, is2NdCondition);
                    else if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                             !LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = "0";
                    else
                        syntaxHv = GetBscansyntax(result, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                            bscanCategory, orgCategory, type, PinSelector.Hv, is2NdCondition);

                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = GetBscansyntax(result, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                            bscanCategory, orgCategory, type, PinSelector.Ulv, is2NdCondition);
                    else if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                             !LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = "0";
                    else
                        syntaxLv = GetBscansyntax(result, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                            bscanCategory, orgCategory, type, PinSelector.Lv, is2NdCondition);
                }
                else
                {
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        !LocalSpecs.HasUltraVoltageUHv)
                        syntaxHv = "0";
                    else
                        syntaxHv = syntaxNv;
                    if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                        !LocalSpecs.HasUltraVoltageULv)
                        syntaxLv = "0";
                    else
                        syntaxLv = syntaxNv;
                }

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase))
                    syntaxNv = "0";
                dcSpec.AddCategory(GetCategoryItem(orgCategory, syntaxNv, syntaxHv, syntaxLv));
            }

            return dcSpec;
        }

        private string Getsyntax(string voltage, IoLevelsItem ioLevelsItem,
            List<string> categories, string vdd,
            string domain, string category, string type, PinSelector pinSelector, bool is2NdCondition = false)
        {
            var syntax = "";
            if (voltage.Contains("*"))
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(category, StringComparison.OrdinalIgnoreCase));
                var globalSpecVar = "_" + GetGlobalSpecName(vdd, domain, category, isTheSameDomain, "", pinSelector);

                if (is2NdCondition) isTheSameDomain = false;

                if (Regex.IsMatch(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase))
                {
                    var global = "_" + GetGlobalSpecName(vdd, domain, category, isTheSameDomain, type, PinSelector.Nv,
                        is2NdCondition) + "*";
                    var match = Regex.Match(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase);
                    var factor = match.Groups["value"] + "*";
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else if (Regex.IsMatch(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                             RegexOptions.IgnoreCase))
                {
                    var global = "*" + "_" + GetGlobalSpecName(vdd, domain, category, isTheSameDomain, type,
                        PinSelector.Nv, is2NdCondition);
                    var match = Regex.Match(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                        RegexOptions.IgnoreCase);
                    var factor = "*" + match.Groups["value"];
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else
                {
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                }
            }
            else
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(category, StringComparison.OrdinalIgnoreCase));
                var globalSpecVar = "_" + GetGlobalSpecName(vdd, domain, category, isTheSameDomain, "", pinSelector);

                if (is2NdCondition) isTheSameDomain = false;
                var global = "_" + GetGlobalSpecName(vdd, domain, category, isTheSameDomain, type, PinSelector.Nv,
                    is2NdCondition);
                if (voltage.Contains(ioLevelsItem.Domain))
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                else
                    syntax = global;
            }

            return syntax;
        }

        private string GetBscansyntax(string voltage, IoLevelsItem ioLevelsItem,
            List<string> categories, Dictionary<string, Tuple<string, string>> dcSpecCategoryMapping, string vdd,
            string domain, string bscanCategory, string orgCategory, string type, PinSelector pinSelector,
            bool is2NdCondition = false)
        {
            if (dcSpecCategoryMapping.ContainsKey(orgCategory))
                return GetBscanDomainSyntax(voltage, ioLevelsItem, categories, dcSpecCategoryMapping, vdd, domain,
                    bscanCategory, orgCategory, type, pinSelector, is2NdCondition);

            var syntax = "";
            if (voltage.Contains("*"))
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(bscanCategory, StringComparison.OrdinalIgnoreCase));
                var globalSpecVar =
                    "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, "", pinSelector);

                if (is2NdCondition) isTheSameDomain = false;

                if (Regex.IsMatch(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase))
                {
                    var global = "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type,
                        PinSelector.Nv, is2NdCondition) + "*";
                    var match = Regex.Match(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase);
                    var factor = match.Groups["value"] + "*";
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else if (Regex.IsMatch(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                             RegexOptions.IgnoreCase))
                {
                    var global = "*" + "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type,
                        PinSelector.Nv, is2NdCondition);
                    var match = Regex.Match(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                        RegexOptions.IgnoreCase);
                    var factor = "*" + match.Groups["value"];
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else
                {
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                }
            }
            else
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(bscanCategory, StringComparison.OrdinalIgnoreCase));
                var globalSpecVar =
                    "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, "", pinSelector);

                if (is2NdCondition) isTheSameDomain = false;
                var global = "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type, PinSelector.Nv,
                    is2NdCondition);
                if (voltage.Contains(ioLevelsItem.Domain))
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                else
                    syntax = global;
            }

            return syntax;
        }

        private string GetBscanDomainSyntax(string voltage, IoLevelsItem ioLevelsItem,
            List<string> categories, Dictionary<string, Tuple<string, string>> dcSpecCategoryMapping, string vdd,
            string domain, string bscanCategory, string orgCategory, string type, PinSelector pinSelector,
            bool is2NdCondition = false)
        {
            var syntax = "";
            if (voltage.Contains("*"))
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(bscanCategory, StringComparison.OrdinalIgnoreCase));
                //for example: _VIN_1p2v_LDO9_GLB_LV , used BSCAN_CHAR domain
                var globalSpecVar = "_" + GetBscanDomainGlobalSpecName(vdd, dcSpecCategoryMapping[orgCategory].Item1,
                    bscanCategory, true, "", pinSelector);

                if (is2NdCondition) isTheSameDomain = false;

                if (Regex.IsMatch(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase))
                {
                    var global = "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type,
                        PinSelector.Nv, is2NdCondition) + "*";
                    var match = Regex.Match(voltage, @"(?<value>\d+([.]\d+)?)\*" + ioLevelsItem.Domain,
                        RegexOptions.IgnoreCase);
                    var factor = match.Groups["value"] + "*";
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else if (Regex.IsMatch(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                             RegexOptions.IgnoreCase))
                {
                    var global = "*" + "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type,
                        PinSelector.Nv, is2NdCondition);
                    var match = Regex.Match(voltage, ioLevelsItem.Domain + @"\*(?<value>\d+([.]\d+)?)",
                        RegexOptions.IgnoreCase);
                    var factor = "*" + match.Groups["value"];
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar).Replace(factor, global);
                }
                else
                {
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                }
            }
            else
            {
                var isTheSameDomain = IsTheSameRow ||
                                      !categories.Exists(
                                          x => x.Equals(bscanCategory, StringComparison.OrdinalIgnoreCase));
                var globalSpecVar = "_" + GetBscanDomainGlobalSpecName(vdd, dcSpecCategoryMapping[orgCategory].Item1,
                    bscanCategory, true, type, pinSelector);

                if (is2NdCondition) isTheSameDomain = false;
                var global = "_" + GetGlobalSpecName(vdd, domain, bscanCategory, isTheSameDomain, type, PinSelector.Nv,
                    is2NdCondition);
                if (type == "IOH" || type == "IOL")
                    global = global.Replace(bscanCategory + "_GLB",
                        dcSpecCategoryMapping[orgCategory].Item2 + "_" + bscanCategory + "_GLB");
                if (voltage.Contains(ioLevelsItem.Domain))
                    syntax = voltage.Replace(ioLevelsItem.Domain, globalSpecVar);
                else
                    syntax = global;
            }

            return syntax;
        }


        private CategoryInSpec GetCategoryItem(string category, string glbSymbolTyp, string glbSymbolMax,
            string glbSymbolMin)
        {
            var categoryItem = new CategoryInSpec(category);
            categoryItem.Typ = GetBaseSymbol(glbSymbolTyp);
            //categoryItem.Max = getBaseSymbol(glbSymbolMax) + "*_IO_Pins_GLB_Plus";
            //categoryItem.Min = getBaseSymbol(glbSymbolMin) + "*_IO_Pins_GLB_Minus";
            categoryItem.Max = GetBaseSymbol(glbSymbolMax);
            categoryItem.Min = GetBaseSymbol(glbSymbolMin);
            return categoryItem;
        }

        private string GetBaseSymbol(string glbSymbol)
        {
            return IsFormula(glbSymbol) ? "=(" + glbSymbol + ")" : "=" + glbSymbol;
        }

        private bool IsFormula(string symbol)
        {
            if (symbol.Contains('+'))
                return true;
            if (symbol.Contains('-'))
                return true;
            if (symbol.Contains('*'))
                return true;
            if (symbol.Contains('/'))
                return true;

            return false;
        }

        public DcSpec GetVtDcSpecs(List<string> categoryList)
        {
            var domain = IsGroupPin ? Domain : PinName;
            var dcSpec = new DcSpec(domain + "_VT_VAR", GetSelectorList());
            foreach (var category in categoryList)
            {
                var vtSyntax = "=(_" + domain + "_VOH_VAR+_" + domain + "_VOL_VAR)/2";
                var categoryInSpec = new CategoryInSpec(category);
                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase))
                    categoryInSpec.Typ = "0";
                else
                    categoryInSpec.Typ = vtSyntax;

                //categoryInSpec.Max = vtSyntax + "*_IO_Pins_GLB_Plus";
                //categoryInSpec.Min = vtSyntax + "*_IO_Pins_GLB_Minus";

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                    !LocalSpecs.HasUltraVoltageUHv)
                    categoryInSpec.Max = "0";
                else
                    categoryInSpec.Max = vtSyntax;

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                    !LocalSpecs.HasUltraVoltageULv)
                    categoryInSpec.Min = "0";
                else
                    categoryInSpec.Min = vtSyntax;
                dcSpec.AddCategory(categoryInSpec);
            }

            return dcSpec;
        }

        public DcSpec GetBscanApplyPinsVtDcSpecs(List<string> categoryList)
        {
            var domain = IsGroupPin ? Domain : PinName;
            domain = domain + "_BSCAN_ApplyPins";
            var dcSpec = new DcSpec(domain + "_VT_VAR", GetSelectorList());
            foreach (var category in categoryList)
            {
                var vtSyntax = "=(_" + domain + "_VOH_VAR+_" + domain + "_VOL_VAR)/2";
                var categoryInSpec = new CategoryInSpec(category);
                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase))
                    categoryInSpec.Typ = "0";
                else
                    categoryInSpec.Typ = vtSyntax;
                //categoryInSpec.Max = vtSyntax + "*_IO_Pins_GLB_Plus";
                //categoryInSpec.Min = vtSyntax + "*_IO_Pins_GLB_Minus";

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                    !LocalSpecs.HasUltraVoltageUHv)
                    categoryInSpec.Max = "0";
                else
                    categoryInSpec.Max = vtSyntax;

                if (category.EndsWith("_UltraVoltage", StringComparison.OrdinalIgnoreCase) &&
                    !LocalSpecs.HasUltraVoltageULv)
                    categoryInSpec.Min = "0";
                else
                    categoryInSpec.Min = vtSyntax;
                dcSpec.AddCategory(categoryInSpec);
            }

            return dcSpec;
        }

        public List<DcSpec> GenDcSpecs(List<string> categoryList)
        {
            var dcSpecsList = new List<DcSpec>();
            dcSpecsList.Add(GetDcSpecs(categoryList, "VIH"));
            dcSpecsList.Add(GetDcSpecs(categoryList, "VIL"));
            dcSpecsList.Add(GetDcSpecs(categoryList, "VOH"));
            dcSpecsList.Add(GetDcSpecs(categoryList, "VOL"));
            dcSpecsList.Add(GetDcSpecs(categoryList, "IOH"));
            dcSpecsList.Add(GetDcSpecs(categoryList, "IOL"));
            dcSpecsList.Add(GetVtDcSpecs(categoryList));
            return dcSpecsList;
        }
    }

    public class IoLevelsSheet : MySheet
    {
        #region Constructor

        public IoLevelsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<IoLevelsRow>();
            _domainTypeValues = new Dictionary<string, Dictionary<string, List<string>>>();
            DcSpecCategoryMapping = new Dictionary<string, Tuple<string, string>>();
        }

        #endregion

        public void UpdateDcCategory(List<DcCategory> categories, string block)
        {
            var levels = Rows.SelectMany(x => x.IoLevelDate.Select(y => y.Level)).Distinct().ToList();
            foreach (var level in levels)
                if (!categories.Exists(x => x.CategoryName.Equals(level, StringComparison.OrdinalIgnoreCase)))
                    categories.Add(new DcCategory(level, block, level, DcCategoryType.Pmic));
                else
                    categories[
                            categories.FindIndex(x => x.CategoryName.Equals(level, StringComparison.OrdinalIgnoreCase))]
                        .Type = DcCategoryType.Pmic;

            if (LocalSpecs.HasUltraVoltage)
            {
                categories.Add(new DcCategory("Common_UltraVoltage", block, "Common_UltraVoltage",
                    DcCategoryType.Pmic));
                LocalSpecs.UltraVoltageCategory.Add("Common", "Common_UltraVoltage");
                for (var i = 1; i < Rows[0].IoLevelDate.Count; i++)
                {
                    var level = Rows[i].IoLevelDate[i].Level;
                    for (var j = 0; j < Rows.Count; j++)
                        if (Rows[j].IoLevelDate[i].Vdd != Rows[j].IoLevelDate[0].Vdd ||
                            Rows[j].IoLevelDate[i].Vih != Rows[j].IoLevelDate[0].Vih ||
                            Rows[j].IoLevelDate[i].Vil != Rows[j].IoLevelDate[0].Vil ||
                            Rows[j].IoLevelDate[i].Voh != Rows[j].IoLevelDate[0].Voh ||
                            Rows[j].IoLevelDate[i].Vol != Rows[j].IoLevelDate[0].Vol)
                        {
                            if (categories.Exists(x =>
                                    x.CategoryName.Equals(level, StringComparison.OrdinalIgnoreCase)))
                                level = categories.Find(x =>
                                    x.CategoryName.Equals(level, StringComparison.OrdinalIgnoreCase)).CategoryName;
                            LocalSpecs.UltraVoltageCategory.Add(level, level + "_UltraVoltage");
                            categories.Add(new DcCategory(level + "_UltraVoltage", block, level + "_UltraVoltage",
                                DcCategoryType.Pmic));
                            if (level.Equals("BScan", StringComparison.CurrentCultureIgnoreCase))
                                categories.AddRange(GenBScanUltraVoltageDcCategory(""));
                            break;
                        }
                }
            }
        }

        public List<DcCategory> GenBScanDcCategory(string block)
        {
            var bscanDcCategory = new List<DcCategory>();
            if (StaticTestPlan.BscanCharSheet == null)
                return bscanDcCategory;

            var domainCurrentMap = StaticTestPlan.BscanCharSheet.GetDomainCurrentMapping();
            var currents = StaticTestPlan.BscanCharSheet.GetDomainCurrents();

            foreach (var current in currents)
                foreach (var domainCurrentItem in domainCurrentMap)
                    if (domainCurrentItem.Value.Contains(current))
                    {
                        var categoryName = string.Format("BSCAN_{0}_{1}mA", domainCurrentItem.Key,
                            current.ToString().Replace('.', 'p'));
                        bscanDcCategory.Add(new DcCategory(categoryName, block, categoryName, DcCategoryType.Pmic));
                        if (!DcSpecCategoryMapping.ContainsKey(categoryName))
                            DcSpecCategoryMapping.Add(categoryName,
                                Tuple.Create(domainCurrentItem.Key, current.ToString().Replace('.', 'p') + "m"));
                    }

            return bscanDcCategory;
        }

        public List<DcCategory> GenBScanUltraVoltageDcCategory(string block)
        {
            var bscanUvDcCategory = new List<DcCategory>();
            if (StaticTestPlan.BscanCharSheet == null)
                return bscanUvDcCategory;

            var domainCurrentMap = StaticTestPlan.BscanCharSheet.GetDomainCurrentMapping();
            var currents = StaticTestPlan.BscanCharSheet.GetDomainCurrents();

            foreach (var current in currents)
                foreach (var domainCurrentItem in domainCurrentMap)
                    if (domainCurrentItem.Value.Contains(current))
                    {
                        var categoryName = string.Format("BSCAN_{0}_{1}mA_UltraVoltage", domainCurrentItem.Key,
                            current.ToString().Replace('.', 'p'));
                        //LocalSpecs.UltraVoltageCategory.Add(categoryName, categoryName);
                        bscanUvDcCategory.Add(new DcCategory(categoryName, block, categoryName, DcCategoryType.Pmic));
                    }

            return bscanUvDcCategory;
        }

        public List<GlobalSpec> GenGlbSymbol()
        {
            var globalSpecs = new List<GlobalSpec>();
            GenBscanDomainTypeValues();

            foreach (var row in Rows)
            {
                //Default
                var domain = row.IsGroupPin ? row.Domain : row.PinName;
                {
                    var data = row.IoLevelDate.First();
                    var level = data.Level;
                    var vdd = data.Vdd;
                    GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, true);

                    if (LocalSpecs.VddRefInfoList.ContainsKey(domain))
                    {
                        GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, true, PinSelector.Hv, true);
                        GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, true, PinSelector.Lv, true);
                        if (LocalSpecs.HasUltraVoltageUHv)
                            GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, true, PinSelector.Uhv, true);
                        if (LocalSpecs.HasUltraVoltageULv)
                            GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, true, PinSelector.Ulv, true);
                    }
                }

                if (!row.IsTheSameRow)
                    foreach (var data in row.IoLevelDate)
                    {
                        var level = data.Level;
                        var vdd = data.Vdd;
                        GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, false);
                        if (LocalSpecs.VddRefInfoList.ContainsKey(domain))
                        {
                            GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, false, PinSelector.Hv, true);
                            GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, false, PinSelector.Lv, true);
                            if (LocalSpecs.HasUltraVoltageUHv)
                                GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, false, PinSelector.Uhv,
                                    true);
                            if (LocalSpecs.HasUltraVoltageULv)
                                GetGlobalSpec(row, vdd, domain, level, ref globalSpecs, data, false, PinSelector.Ulv,
                                    true);
                        }
                    }
            }

            return globalSpecs;
        }

        private void GetGlobalSpec(IoLevelsRow row, string vdd, string domain, string level,
            ref List<GlobalSpec> globalSpecs, IoLevelsItem data, bool isTheSameDomain,
            PinSelector pinSelector = PinSelector.Nv, bool isRefGlobalSpec = false)
        {
            var globalSpec = row.GetGlobalSpec(vdd, vdd, domain, level, isTheSameDomain, "", pinSelector);
            if (!globalSpecs.Any(x => x.Symbol.Equals(globalSpec.Symbol, StringComparison.OrdinalIgnoreCase)))
                globalSpecs.Add(globalSpec);

            if (!isRefGlobalSpec)
            {
                var vih = GetFactor(data.Vih, data.Domain);
                var globalSpecVih = row.GetGlobalSpec(vdd, vih, domain, level, isTheSameDomain, "VIH", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecVih.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecVih);

                var vil = GetFactor(data.Vil, data.Domain);
                var globalSpecVil = row.GetGlobalSpec(vdd, vil, domain, level, isTheSameDomain, "VIL", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecVil.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecVil);

                var voh = GetFactor(data.Voh, data.Domain);
                var globalSpecVoh = row.GetGlobalSpec(vdd, voh, domain, level, isTheSameDomain, "VOH", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecVoh.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecVoh);

                var vol = GetFactor(data.Vol, data.Domain);
                var globalSpecVol = row.GetGlobalSpec(vdd, vol, domain, level, isTheSameDomain, "VOL", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecVol.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecVol);

                var ioh = GetFactor(data.Ioh, data.Domain);
                var globalSpecIoh = row.GetGlobalSpec(vdd, ioh, domain, level, isTheSameDomain, "IOH", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecIoh.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecIoh);

                var iol = GetFactor(data.Iol, data.Domain);
                var globalSpecIol = row.GetGlobalSpec(vdd, iol, domain, level, isTheSameDomain, "IOL", pinSelector);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecIol.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecIol);

                if (level.Equals("BSCAN", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (_domainTypeValues.ContainsKey(domain))
                        GenBscan2NdGolbalSpecs(globalSpecs, row, vdd, data.Vih, data.Vil, data.Voh, data.Vol, data.Ioh,
                            data.Iol, domain, level, isTheSameDomain, pinSelector);

                    // gen BSCAN ioh iol specs by BSCAN_CHAR
                    var bscanGlobalSpecs = GetBscanIohAndIolGlobalSpecs(row, vdd, ioh, iol, domain, level,
                        isTheSameDomain, pinSelector);
                    foreach (var bscanGlobalSpec in bscanGlobalSpecs)
                        if (!globalSpecs.Any(x =>
                                x.Symbol.Equals(bscanGlobalSpec.Symbol, StringComparison.OrdinalIgnoreCase)))
                            globalSpecs.Add(bscanGlobalSpec);
                }
            }
        }

        private string GetFactor(string value, string domain)
        {
            double outValue = 0;
            if (double.TryParse(value, out outValue))
                return outValue.ToString();

            if (!value.Contains("*")) return "";
            var regexPattern = @"(?<value>[+|-]?\d+([.]\d+)?)";
            //0.7 * LDO9
            if (Regex.IsMatch(value, regexPattern + @"\*" + domain, RegexOptions.IgnoreCase))
            {
                var match = Regex.Match(value, regexPattern + @"\*" + domain, RegexOptions.IgnoreCase);
                var factor = match.Groups["value"].ToString();
                return factor;
            }

            //LDO9* 0.7
            if (Regex.IsMatch(value, domain + @"\*" + regexPattern, RegexOptions.IgnoreCase))
            {
                var match = Regex.Match(value, domain + @"\*" + regexPattern, RegexOptions.IgnoreCase);
                var factor = match.Groups["value"].ToString();
                return factor;
            }

            return "";
        }

        public List<DcSpec> GenDcSpecForIoPins(List<string> categoryList)
        {
            var dcSpecsList = new List<DcSpec>();
            foreach (var row in Rows) dcSpecsList.AddRange(row.GenDcSpecs(categoryList));
            GenBscanApplyPinsDcSpecs(categoryList, dcSpecsList);
            return dcSpecsList;
        }

        private List<GlobalSpec> GetBscanIohAndIolGlobalSpecs(IoLevelsRow row, string vdd, string iohValue,
            string iolValue, string domain, string level, bool isTheSameDomain,
            PinSelector pinSelector = PinSelector.Nv)
        {
            var bscanGlbSpec = new List<GlobalSpec>();

            if (StaticTestPlan.BscanCharSheet == null)
                return bscanGlbSpec;

            var currentKindList = StaticTestPlan.BscanCharSheet.GetDomainCurrents();

            foreach (var current in currentKindList)
            {
                var iohType = string.Format("IOH_{0}m", current.ToString().Replace('.', 'p'));
                var iohGlbSymbol = row.GetGlobalSpecName(vdd, domain, level, isTheSameDomain, iohType, pinSelector);

                bscanGlbSpec.Add(new GlobalSpec(iohGlbSymbol,
                    SpecFormat.GenSpecValueSingleValue((0 - current / 1000).ToString()), "", "Come from BSCAN_CHAR"));

                var iolType = string.Format("IOL_{0}m", current.ToString().Replace('.', 'p'));
                var iolGlbSymbol = row.GetGlobalSpecName(vdd, domain, level, isTheSameDomain, iolType, pinSelector);
                bscanGlbSpec.Add(new GlobalSpec(iolGlbSymbol,
                    SpecFormat.GenSpecValueSingleValue((current / 1000).ToString()), "", "Come from BSCAN_CHAR"));
            }

            return bscanGlbSpec;
        }

        private void GenBscanDomainTypeValues()
        {
            _domainTypeValues.Clear();
            foreach (var row in Rows)
            {
                if (!_domainTypeValues.ContainsKey(row.Domain))
                    _domainTypeValues.Add(row.Domain, new Dictionary<string, List<string>>());

                var typeValues = _domainTypeValues[row.Domain];
                var bscanItem = row.IoLevelDate
                    .Where(o => o.Level.Equals("BSCAN", StringComparison.CurrentCultureIgnoreCase)).First();

                //VIH
                if (typeValues.ContainsKey("VIH") && !typeValues["VIH"].Contains(bscanItem.Vih))
                    typeValues["VIH"].Add(bscanItem.Vih);
                else if (!typeValues.ContainsKey("VIH"))
                    typeValues.Add("VIH", new List<string> { bscanItem.Vih });
                //VIL
                if (typeValues.ContainsKey("VIL") && !typeValues["VIL"].Contains(bscanItem.Vil))
                    typeValues["VIL"].Add(bscanItem.Vil);
                else if (!typeValues.ContainsKey("VIL"))
                    typeValues.Add("VIL", new List<string> { bscanItem.Vil });
                //VOH
                if (typeValues.ContainsKey("VOH") && !typeValues["VOH"].Contains(bscanItem.Voh))
                    typeValues["VOH"].Add(bscanItem.Voh);
                else if (!typeValues.ContainsKey("VOH"))
                    typeValues.Add("VOH", new List<string> { bscanItem.Voh });
                //VOL
                if (typeValues.ContainsKey("VOL") && !typeValues["VOL"].Contains(bscanItem.Vol))
                    typeValues["VOL"].Add(bscanItem.Vol);
                else if (!typeValues.ContainsKey("VOL"))
                    typeValues.Add("VOL", new List<string> { bscanItem.Vol });
                //IOH
                if (typeValues.ContainsKey("IOH") && !typeValues["IOH"].Contains(bscanItem.Ioh))
                    typeValues["IOH"].Add(bscanItem.Ioh);
                else if (!typeValues.ContainsKey("IOH"))
                    typeValues.Add("IOH", new List<string> { bscanItem.Ioh });
                //IOL
                if (typeValues.ContainsKey("IOL") && !typeValues["IOL"].Contains(bscanItem.Iol))
                    typeValues["IOL"].Add(bscanItem.Iol);
                else if (!typeValues.ContainsKey("IOL"))
                    typeValues.Add("IOL", new List<string> { bscanItem.Iol });
            }
        }

        private void GenBscan2NdGolbalSpecs(List<GlobalSpec> globalSpecs, IoLevelsRow row, string vdd, string vih,
            string vil, string voh, string vol, string ioh, string iol, string domain, string level,
            bool isTheSameDomain,
            PinSelector pinSelector = PinSelector.Nv)
        {
            var typeValues = _domainTypeValues[domain];

            //VIH
            GetBscan2NdGolbalSpec("VIH", typeValues, globalSpecs, row, vdd, vih, domain, level, isTheSameDomain,
                pinSelector);
            //VIL
            GetBscan2NdGolbalSpec("VIL", typeValues, globalSpecs, row, vdd, vil, domain, level, isTheSameDomain,
                pinSelector);
            //VOH
            GetBscan2NdGolbalSpec("VOH", typeValues, globalSpecs, row, vdd, voh, domain, level, isTheSameDomain,
                pinSelector);
            //VOL
            GetBscan2NdGolbalSpec("VOL", typeValues, globalSpecs, row, vdd, vol, domain, level, isTheSameDomain,
                pinSelector);
            //IOH
            GetBscan2NdGolbalSpec("IOH", typeValues, globalSpecs, row, vdd, ioh, domain, level, isTheSameDomain,
                pinSelector);
            //IOL
            GetBscan2NdGolbalSpec("IOL", typeValues, globalSpecs, row, vdd, iol, domain, level, isTheSameDomain,
                pinSelector);
        }

        private void GetBscan2NdGolbalSpec(string type, Dictionary<string, List<string>> typeValues,
            List<GlobalSpec> globalSpecs, IoLevelsRow row, string vdd, string value, string domain, string level,
            bool isTheSameDomain,
            PinSelector pinSelector = PinSelector.Nv)
        {
            if (!typeValues.ContainsKey(type)) return;

            if (typeValues[type].Count > 1)
            {
                var other = typeValues[type].Find(o => !o.Equals(value));
                var otherValue = GetFactor(other, domain);
                var globalSpecVih2Nd = row.GetGlobalSpec(vdd, otherValue, domain, level, isTheSameDomain, type,
                    pinSelector, true);
                if (!globalSpecs.Any(x =>
                        x.Symbol.Equals(globalSpecVih2Nd.Symbol, StringComparison.OrdinalIgnoreCase)))
                    globalSpecs.Add(globalSpecVih2Nd);
            }
        }

        public List<PinGroup> GenBscanApplyPinGroups()
        {
            var groups = new List<PinGroup>();
            var applyPinRows = new List<IoLevelsRow>();
            foreach (var row in Rows)
            {
                var bscanDataItem = row.IoLevelDate
                    .Where(o => o.Level.Equals("BSCAN", StringComparison.CurrentCultureIgnoreCase)).Select(o => o)
                    .FirstOrDefault();
                if (bscanDataItem == null) continue;

                if (bscanDataItem.CharAppliedPins.Equals("v", StringComparison.CurrentCultureIgnoreCase))
                    applyPinRows.Add(row);
            }

            var group = applyPinRows.GroupBy(x => x.Domain).ToList();
            foreach (var item in group)
            {
                var rows = item.ToList();
                if (string.IsNullOrEmpty(rows[0].Domain)) continue;
                var pinGroup = new PinGroup(rows[0].Domain + "_BSCAN_ApplyPins");
                var counter = 0;
                foreach (var row in rows.OrderBy(o => o.PinName))
                {
                    var comment = string.Empty;
                    if (counter == 0)
                    {
                        comment = "Depend on CHAR_Applied_Pins column";
                        counter++;
                    }
                    else if (counter == 1)
                    {
                        comment = "Remove no \"v\" pin";
                        counter++;
                    }

                    var newPin = new Pin(row.PinName, PinMapConst.TypeIo, comment);
                    pinGroup.AddPin(newPin);
                }

                groups.Add(pinGroup);
            }

            return groups;
        }

        private void GenBscanApplyPinsDcSpecs(List<string> categoryList, List<DcSpec> dcSpecsList)
        {
            var bscanApplyPinsDomains = Rows.Where(o => o.IsBscanApplyPins).Select(o => o.Domain).Distinct().ToList();
            if (!bscanApplyPinsDomains.Any())
                return;

            var domainTypeValues = GetBscanDomainTypeValues();

            foreach (var bscanDomain in bscanApplyPinsDomains)
            {
                var domainRow = Rows.Find(o =>
                    o.Domain.Equals(bscanDomain, StringComparison.CurrentCultureIgnoreCase) && o.IsGroupPin);
                if (domainRow == null)
                    continue;

                var typeValues2NdCondition = new List<KeyValuePair<string, List<string>>>();
                if (domainTypeValues.ContainsKey(bscanDomain))
                {
                    var typeValues = domainTypeValues[bscanDomain];
                    typeValues2NdCondition = typeValues.Where(o => o.Value.Count > 1).Select(o => o).ToList();
                    //If domain's Char_Apply_Pins has 'v' value, it will gen Applypins symbols.
                    //if (typeValues_2ndCondition.Count == 0)
                    //    continue;
                }
                else
                {
                    continue;
                }


                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "VIH",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "VIL",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "VOH",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "VOL",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "IOH",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsDcSpecs(categoryList, DcSpecCategoryMapping, "IOL",
                    typeValues2NdCondition));
                dcSpecsList.Add(domainRow.GetBscanApplyPinsVtDcSpecs(categoryList));
            }
        }

        public Dictionary<string, Dictionary<string, List<string>>> GetBscanDomainTypeValues()
        {
            if (_domainTypeValues == null || _domainTypeValues.Count == 0)
                GenBscanDomainTypeValues();

            return _domainTypeValues;
        }

        internal void CheckSheet()
        {
            foreach (var row in Rows)
            {
                var firstDomain = row.IoLevelDate.First().Domain;
                var firstLevel = row.IoLevelDate.First().Level;
                for (int i = 1; i < row.IoLevelDate.Count; i++)
                {
                    var ioLevelDate = row.IoLevelDate.ElementAt(i);
                    var domain = ioLevelDate.Domain;
                    var level = ioLevelDate.Level;
                    if (!firstDomain.Equals(domain, StringComparison.CurrentCulture))
                        AddError(EnumErrorType.FormatError, EnumErrorLevel.Warning, SheetName, row.RowNum, ioLevelDate.DomainIndex,
                           string.Format("Please comfirm domain \"{0}\" in \"{1}\" is different than \"{2}\" in \"{3}\" !!!",
                            domain, level, firstDomain, firstLevel));
                }
            }
        }

        #region Property

        //rows which don't contain floating domain.
        public List<IoLevelsRow> Rows { get; set; }

        //total rows which contain floating domain.
        public List<IoLevelsRow> TotalRows { get; set; }

        //key:demoin  value: type and value
        // VDDIO_1V8:
        //        VOH: 0.75*VDDIO_1V8, 0.77*VDDIO_1V8
        //        VOL: 0.25*VDDIO_1V8, 0.33*VDDIO_1V8
        private readonly Dictionary<string, Dictionary<string, List<string>>> _domainTypeValues;

        //key:category value: (domain,current)
        public Dictionary<string, Tuple<string, string>> DcSpecCategoryMapping;
        public int TypeIndex = -1;
        public int PinNameIndex = -1;
        public int FsddIndex = -1;

        #endregion
    }

    public class IoLevelsReader
    {
        private const string HeaderType = "Type";
        private const string HeaderPinName = "PinName";
        private const string HeaderFsdd = "FS/DD";
        private const string HeaderDomain = "Domain";
        private const string HeaderVdd = "VDD";
        private const string HeaderVih = "VIH";
        private const string HeaderVil = "VIL";
        private const string HeaderVoh = "VOH";
        private const string HeaderVol = "VOL";
        private const string HeaderIoh = "IOH";
        private const string HeaderIol = "IOL";
        private const string HeaderCharApplyPins = "CHAR_Applied_Pins";
        private readonly string _floatingString = "FLOATING";
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _fsddIndex = -1;
        private IoLevelsSheet _iOLevelsSheet;
        private int _pinNameIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _typeIndex = -1;

        public IoLevelsSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _iOLevelsSheet = new IoLevelsSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _iOLevelsSheet = ReadSheetData();

            //ignore rows which domain is floating.
            var rows = new List<IoLevelsRow>();
            rows.AddRange(_iOLevelsSheet.Rows.Where(y =>
                !y.Domain.Equals(_floatingString, StringComparison.CurrentCultureIgnoreCase) && y.IsGroupPin));
            rows.AddRange(_iOLevelsSheet.Rows.Where(x =>
                !x.Domain.Equals(_floatingString, StringComparison.CurrentCultureIgnoreCase) && !x.IsGroupPin));

            var totalRows = new List<IoLevelsRow>();
            totalRows.AddRange(_iOLevelsSheet.Rows.Where(y => y.IsGroupPin));
            totalRows.AddRange(_iOLevelsSheet.Rows.Where(x => !x.IsGroupPin));

            _iOLevelsSheet.Rows = rows;
            _iOLevelsSheet.TotalRows = totalRows;

            _iOLevelsSheet.CheckSheet();
            ErrorManager.AddErrors(_iOLevelsSheet.Errors);

            return _iOLevelsSheet;
        }

        private void FillEmptyCell(IoLevelsSheet ioLevelsSheet)
        {
            foreach (var row in ioLevelsSheet.Rows)
                foreach (var item in row.IoLevelDate)
                {
                    var firstRow = _iOLevelsSheet.Rows.SelectMany(y => y.IoLevelDate).ToList().Find(x =>
                        x.Level == item.Level && x.Domain == item.Domain && !string.IsNullOrEmpty(x.Vdd));
                    if (string.IsNullOrEmpty(item.Vdd) && firstRow != null)
                    {
                        item.Vdd = firstRow.Vdd;
                        item.Vih = firstRow.Vih;
                        item.Vil = firstRow.Vil;
                        item.Voh = firstRow.Voh;
                        item.Vol = firstRow.Vol;
                        item.Ioh = firstRow.Ioh;
                        item.Iol = firstRow.Iol;
                        item.CharAppliedPins = firstRow.CharAppliedPins;
                    }
                }
        }

        private void GetTheSameDomain(IoLevelsSheet ioLevelsSheet)
        {
            foreach (var row in ioLevelsSheet.Rows)
            {
                row.IsTheSameRow = true;
                row.IsGroupPin = true;
                foreach (var data in row.IoLevelDate)
                {
                    if (!data.Domain.Equals(row.IoLevelDate.First().Domain, StringComparison.CurrentCulture))
                        data.IsSameDomain = false;

                    var ioLevelDate = ioLevelsSheet.Rows.SelectMany(x => x.IoLevelDate).First(x =>
                        x.Level.Equals(data.Level, StringComparison.CurrentCulture) &&
                        x.Domain.Equals(data.Domain, StringComparison.CurrentCulture));
                    if (ioLevelDate != null && !data.Vdd.Equals(ioLevelDate.Vdd, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Vih.Equals(ioLevelDate.Vih, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Vil.Equals(ioLevelDate.Vil, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Voh.Equals(ioLevelDate.Voh, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Vol.Equals(ioLevelDate.Vol, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Ioh.Equals(ioLevelDate.Ioh, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    if (ioLevelDate != null && !data.Iol.Equals(ioLevelDate.Iol, StringComparison.CurrentCulture))
                        row.IsGroupPin = true;
                    //if (ioLevelDate != null && !data.CharAppliedPins.Equals(ioLevelDate.CharAppliedPins, StringComparison.CurrentCulture))
                    //    row.IsGroupPin = false;
                }

                if (row.IoLevelDate.Select(x => x.Domain).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                    row.IsGroupPin = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vdd).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vih).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vil).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Voh).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Vol).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Ioh).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.Iol).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }
                else if (row.IoLevelDate.Select(x => x.CharAppliedPins).Distinct().Count() != 1)
                {
                    row.IsTheSameRow = false;
                }

                row.Domain = row.IoLevelDate.Select(x => x.Domain).Distinct().First();
            }
        }

        private void GetBscanApplyPins(IoLevelsSheet ioLevelsSheet)
        {
            foreach (var row in ioLevelsSheet.Rows)
            {
                row.IsBscanApplyPins = false;

                var bscanApplyPinsData = row.IoLevelDate
                    .Where(o => o.CharAppliedPins.Equals("v", StringComparison.CurrentCultureIgnoreCase)).Select(o => o)
                    .ToList();
                if (bscanApplyPinsData.Any())
                    row.IsBscanApplyPins = true;
            }
        }

        private IoLevelsSheet ReadSheetData()
        {
            var ioLevelsSheet = new IoLevelsSheet(_sheetName);
            var ioLevelsItem = new IoLevelsItem();
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new IoLevelsRow(_sheetName);
                row.RowNum = i;
                if (_typeIndex != -1)
                    row.Type = _excelWorksheet.GetMergedCellValue(i, _typeIndex).Trim();
                if (_pinNameIndex != -1)
                    row.PinName = _excelWorksheet.GetMergedCellValue(i, _pinNameIndex).Trim();
                if (_fsddIndex != -1)
                    row.Fsdd = _excelWorksheet.GetMergedCellValue(i, _fsddIndex).Trim();

                var cnt = 0;
                for (var j = _startColNumber + 1; j <= _endColNumber; j++)
                {
                    var levelName = _excelWorksheet.GetMergedCellValue(_startRowNumber - 1, j).Trim();
                    var headerName = _excelWorksheet.GetMergedCellValue(_startRowNumber, j).Trim();

                    if (!string.IsNullOrEmpty(levelName) &&
                        headerName.Equals(HeaderDomain, StringComparison.OrdinalIgnoreCase))
                    {
                        if (cnt != 0)
                            row.IoLevelDate.Add(ioLevelsItem);
                        ioLevelsItem = new IoLevelsItem(levelName);
                        cnt++;
                    }

                    var value = _excelWorksheet.GetMergedCellValue(i, j).Trim();
                    if (headerName.Equals(HeaderDomain, StringComparison.OrdinalIgnoreCase))
                    {
                        ioLevelsItem.Domain = value.Trim();
                        ioLevelsItem.DomainIndex = j;
                    }
                    else if (headerName.Equals(HeaderVdd, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Vdd = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderVih, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Vih = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderVil, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Vil = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderVoh, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Voh = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderVol, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Vol = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderIoh, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Ioh = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderIol, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.Iol = value.Trim().Replace(" ", "");
                    else if (headerName.Equals(HeaderCharApplyPins, StringComparison.OrdinalIgnoreCase))
                        ioLevelsItem.CharAppliedPins = value.Trim().Replace(" ", "");

                    if (j == _endColNumber)
                    {
                        if (cnt != 0)
                            row.IoLevelDate.Add(ioLevelsItem);
                        ioLevelsItem = new IoLevelsItem(levelName);
                        cnt++;
                    }
                }

                ioLevelsSheet.Rows.Add(row);
            }

            ioLevelsSheet.TypeIndex = _typeIndex;
            ioLevelsSheet.PinNameIndex = _pinNameIndex;
            ioLevelsSheet.FsddIndex = _fsddIndex;
            FillEmptyCell(ioLevelsSheet);
            GetTheSameDomain(ioLevelsSheet);
            GetBscanApplyPins(ioLevelsSheet);

            return ioLevelsSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderType, StringComparison.OrdinalIgnoreCase))
                {
                    _typeIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPinName, StringComparison.OrdinalIgnoreCase))
                {
                    _pinNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderFsdd, StringComparison.OrdinalIgnoreCase)) _fsddIndex = i;
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i < rowNum; i++)
                for (var j = 1; j < colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(HeaderPinName, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _typeIndex = -1;
            _pinNameIndex = -1;
            _fsddIndex = -1;
        }
    }
}