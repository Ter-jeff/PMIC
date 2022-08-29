using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.Others
{
    public class SpecFinder
    {
        private readonly List<AcSpecSheet> _acSpecSheets;
        private readonly List<DcSpecSheet> _dcSpecSheets;
        private readonly ExcelWorksheet _excelWorksheet;
        private readonly List<GlobalSpec> _globalSpecs;
        private readonly List<LevelSheet> _levelSheets;

        public SpecFinder(List<GlobalSpecSheet> globalSpecsSheets, List<AcSpecSheet> acSpecSheets,
            List<DcSpecSheet> dcSpecSheets, List<LevelSheet> levelSheets)
        {
            _acSpecSheets = acSpecSheets;
            _dcSpecSheets = dcSpecSheets;
            _levelSheets = levelSheets;
            _globalSpecs = globalSpecsSheets.Where(x => x != null).SelectMany(x => x.GetGlobalSpecs()).ToList();
            var excelPackage = new ExcelPackage(new FileInfo("Test"));
            _excelWorksheet = excelPackage.Workbook.Worksheets.Add("Test");
        }

        public string GetValue(InstanceRow instanceRow, string formula, string frequencyName)
        {
            var replace = GetReplaceString(instanceRow, frequencyName);
            formula = formula.Replace(frequencyName, replace);
            return GetFormulaValue(formula, _globalSpecs);
        }

        public string GetValue(InstanceRow instanceRow, string formula)
        {
            if (instanceRow == null) return "";
            if (_levelSheets.Exists(x =>
                    x.SheetName.Equals(instanceRow.PinLevels, StringComparison.CurrentCultureIgnoreCase)))
            {
                var levelSheet = _levelSheets.Find(x =>
                    x.SheetName.Equals(instanceRow.PinLevels, StringComparison.CurrentCultureIgnoreCase));
                if (levelSheet.LevelRows.Exists(x =>
                        x.PinName.Equals(formula, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var row = levelSheet.LevelRows
                        .Where(x => x.PinName.Equals(formula, StringComparison.CurrentCultureIgnoreCase))
                        .First(x => x.Parameter.Equals("Vps", StringComparison.CurrentCultureIgnoreCase) ||
                                    x.Parameter.Equals("Vmain", StringComparison.CurrentCultureIgnoreCase));
                    formula = row.Value.Trim(' ').Trim('=').Trim('_');
                }
            }

            formula = GetReplaceString(instanceRow, formula);
            return GetFormulaValue(formula, _globalSpecs);
        }

        private string GetReplaceString(InstanceRow instanceRow, string frequencyName)
        {
            foreach (var acSpecSheet in _acSpecSheets)
                if (acSpecSheet.FindValue(instanceRow.AcCategory, instanceRow.AcSelector, ref frequencyName))
                    return frequencyName;

            foreach (var dcSpecSheet in _dcSpecSheets)
                if (dcSpecSheet.FindValue(instanceRow.DcCategory, instanceRow.DcSelector, ref frequencyName))
                    return frequencyName;

            return frequencyName;
        }

        private string GetFormulaValue(string formula, List<GlobalSpec> globalSpecs)
        {
            try
            {
                formula = formula.TrimStart(' ').TrimStart('=');
                formula = ReplaceGlobalValue(formula, globalSpecs);
                _excelWorksheet.Cells["A1"].Formula = formula;
                _excelWorksheet.Cells["A1"].Calculate();
                double value;
                if (double.TryParse(_excelWorksheet.Cells["A1"].Value.ToString(), out value))
                    return _excelWorksheet.Cells["A1"].Value.ToString();
                return "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private string ReplaceGlobalValue(string formula, List<GlobalSpec> globalSpecs)
        {
            const string regex = @"(?<Name>\w+)";
            foreach (var item in Regex.Matches(formula, regex, RegexOptions.IgnoreCase))
                if (globalSpecs.Exists(x => item.ToString().Equals("_" + x.Symbol, StringComparison.OrdinalIgnoreCase)))
                {
                    var row = globalSpecs.Find(x =>
                        item.ToString().Equals("_" + x.Symbol, StringComparison.OrdinalIgnoreCase));
                    var value = row.Value.TrimStart(' ').TrimStart('=');
                    var newValue = "";
                    var global = "_" + row.Symbol.ToUpper();
                    double output;
                    while (!double.TryParse(value, out output) && newValue != value)
                        newValue = ReplaceGlobalValue(value, globalSpecs);
                    return Regex.Replace(formula, global, value, RegexOptions.IgnoreCase);
                }

            return formula;
        }
    }
}