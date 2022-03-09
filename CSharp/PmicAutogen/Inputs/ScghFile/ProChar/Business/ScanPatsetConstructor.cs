using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.Utility;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.Reader;
using PmicAutogen.Local;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Business
{
    public class ScanPatSetConstructor : ProdCharPatSetConstructorBase
    {
        private readonly DataTable _payloadTypeTable;

        public ScanPatSetConstructor(IEnumerable<IProdCharSheetRow> inputRows) : base(inputRows)
        {
            _payloadTypeTable = StaticSetting.PayloadTypeTable;
            PerformanceModeList = new List<string>();
            Block = "Scan";
        }

        public List<ProdCharRowScan> WorkFlow(bool removeNonUsage = false)
        {
            if (removeNonUsage)
            {
                InitList = FilterProChar(InitList);
                PayloadList = FilterProChar(PayloadList);
            }

            var prodCharRowScans = new List<ProdCharRowScan>();
            var prodCharRows = GetPatSetFromProdChar(InitList, PayloadList);
            if (PatternListUsage != null && PatternListUsage.Any())
                prodCharRows = FilterPatSetWithUsedPatterns(prodCharRows);

            foreach (var row in prodCharRows)
            {
                var prodCharRowScan = row.NewProdCharRowScan();
                prodCharRowScan.RowNum = row.RowNum;
                prodCharRowScan.PayloadType = GetPayloadType(row.PayLoadName);
                prodCharRowScan.Prefix = GetPrefix(prodCharRowScan.PayloadType);
                prodCharRowScan.InitPatSetNameByNamingRule = "";
                prodCharRowScan.PatSetName =
                    (ComCombine.CombineByUnderLine(prodCharRowScan.Prefix, prodCharRowScan.InitPatSetNameByNamingRule) +
                    "_" + GetPayLoadName(prodCharRowScan)).Trim('_');
                prodCharRowScan.InstanceName = prodCharRowScan.PatSetName;
                prodCharRowScan.PerformanceMode = GetPerformanceMode(row, PerformanceModeList);
                CheckNop(prodCharRowScan);
                prodCharRowScans.Add(prodCharRowScan);
            }

            return prodCharRowScans;
        }

        protected override string GetPrefix(string payloadType = "")
        {
            return Domain + payloadType;
        }

        private void CheckNop(ProdCharRowScan prodCharRowScan)
        {
            var prodCharRow = (ProdCharSheetRow) prodCharRowScan.ProdCharItem;

            if (prodCharRowScan.InitPatternMissing)
            {
                prodCharRowScan.Nop = true;
                prodCharRowScan.NopType = NopType.BlankInit;
            }

            if (prodCharRow.Usage != "1")
            {
                prodCharRowScan.Nop = true;
                prodCharRowScan.NopType = NopType.NonUsage;
            }
        }

        protected virtual string GetPayloadType(string pattern)
        {
            var lStrResult = "";
            if (_payloadTypeTable == null) return "";
            for (var i = 0; i < _payloadTypeTable.Rows.Count; i++)
            {
                var lBIsMatch = true;
                for (var j = 1; j < _payloadTypeTable.Columns.Count; j++)
                {
                    var lStrMatchPattern = _payloadTypeTable.Rows[i][j].ToString();
                    var subName = GetSubName(pattern, _payloadTypeTable.Columns[j].ColumnName);
                    if (Regex.IsMatch(subName, lStrMatchPattern, RegexOptions.IgnoreCase) == false) lBIsMatch = false;
                }

                if (lBIsMatch)
                {
                    lStrResult = _payloadTypeTable.Rows[i][0].ToString();
                    break;
                }
            }

            return lStrResult;
        }
    }
}