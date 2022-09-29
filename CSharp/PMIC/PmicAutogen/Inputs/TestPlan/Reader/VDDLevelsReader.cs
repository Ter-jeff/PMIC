//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2021-07-08  Bruce Qian     #92             T-auotgen , Support UHV，ULV
//
//------------------------------------------------------------------------------ 

using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.Others;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic.GenDc.DcInitial;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class VddLevelsRow
    {
        #region Constructor

        public VddLevelsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            WsBumpName = "";
            Lv = "";
            Nv = "";
            Hv = "";
            ULv = "";
            UHv = "";
            LvRange = "";
            NvRange = "";
            HvRange = "";
            ULvRange = "";
            UHvRange = "";
            Seq = "";
            Comment = "";
            ReferenceLevel = "";
            FinalSeq = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { set; get; }
        public int RowNum { get; set; }
        public string WsBumpName { set; get; }
        public string Lv { set; get; }
        public string LvRange { set; get; }
        public string Nv { set; get; }
        public string NvRange { set; get; }
        public string Hv { set; get; }
        public string HvRange { set; get; }
        public string UHv { set; get; }
        public string UHvRange { set; get; }
        public string ULv { set; get; }
        public string ULvRange { set; get; }

        public Dictionary<string, string> ExtraSelectors = new Dictionary<string, string>();
        public string Seq { set; get; }
        public string Comment { set; get; }
        public string ReferenceLevel { get; set; }
        public string FinalSeq { get; set; }

        #endregion
    }

    public class VddLevelsSheet
    {
        #region Constructor

        public VddLevelsSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<VddLevelsRow>();
            XRows = new List<VddLevelsRow>();
        }

        #endregion

        #region Property

        public string SheetName { get; set; }

        //rows which Seq column is not equal to 'x'/'X'
        public List<VddLevelsRow> Rows { get; }

        //rows cwhich Seq column is equal to 'x'/'X'
        public List<VddLevelsRow> XRows { get; }
        public Dictionary<string, int> HeaderIndex = new Dictionary<string, int>();
        public List<int> HeaderList = new List<int>();

        public int WsBumpNameIndex = -1;
        public int LvIndex = -1;
        public int NvIndex = -1;
        public int HvIndex = -1;
        public int ULvIndex = -1;
        public int UHvIndex = -1;
        public int SeqIndex = -1;
        public int CommentIndex = -1;
        public int ReferenceLevelIndex = -1;
        public int FinalSeqIndex = -1;

        public bool ULvAllNa;
        public bool UHvAllNa;

        public const string VariblePrefix = "VDD_Levels_Information!";

        #endregion

        #region global spec

        public List<GlobalSpec> GenGlbSymbol(IfoldPowerTableSheet ifoldPowerTable)
        {
            var totalGlobalSpec = new List<GlobalSpec>();
            var globalGlb = new List<GlobalSpec>();
            var globalPlus = new List<GlobalSpec>();
            var globalMinus = new List<GlobalSpec>();
            var globalPlusUHv = new List<GlobalSpec>();
            var globalMinusULv = new List<GlobalSpec>();
            //var glbSymbolList = new List<GlobalSpec>();
            var ifoldGlbSpecsList = new List<GlobalSpec>();
            var tdelayGlbSpecsList = new List<GlobalSpec>();
            var powerSeqList = new List<GlobalSpec>();
            foreach (var row in Rows)
            {
                double nv;
                double lv;
                double hv;
                double ulv;
                double uhv;
                double.TryParse(row.Nv, out nv);
                double.TryParse(row.Lv, out lv);
                double.TryParse(row.Hv, out hv);
                double.TryParse(row.ULv, out ulv);
                double.TryParse(row.UHv, out uhv);

                var nvValue = nv == 0 ? "0" : VariblePrefix + row.NvRange;
                var plus = nv == 0 ? "0" : VariblePrefix + row.HvRange + "/" + nvValue;
                var minus = nv == 0 ? "0" : VariblePrefix + row.LvRange + "/" + nvValue;
                var plusUHv = nv == 0 ? "0" : VariblePrefix + row.UHvRange + "/" + nvValue;
                var minusULv = nv == 0 ? "0" : VariblePrefix + row.ULvRange + "/" + nvValue;

                //var plus = nv == 0 ? "0" : (hv / nv).ToString(CultureInfo.InvariantCulture);
                //var minus = nv == 0 ? "0" : (lv / nv).ToString(CultureInfo.InvariantCulture);
                //var plus_UHv = nv == 0 ? "0" : (uhv / nv).ToString(CultureInfo.InvariantCulture);
                //var minus_ULv = nv == 0 ? "0" : (ulv / nv).ToString(CultureInfo.InvariantCulture);

                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName); // _GLB
                globalGlb.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(nvValue)));

                glbSymbol = SpecFormat.GenGlbPlus(row.WsBumpName); // _GLB_Plus
                globalPlus.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(plus)));

                glbSymbol = SpecFormat.GenGlbMinus(row.WsBumpName); // _GLB_Minus
                globalMinus.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(minus)));

                if (!UHvAllNa && row.UHv.Trim() != "")
                {
                    glbSymbol = SpecFormat.GenGlbPlusUHv(row.WsBumpName); // _GLB_Plus
                    globalPlusUHv.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(plusUHv)));
                }

                if (!ULvAllNa && row.ULv.Trim() != "")
                {
                    glbSymbol = SpecFormat.GenGlbMinusULv(row.WsBumpName); // _GLB_Minus
                    globalMinusULv.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(minusULv)));
                }

                //foreach (var item in row.ExtraSelectors)
                //{
                //    double value;
                //    double.TryParse(item.Value, out value);
                //    var context = nv == 0 ? "0" : (value / nv).ToString(CultureInfo.InvariantCulture);
                //    glbSymbol = SpecFormat.GenGlbOther(row.WsBumpName, item.Key); // _GLB_HV2
                //    glbSymbolList.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(context)));
                //}

                //Modified by terry
                if (!Regex.IsMatch(row.WsBumpName, "DC30", RegexOptions.IgnoreCase))
                {
                    glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName + "_" + "Ifold"); // _Ifold_GLB
                    var ifold = GetIfoldValue(row.WsBumpName, ifoldPowerTable);
                    ifoldGlbSpecsList.Add(GenGlbSymbolWithUnCertainValue(glbSymbol, ifold, "1"));

                    glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName + "_" + "Tdelay"); // _Tdelay_GLB
                    tdelayGlbSpecsList.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue("0")));
                }

                glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName + "_" + "PowerSequence"); // _PowerSequence_GLB
                powerSeqList.Add(new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(GenPowerSeqValue(row))));
            }

            totalGlobalSpec.AddRange(globalGlb);
            totalGlobalSpec.AddRange(globalPlus);
            totalGlobalSpec.AddRange(globalMinus);
            totalGlobalSpec.AddRange(globalPlusUHv);
            totalGlobalSpec.AddRange(globalMinusULv);
            //totalGlobalSpec.AddRange(glbSymbolList);
            totalGlobalSpec.AddRange(ifoldGlbSpecsList);
            totalGlobalSpec.AddRange(tdelayGlbSpecsList);
            totalGlobalSpec.AddRange(powerSeqList);


            LocalSpecs.HasUltraVoltageUHv = globalPlusUHv.Count > 0;
            LocalSpecs.HasUltraVoltageULv = globalMinusULv.Count > 0;
            LocalSpecs.HasUltraVoltage = LocalSpecs.HasUltraVoltageUHv | LocalSpecs.HasUltraVoltageULv;

            return totalGlobalSpec;
        }

        private string GetIfoldValue(string pinName, IfoldPowerTableSheet ifoldPowerTable)
        {
            var ifold = "";
            if (ifoldPowerTable != null)
            {
                var targetIfold =
                    ifoldPowerTable.Rows.Find(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                if (targetIfold != null) ifold = targetIfold.Current;
            }


            if (!string.IsNullOrEmpty(ifold))
                return ifold;
            return "1";
        }

        public GlobalSpec GenGlbSymbolWithUnCertainValue(string glbSymbol, string value, string defaultValue)
        {
            if (string.IsNullOrEmpty(value))
            {
                double outDouble;
                var defaultSpecValue = double.TryParse(defaultValue, out outDouble)
                    ? SpecFormat.GenSpecValueSingleValue(defaultValue)
                    : SpecFormat.GenSpecValueSingleSpec(defaultValue);
                return new GlobalSpec(glbSymbol, defaultSpecValue, "", "Not specified in Power table");
            }

            return new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(value));
        }

        public string GenPowerSeqValue(VddLevelsRow row)
        {
            string lStrResult;
            const string lStrMatchPattern = @"^[a-zA-Z]*(?<str>\d+)";
            if (Regex.IsMatch(row.Seq, lStrMatchPattern))
                lStrResult = Regex.Match(row.Seq, lStrMatchPattern).Groups["str"].ToString();
            else if (row.Seq.ToUpper().Contains("EFUSE") && (row.Seq.ToUpper().Equals("OFF") || row.Seq == ""))
                lStrResult = "99";
            else
                lStrResult = "99";
            return lStrResult;
        }

        #endregion

        #region DC spec

        public List<DcSpec> GenDcSymbol(List<DcCategory> categories)
        {
            var dcSpecList = new List<DcSpec>();
            var selectorList = GetSelectorList();
            //var group = Rows.SelectMany(x => x.ExtraSelectors.Keys).Distinct().ToList();
            var group = new List<string>();
            foreach (var row in Rows)
            {
                var dcSpecSymbol = SpecFormat.GenDcSpecSymbol(row.WsBumpName);
                var dcSpec = new DcSpec(dcSpecSymbol);
                dcSpec.SelectorList = selectorList;
                SetDcSpecValue(row, ref dcSpec, categories, group);
                dcSpecList.Add(dcSpec);
            }

            return dcSpecList;
        }

        public List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }

        private void SetDcSpecValue(VddLevelsRow row, ref DcSpec dcSpec, List<DcCategory> categories,
            List<string> group)
        {
            var categoryName = string.Empty;
            try
            {
                foreach (var category in categories)
                {
                    categoryName = category.CategoryName;
                    var categoryInSpec = new CategoryInSpec(categoryName);
                    if (categoryName == Category.Conti)
                    {
                        categoryInSpec.Max = "0";
                        categoryInSpec.Min = "0";
                        categoryInSpec.Typ = "0";
                    }
                    else
                    {
                        if (categoryName.EndsWith("UltraVoltage", StringComparison.OrdinalIgnoreCase))
                        {
                            categoryInSpec.Max = "0";
                            categoryInSpec.Min = "0";
                            categoryInSpec.Typ = "0";

                            if (!UHvAllNa && row.UHv != null && row.UHv != "")
                            {
                                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                                var glbSymbolPlus = SpecFormat.GenGlbPlusUHv(row.WsBumpName);
                                categoryInSpec.Max = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolPlus);
                            }

                            if (!ULvAllNa && row.ULv != null && row.ULv != "")
                            {
                                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                                var glbSymbolMinus = SpecFormat.GenGlbMinusULv(row.WsBumpName);
                                categoryInSpec.Min = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolMinus);
                            }
                        }
                        else
                        {
                            if (row.Hv != null)
                            {
                                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                                var glbSymbolPlus = SpecFormat.GenGlbPlus(row.WsBumpName);
                                categoryInSpec.Max = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolPlus);
                            }

                            if (row.Lv != null)
                            {
                                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                                var glbSymbolMinus = SpecFormat.GenGlbMinus(row.WsBumpName);
                                categoryInSpec.Min = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolMinus);
                            }

                            if (row.Nv != null)
                            {
                                var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                                categoryInSpec.Typ = "=_" + glbSymbol;
                            }
                        }
                    }

                    if (!dcSpec.ContainsCategory(categoryInSpec.Name))
                        dcSpec.AddCategory(categoryInSpec);
                }

                foreach (var category in categories)
                {
                    categoryName = category.CategoryName;
                    foreach (var groupName in group)
                    {
                        categoryName = category.CategoryName + "_" + groupName;
                        var categoryInSpec = new CategoryInSpec(categoryName);
                        var glbSymbol = SpecFormat.GenGlbSpecSymbol(row.WsBumpName);
                        var glbSymbolOther = SpecFormat.GenGlbOther(row.WsBumpName, groupName);
                        categoryInSpec.Typ = "=_" + glbSymbol;
                        categoryInSpec.Max = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolOther);
                        categoryInSpec.Min = SpecFormat.GenGlbRatio(glbSymbol, glbSymbolOther);

                        if (!dcSpec.ContainsCategory(categoryInSpec.Name)) dcSpec.AddCategory(categoryInSpec);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Error occurred in converting TestSetting category:{0} " + e.Message,
                    categoryName));
            }
        }

        #endregion

        #region Power Table

        //public PowerTableSheet ConvertorPowerTable()
        //{
        //    PowerTableSheet powerTableSheet = new PowerTableSheet();
        //    foreach (var row in Rows)
        //    {
        //        PowerTableRow powerTableRow = new PowerTableRow();
        //        powerTableRow.PinName = row.WsBumpName;
        //        powerTableRow.PinType = PowerTablePinType.HardIpPower;
        //        powerTableRow.Vmain = row.Nv;
        //        powerTableRow.Lv = row.Lv;
        //        powerTableRow.Hv = row.Hv;
        //        powerTableSheet.AddRow(powerTableRow);
        //    }
        //    return powerTableSheet;
        //}

        #endregion
    }

    public class VddLevelsReader
    {
        private const string HeaderWsBumpName = "WS Bump Name";
        private const string HeaderLv = "LV";
        private const string HeaderNv = "NV";
        private const string HeaderHv = "HV";
        private const string HeaderULv = "ULV";
        private const string HeaderUHv = "UHV";
        private const string HeaderSeq = "SEQ";
        private const string HeaderComment = "Comment";
        private const string HeaderReferenceLevel = "Reference_Level";
        private const string HeaderFinalSeq = "Final_SEQ";
        private int _commentIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _finalSeqIndex = -1;
        private int _hvIndex = -1;
        private int _lvIndex = -1;
        private int _nvIndex = -1;
        private int _referenceLevelIndex = -1;
        private int _seqIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _uhvIndex = -1;
        private int _ulvIndex = -1;
        private VddLevelsSheet _vddLevelsSheet;

        private int _wsBumpNameIndex = -1;

        public VddLevelsSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _vddLevelsSheet = new VddLevelsSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _vddLevelsSheet = ReadSheetData();
            PostAction();

            return _vddLevelsSheet;
        }

        private VddLevelsSheet ReadSheetData()
        {
            var vddLevelsSheet = new VddLevelsSheet(_sheetName);
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new VddLevelsRow(_sheetName);
                row.RowNum = i;
                if (_wsBumpNameIndex != -1)
                    row.WsBumpName = _excelWorksheet.GetMergedCellValue(i, _wsBumpNameIndex).Trim();
                if (_lvIndex != -1)
                {
                    var lvRange = string.Empty;
                    row.Lv = _excelWorksheet.GetMergedCellValueAndAddress(i, _lvIndex, ref lvRange)
                        .Trim();
                    row.LvRange = lvRange;
                }

                if (_nvIndex != -1)
                {
                    var nvRange = string.Empty;
                    row.Nv = _excelWorksheet.GetMergedCellValueAndAddress(i, _nvIndex, ref nvRange)
                        .Trim();
                    row.NvRange = nvRange;
                }

                if (_hvIndex != -1)
                {
                    var hvRange = string.Empty;
                    row.Hv = _excelWorksheet.GetMergedCellValueAndAddress(i, _hvIndex, ref hvRange)
                        .Trim();
                    row.HvRange = hvRange;
                }

                if (_ulvIndex != -1)
                {
                    var ulvRange = string.Empty;
                    row.ULv = _excelWorksheet.GetMergedCellValueAndAddress(i, _ulvIndex, ref ulvRange)
                        .Trim();
                    row.ULvRange = ulvRange;
                }

                if (_uhvIndex != -1)
                {
                    var uhvRange = string.Empty;
                    row.UHv = _excelWorksheet.GetMergedCellValueAndAddress(i, _uhvIndex, ref uhvRange)
                        .Trim();
                    row.UHvRange = uhvRange;
                }

                if (_seqIndex != -1)
                    row.Seq = _excelWorksheet.GetMergedCellValue(i, _seqIndex).Trim();
                if (_commentIndex != -1)
                    row.Comment = _excelWorksheet.GetMergedCellValue(i, _commentIndex).Trim();
                if (_referenceLevelIndex != -1)
                    row.ReferenceLevel = _excelWorksheet.GetMergedCellValue(i, _referenceLevelIndex)
                        .Trim();
                if (_finalSeqIndex != -1)
                    row.FinalSeq = _excelWorksheet.GetMergedCellValue(i, _finalSeqIndex).Trim();

                //Alec's request(2021/6/17), if empty row, ignore all the context below this row
                if (string.IsNullOrEmpty(row.WsBumpName) && string.IsNullOrEmpty(row.Lv)
                                                         && string.IsNullOrEmpty(row.Hv) &&
                                                         string.IsNullOrEmpty(row.Seq) &&
                                                         string.IsNullOrEmpty(row.Comment))
                    break;

                foreach (var col in _vddLevelsSheet.HeaderList)
                {
                    var header = _excelWorksheet.GetMergedCellValue(_startRowNumber, col).Trim();
                    var context = _excelWorksheet.GetMergedCellValue(i, col).Trim();
                    if (header.Equals(HeaderUHv, StringComparison.CurrentCultureIgnoreCase)
                        || header.Equals(HeaderULv, StringComparison.CurrentCultureIgnoreCase))
                        row.ExtraSelectors.Add(header, context);
                }

                if (!string.IsNullOrEmpty(row.WsBumpName))
                {
                    if (!row.Seq.Equals("x", StringComparison.CurrentCultureIgnoreCase))
                        vddLevelsSheet.Rows.Add(row);
                    else
                        vddLevelsSheet.XRows.Add(row);
                }
            }

            vddLevelsSheet.WsBumpNameIndex = _wsBumpNameIndex;
            vddLevelsSheet.LvIndex = _lvIndex;
            vddLevelsSheet.NvIndex = _nvIndex;
            vddLevelsSheet.HvIndex = _hvIndex;
            vddLevelsSheet.ULvIndex = _ulvIndex;
            vddLevelsSheet.UHvIndex = _uhvIndex;
            vddLevelsSheet.SeqIndex = _seqIndex;
            vddLevelsSheet.CommentIndex = _commentIndex;
            vddLevelsSheet.ReferenceLevelIndex = _referenceLevelIndex;
            vddLevelsSheet.FinalSeqIndex = _finalSeqIndex;
            return vddLevelsSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderWsBumpName, StringComparison.OrdinalIgnoreCase))
                {
                    _wsBumpNameIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderWsBumpName, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderLv, StringComparison.OrdinalIgnoreCase))
                {
                    _lvIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderLv, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderNv, StringComparison.OrdinalIgnoreCase))
                {
                    _nvIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderNv, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderHv, StringComparison.OrdinalIgnoreCase))
                {
                    _hvIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderHv, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderULv, StringComparison.OrdinalIgnoreCase))
                {
                    _ulvIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderULv, i);
                    //continue;
                }

                if (lStrHeader.Equals(HeaderUHv, StringComparison.OrdinalIgnoreCase))
                {
                    _uhvIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderUHv, i);
                    //continue;
                }

                if (lStrHeader.Equals(HeaderSeq, StringComparison.OrdinalIgnoreCase))
                {
                    _seqIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderSeq, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderComment, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderReferenceLevel, StringComparison.OrdinalIgnoreCase))
                {
                    _referenceLevelIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderReferenceLevel, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderFinalSeq, StringComparison.OrdinalIgnoreCase))
                {
                    _finalSeqIndex = i;
                    _vddLevelsSheet.HeaderIndex.Add(HeaderFinalSeq, i);
                    continue;
                }

                if (!string.IsNullOrEmpty(lStrHeader))
                    _vddLevelsSheet.HeaderList.Add(i);
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
                        .Equals(HeaderWsBumpName, StringComparison.OrdinalIgnoreCase))
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
            _wsBumpNameIndex = -1;
            _lvIndex = -1;
            _nvIndex = -1;
            _hvIndex = -1;
            _ulvIndex = -1;
            _uhvIndex = -1;
            _seqIndex = -1;
            _commentIndex = -1;
            _referenceLevelIndex = -1;
            _finalSeqIndex = -1;
        }

        private void PostAction()
        {
            if (_vddLevelsSheet == null || _vddLevelsSheet.Rows.Count <= 0)
                return;

            var uHvNaRows = _vddLevelsSheet.Rows
                .FindAll(row => row.UHv.Equals("NA", StringComparison.CurrentCultureIgnoreCase)).ToList();
            if (uHvNaRows.Any())
                _vddLevelsSheet.UHvAllNa = true;

            var uLvNaRows = _vddLevelsSheet.Rows
                .FindAll(row => row.ULv.Equals("NA", StringComparison.CurrentCultureIgnoreCase)).ToList();
            if (uLvNaRows.Any())
                _vddLevelsSheet.ULvAllNa = true;
        }
    }
}