using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonReaderLib.Input;
using CommonReaderLib.PatternListCsv;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommonReaderLib.DebugPlan
{
    public class AiTestPlanReader : MySheetReader
    {
        private const string ConHeaderUseNotUse = "Use/Not Use";
        private const string ConHeaderComment = "Comment";
        private const string ConHeaderTestInstanceName = "Test instance name";
        private const string ConHeaderAiType = "AI type";
        private const string ConHeaderDataLoggingSetting = "Data logging setting";
        private const string ConHeaderTimeset = "Timeset";
        private const string ConHeaderVoltageCategory = "Voltage Category";
        private const string ConHeaderOrder = "Order";
        private const string ConHeaderSearch = "Search";
        private const string ConHeaderTempCondition = "Temp. Condition";
        private const string ConHeaderPattern = @"Pattern\d+";
        private const string ConStart = "Start";
        private const string ConStop = "Stop";
        private const string ConStep = "Step";

        private const string ConSelsramDssc = "SELSRAM_DSSC";

        private readonly List<int> _indexPatterns = new List<int>();
        private readonly Dictionary<string, string> _indexPins = new Dictionary<string, string>();
        private int _indexAIType = -1;
        private int _indexComment = -1;
        private int _indexDataLoggingSetting = -1;
        private int _indexOrder = -1;
        private int _indexSearch = -1;
        private int _indexTempCondition = -1;
        private int _indexTestInstanceName = -1;
        private int _indexTimeset = -1;

        private int _indexUseNotUse = -1;
        private int _indexVoltageCategory = -1;

        private int _indexSelsramDssc = -1;

        public AiTestPlanSheet ReadSheet(ExcelWorksheet worksheet)
        {
            var sheetName = worksheet.Name;

            var sheet = new AiTestPlanSheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                sheet.AddDimensionError();
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                sheet.AddFirstHeaderError(ConHeaderUseNotUse);
                return null;
            }

            GetHeaderIndex();

            sheet = ReadSheet(sheetName);

            return sheet;
        }

        private AiTestPlanSheet ReadSheet(string sheetName)
        {
            var sheet = new AiTestPlanSheet(sheetName);
            for (var i = StartRowNumber + 1; i <= EndRowNumber; i++)
            {
                var row = new AiTestPlanRow(sheetName);
                row.RowNum = i;
                if (_indexUseNotUse != -1)
                    row.UseNotUse = ExcelWorksheet.GetMergedCellValue(i, _indexUseNotUse).Trim();
                if (_indexComment != -1)
                    row.Comment = ExcelWorksheet.GetMergedCellValue(i, _indexComment).Trim();
                if (_indexTestInstanceName != -1)
                    row.TestInstanceName = ExcelWorksheet.GetMergedCellValue(i, _indexTestInstanceName).Trim();
                if (_indexAIType != -1)
                    row.AiType = ExcelWorksheet.GetMergedCellValue(i, _indexAIType).Trim();
                if (_indexDataLoggingSetting != -1)
                    row.DataLoggingSetting = ExcelWorksheet.GetMergedCellValue(i, _indexDataLoggingSetting).Trim();
                if (_indexTimeset != -1)
                    row.Timeset = ExcelWorksheet.GetMergedCellValue(i, _indexTimeset).Trim();
                if (_indexVoltageCategory != -1)
                    row.VoltageCategory = ExcelWorksheet.GetMergedCellValue(i, _indexVoltageCategory).Trim();
                if (_indexOrder != -1)
                    row.Order = ExcelWorksheet.GetMergedCellValue(i, _indexOrder).Trim();
                if (_indexSearch != -1)
                    row.Search = ExcelWorksheet.GetMergedCellValue(i, _indexSearch).Trim();
                if (_indexTempCondition != -1)
                    row.TempCondition = ExcelWorksheet.GetMergedCellValue(i, _indexTempCondition).Trim();
                if (_indexSelsramDssc != -1)
                    row.SelsramDssc = ExcelWorksheet.GetMergedCellValue(i, _indexSelsramDssc).Trim();
                var initStop = false;
                foreach (var index in _indexPatterns)
                {
                    var pattern = ExcelWorksheet.GetMergedCellValue(i, index).Trim();
                    if (!string.IsNullOrEmpty(pattern))
                    {
                        row.Patterns.Add(new PatternDate(pattern));
                        if (pattern.IsInit() && initStop == false)
                        {
                            row.Inits.Add(new PatternDate(pattern));
                        }
                        else
                        {
                            initStop = true;
                            row.Payloads.Add(new PatternDate(pattern));
                        }
                    }
                }

                foreach (var index in _indexPins)
                {
                    var pin = new Pin();
                    pin.Name = index.Value;
                    var arr = index.Key.Split(';').ToList();
                    var num = -1;
                    if (int.TryParse(arr.First(), out num) && num != -1)
                    {
                        pin.Start = ExcelWorksheet.GetMergedCellValue(i, num).Trim();
                        pin.IndexStart = num;
                    }

                    if (arr.Count >= 2 && !string.IsNullOrEmpty(arr.ElementAt(1)))
                        if (int.TryParse(arr.ElementAt(1), out num) && num != -1)
                        {
                            pin.Stop = ExcelWorksheet.GetMergedCellValue(i, num).Trim();
                            pin.IndexStop = num;
                        }

                    if (arr.Count >= 3 && !string.IsNullOrEmpty(arr.ElementAt(2)))
                        if (int.TryParse(arr.ElementAt(2), out num) && num != -1)
                        {
                            pin.Step = ExcelWorksheet.GetMergedCellValue(i, num).Trim();
                            pin.IndexStep = num;
                        }

                    row.Pins.Add(pin);
                }
                if (!string.IsNullOrEmpty(row.UseNotUse))
                    sheet.Rows.Add(row);
            }

            sheet.IndexUseNotUse = _indexUseNotUse;
            sheet.IndexComment = _indexComment;
            sheet.IndexTestInstanceName = _indexTestInstanceName;
            sheet.IndexAiType = _indexAIType;
            sheet.IndexDataLoggingSetting = _indexDataLoggingSetting;
            sheet.IndexTimeset = _indexTimeset;
            sheet.IndexVoltageCategory = _indexVoltageCategory;
            sheet.IndexOrder = _indexOrder;
            sheet.IndexSearch = _indexSearch;
            sheet.IndexTempCondition = _indexTempCondition;
            sheet.IndexSelsramDssc = _indexSelsramDssc;
            sheet.IndexPatternStart = _indexPatterns.First();
            sheet.IndexStartRow = StartRowNumber;

            return sheet;
        }

        private void GetHeaderIndex()
        {
            for (var i = StartColNumber; i <= EndColNumber; i++)
            {
                var header = ExcelWorksheet.GetCellValue( StartRowNumber, i).Trim();
                if (header.Equals(ConHeaderUseNotUse, StringComparison.OrdinalIgnoreCase))
                {
                    _indexUseNotUse = i;
                    continue;
                }

                if (header.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _indexComment = i;
                    continue;
                }

                if (header.Equals(ConHeaderTestInstanceName, StringComparison.OrdinalIgnoreCase))
                {
                    _indexTestInstanceName = i;
                    continue;
                }

                if (header.Equals(ConHeaderAiType, StringComparison.OrdinalIgnoreCase))
                {
                    _indexAIType = i;
                    continue;
                }

                if (header.Equals(ConHeaderDataLoggingSetting, StringComparison.OrdinalIgnoreCase))
                {
                    _indexDataLoggingSetting = i;
                    continue;
                }

                if (header.Equals(ConHeaderTimeset, StringComparison.OrdinalIgnoreCase))
                {
                    _indexTimeset = i;
                    continue;
                }

                if (header.Equals(ConHeaderVoltageCategory, StringComparison.OrdinalIgnoreCase))
                {
                    _indexVoltageCategory = i;
                    continue;
                }

                if (header.Equals(ConHeaderOrder, StringComparison.OrdinalIgnoreCase))
                {
                    _indexOrder = i;
                    continue;
                }

                if (header.Equals(ConHeaderSearch, StringComparison.OrdinalIgnoreCase))
                {
                    _indexSearch = i;
                    continue;
                }

                if (header.Equals(ConHeaderTempCondition, StringComparison.OrdinalIgnoreCase))
                {
                    _indexTempCondition = i;
                    continue;
                }

                if (Regex.IsMatch(header, ConHeaderPattern, RegexOptions.IgnoreCase))
                {
                    _indexPatterns.Add(i);
                    continue;
                }

                if (header.Equals(ConSelsramDssc, StringComparison.OrdinalIgnoreCase))
                {
                    _indexSelsramDssc = i;
                    continue;
                }

                var topHeader = ExcelWorksheet.GetCellValue(StartRowNumber - 1, i).Trim();
                if (Regex.IsMatch(header, ConStart, RegexOptions.IgnoreCase) &&
                    !string.IsNullOrEmpty(topHeader))
                {
                    var indexStop = "";
                    var header_1 = ExcelWorksheet.GetCellValue(StartRowNumber, i + 1).Trim();
                    if (Regex.IsMatch(header_1, ConStop, RegexOptions.IgnoreCase))
                        indexStop = (i + 1).ToString();

                    var indexStep = "";
                    var header_2 = ExcelWorksheet.GetCellValue(StartRowNumber, i + 2).Trim();
                    if (Regex.IsMatch(header_2, ConStep, RegexOptions.IgnoreCase))
                        indexStep = (i + 2).ToString();

                    _indexPins.Add(i + ";" + indexStop + ";" + indexStep, topHeader);
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            var colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (ExcelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(ConHeaderUseNotUse, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRowNumber = i;
                        return true;
                    }

            return false;
        }
    }

    public class AiTestPlanSheet : MySheet
    {
        public int IndexAiType = -1;
        public int IndexComment = -1;
        public int IndexDataLoggingSetting = -1;
        public int IndexOrder = -1;
        public int IndexSearch = -1;
        public int IndexTempCondition = -1;
        public int IndexSelsramDssc = -1;
        public int IndexPatternStart = -1;
        public int IndexTestInstanceName = -1;
        public int IndexTimeset = -1;
        public int IndexUseNotUse = -1;
        public int IndexVoltageCategory = -1;

        public int IndexStartRow = -1;

        #region Constructor

        public AiTestPlanSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<AiTestPlanRow>();
        }

        #endregion

        public List<AiTestPlanRow> Rows { set; get; }

        internal void Chcek()
        {
            CheckByColumn();
        }

        private void CheckByColumn()
        {
            foreach (var row in Rows)
            {
                if (!row.UseNotUse.Equals("Use", StringComparison.CurrentCultureIgnoreCase) &&
                      !row.UseNotUse.Equals("Not Use", StringComparison.CurrentCultureIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.FormatError,
                        ErrorLevel = EnumErrorLevel.Warning,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexUseNotUse,
                        Message = string.Format("[UseNotUse] The syntax \"{0}\" can not be identified, and this row will be ignored !!!", row.UseNotUse)
                    });
                }

                if (!row.AiType.Equals("Data log", StringComparison.CurrentCultureIgnoreCase) &&
                      !row.AiType.Equals("1D", StringComparison.CurrentCultureIgnoreCase) &&
                      !row.AiType.Equals("2D", StringComparison.CurrentCultureIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.FormatError,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexAiType,
                        Message = string.Format("[AIType] The syntax \"{0}\" can not be identified, and treat this as Data log !!!", row.AiType)
                    });
                }

                if (!row.DataLoggingSetting.Equals("NA", StringComparison.CurrentCultureIgnoreCase) &&
                      !row.DataLoggingSetting.StartsWith("DFC", StringComparison.CurrentCultureIgnoreCase) &&
                      !Regex.IsMatch(row.DataLoggingSetting, @"\d+\s?FC", RegexOptions.IgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.FormatError,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexDataLoggingSetting,
                        Message = string.Format("[Data Logging Setting] The syntax \"{0}\" can not be identified, and treat this as NA !!!", row.DataLoggingSetting)
                    });
                }

                if (!row.SelsramDssc.StartsWith("SELSRM", StringComparison.CurrentCultureIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.FormatError,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexSelsramDssc,
                        Message = string.Format("[SELSRM] The syntax \"{0}\" can not be identified, and this will be ignored !!!", row.SelsramDssc)
                    });
                }

                #region check order
                var arr = row.Order.Split(',').ToList();
                if (!string.IsNullOrEmpty(row.Order))
                {
                    if (row.EnumAiType == EnumAiType.Datalog)
                    {
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            ColNum = IndexOrder,
                            Message = string.Format("[Order] The AI type is datalog, so \"{0}\" will be ignored !!!", row.Order)
                        });
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo_1D && arr.Count() != 1)
                    {
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            ColNum = IndexOrder,
                            Message = string.Format("[Order] The AI type is 1D shmoo, so the count of \"{0}\" should be 1 !!!", row.Order)
                        });
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo_2D && arr.Count() != 2)
                    {
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            ColNum = IndexOrder,
                            Message = string.Format("[Order] The AI type is 1D shmoo, so the count of \"{0}\" should be 2 !!!", row.Order)
                        });
                    }

                    foreach (var item in arr)
                    {
                        var pins = row.Pins.Select(x => x.Name).ToList();
                        if (!pins.Contains(item, StringComparer.OrdinalIgnoreCase))
                        {
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = IndexOrder,
                                Message = string.Format("[Order] The pin \"{0}\" is not exiested !!!", item)
                            });
                        }
                    }
                }
                #endregion

                CheckJump(row);
            }
        }

        private void CheckJump(AiTestPlanRow row)
        {
            if (row.EnumAiType == EnumAiType.Shmoo_1D)
            {
                if (row.IsJump())
                {
                    var testMethods = row.GettestMethods();
                    if (testMethods.Count != 1)
                    {
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            ColNum = IndexSearch,
                            Message = string.Format("[Search] The \"{0}\" don't match AI type Shmoo_1D !!!", row.Search)
                        });
                    }
                    else
                    {
                        var pin = row.GetPinElementAt(0);
                        var stepCnt = pin.Steps;
                        if (testMethods.First().Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase) &&
                            stepCnt < int.Parse(testMethods.First().Arguments))
                        {
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = IndexSearch,
                                Message = string.Format("[Search] The jump steps of \"{0}\" should be smaller than pin \"{1}\" \"{2}\" !!!",
                                    row.Search, pin.Name, stepCnt)
                            });
                        }
                    }
                }
            }
            else if (row.EnumAiType == EnumAiType.Shmoo_2D)
            {
                if (row.IsJump())
                {
                    var testMethods = row.GettestMethods();
                    if (testMethods.Count != 2)
                    {
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            ColNum = IndexSearch,
                            Message = string.Format("[Search] The \"{0}\" don't match AI type Shmoo_2D !!!",
                                row.Search)
                        });
                    }
                    else
                    {
                        var pin = row.GetPinElementAt(0);
                        var stepCnt = pin.Steps;
                        if (testMethods.First().Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase) &&
                            stepCnt < int.Parse(testMethods.First().Arguments))
                        {
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = IndexSearch,
                                Message = string.Format("[Search] The jump steps of \"{0}\" should be smaller than pin \"{1}\" \"{2}\" !!!",
                                    row.Search, pin.Name, stepCnt)
                            });
                        }
                        var testMethod = testMethods.ElementAt(1);
                        if (testMethod.Name.Equals("Jump", StringComparison.CurrentCultureIgnoreCase))
                        {
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = IndexSearch,
                                Message = string.Format("[Search] The Y Shmoo of \"{0}\" can not be jump !!!", row.Search)
                            });
                        }
                    }
                }
            }
        }

        internal void Check()
        {
            foreach (var row in Rows)
            {
                var sweepPinCount = 0;
                if (row.EnumAiType == EnumAiType.Shmoo_1D)
                    sweepPinCount = 1;
                if (row.EnumAiType == EnumAiType.Shmoo_2D)
                    sweepPinCount = 2;

                var pins = row.Pins.Where(x => x.IsSearch).ToList();
                if (!string.IsNullOrEmpty(row.Order))
                {
                    var pinList = new List<Pin>();
                    foreach (var pin in row.Order.Split(','))
                        if (pins.Exists(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                            pinList.Add(
                                pins.Find(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)));

                    if (pinList.Count != sweepPinCount || pins.Count != sweepPinCount)
                        Errors.Add(new Error
                        {
                            EnumErrorType = EnumErrorType.FormatError,
                            ErrorLevel = EnumErrorLevel.Error,
                            SheetName = SheetName,
                            RowNum = row.RowNum,
                            Message = string.Format(
                                "The sweep pins of AI type {0} is not matched with {1} !!!",
                                row.AiType, row.Order)
                        });
                }
            }
        }

        internal void ChcekTimeSet(bool checkTimeSet1, bool checkTimeSet2)
        {
            throw new NotImplementedException();
        }

        internal void ChcekTimeSet(List<string> timeSets)
        {
            foreach (var row in Rows)
            {
                if (!timeSets.Contains(row.Timeset, StringComparer.OrdinalIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.Missing,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexTimeset,
                        Message = string.Format("[Timeset] The {0} is not existed in timeSet folder or test program !!!", row.Timeset)
                    });
                }
            }
        }

        internal void ChcekDcSpec(List<string> dcspecs)
        {
            foreach (var row in Rows)
            {
                if (!dcspecs.Contains(row.VoltageCategory, StringComparer.OrdinalIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.Missing,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexVoltageCategory,
                        Message = string.Format("[VoltageCategory] The {0} is not existed in program !!!", row.VoltageCategory)
                    });
                }
            }
        }

        internal void ChcekPattern(List<string> patterns, PatternListSheet patternListSheet)
        {
            foreach (var row in Rows)
            {
                for (int i = 0; i < row.Patterns.Count; i++)
                {
                    PatternDate pattern = row.Patterns[i];
                    if (!patternListSheet.Rows.Exists(x => x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        if (!patterns.Exists(x => x.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.Missing,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = IndexPatternStart + i,
                                Message = string.Format("[Pattern] The pattern \"{0}\" not existed in pattern folder !!!", pattern.OriName)
                            });
                    }
                }
            }
        }

        internal void ChcekPins(List<string> pins, List<string> acSymbols)
        {
            var firstRow = Rows.First();
            for (int i = 0; i < firstRow.Pins.Count; i++)
            {
                var pin = firstRow.Pins[i];
                if (!pins.Contains(pin.Name, StringComparer.OrdinalIgnoreCase) &&
                    !acSymbols.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                {
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.Missing,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = IndexStartRow - 1,
                        ColNum = pin.IndexStart,
                        Message = string.Format("[Pin] The pin \"{0}\" not existed in test program !!!", pin.Name)
                    });
                }
            }

            foreach (var row in Rows)
            {
                for (int i = 0; i < row.Pins.Count; i++)
                {
                    var pin = row.Pins[i];

                    if (pins.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        string value;
                        if (!string.IsNullOrEmpty(pin.Start) && !pin.Start.TryConvertToVolt(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStart,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Start)
                            });
                        if (!string.IsNullOrEmpty(pin.Stop) && !pin.Stop.TryConvertToVolt(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStop,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Stop)
                            });
                        if (!string.IsNullOrEmpty(pin.Step) && !pin.Step.TryConvertToVolt(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStep,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Step)
                            });
                    }
                    else if (acSymbols.Contains(pin.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        string value;
                        if (!string.IsNullOrEmpty(pin.Start) && !pin.Start.TryConvertToFreq(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStart,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Start)
                            });
                        if (!string.IsNullOrEmpty(pin.Stop) && !pin.Stop.TryConvertToFreq(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStop,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Stop)
                            });
                        if (!string.IsNullOrEmpty(pin.Step) && !pin.Step.TryConvertToFreq(out value))
                            Errors.Add(new Error
                            {
                                EnumErrorType = EnumErrorType.FormatError,
                                ErrorLevel = EnumErrorLevel.Error,
                                SheetName = SheetName,
                                RowNum = row.RowNum,
                                ColNum = pin.IndexStep,
                                Message = string.Format("[Pin] The syntax of value \"{0}\" is not correct !!!", pin.Step)
                            });
                    }
                }
            }
        }
    }

    public class AiTestPlanRow : MyRow
    {
        internal EnumDataLoggingSettingType EnumDataLoggingSettingType
        {
            get
            {
                if (DataLoggingSetting.Contains("DFC"))
                    return EnumDataLoggingSettingType.DFC;
                if (DataLoggingSetting.Contains("FC"))
                    return EnumDataLoggingSettingType.FC;
                return EnumDataLoggingSettingType.NA;
            }
        }


        public List<string> GetTimeSetsByPayloads(PatternListSheet patternListSheet)
        {
            var timeSets = new List<string>();
            foreach (var payload in Payloads)
                if (patternListSheet.Rows.Exists(x =>
                        x.Pattern.Equals(payload.OriName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var timeSet = Path.GetFileNameWithoutExtension(patternListSheet.Rows
                        .Find(x => x.Pattern.Equals(payload.OriName, StringComparison.CurrentCultureIgnoreCase)).TimeSet);
                    timeSets.Add(timeSet);
                }

            return timeSets;
        }

        #region Property

        public string UseNotUse { set; get; }
        public string Comment { set; get; }
        public string TestInstanceName { set; get; }
        public string AiType { set; get; }
        public string DataLoggingSetting { set; get; }
        public string Timeset { set; get; }
        public string VoltageCategory { set; get; }
        public string Order { set; get; }
        public string Search { set; get; }
        public string TempCondition { set; get; }
        public List<PatternDate> Inits { get; set; }
        public List<PatternDate> Payloads { get; set; }
        public List<PatternDate> Patterns { get; set; }
        public List<Pin> Pins { get; set; }
        public string PatSetName { get; set; }

        public string DcCategory
        {
            get
            {
                var arr = VoltageCategory.Split(' ', '_').ToList();
                var last = arr.Last();
                if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase) ||
                    last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase) ||
                    last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                    return string.Join("_", arr.GetRange(0, arr.Count() - 1));
                return VoltageCategory;
            }
        }

        public string DcSelector
        {
            get
            {
                var last = VoltageCategory.Split(' ', '_').Last();
                if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
                    return "Max";
                if (last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase))
                    return "Min";
                if (last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                    return "Typ";
                return "Typ";
            }
        }

        public string TestName
        {
            get
            {
                var last = VoltageCategory.Split(' ', '_').Last();
                var dcSelector = "NV";
                if (last.EndsWith("HV", StringComparison.CurrentCultureIgnoreCase))
                    dcSelector = "HV";
                if (last.EndsWith("LV", StringComparison.CurrentCultureIgnoreCase))
                    dcSelector = "LV";
                if (last.EndsWith("NV", StringComparison.CurrentCultureIgnoreCase))
                    dcSelector = "NV";
                return PatSetName + "_" + dcSelector;
            }
        }

        public EnumAiType EnumAiType
        {
            get
            {
                if (AiType.EndsWith("Data log", StringComparison.CurrentCultureIgnoreCase))
                    return EnumAiType.Datalog;
                if (AiType.Equals("1D", StringComparison.CurrentCultureIgnoreCase))
                    return EnumAiType.Shmoo_1D;
                if (AiType.Equals("2D", StringComparison.CurrentCultureIgnoreCase))
                    return EnumAiType.Shmoo_2D;
                return EnumAiType.Datalog;
            }
        }

        public string Parameter
        {
            get
            {
                if (EnumAiType == EnumAiType.Datalog)
                    return TestName;
                if (EnumAiType == EnumAiType.Shmoo_1D)
                    return TestName + " " + CharName;
                if (EnumAiType == EnumAiType.Shmoo_2D)
                    return TestName + " " + CharName;
                return TestName;
            }
        }

        public string CharName
        {
            get
            {
                if (EnumAiType == EnumAiType.Datalog)
                    return "";
                if (EnumAiType == EnumAiType.Shmoo_1D)
                    return "Char_1D_" + TestName;
                if (EnumAiType == EnumAiType.Shmoo_2D)
                    return "Char_2D_" + TestName;
                return "";
            }
        }

        public string SelsramDssc { get; set; }

        internal object GetBlock()
        {
            var firstSeg = DcCategory.Split('_').First();
            if (firstSeg.Equals("TD", StringComparison.CurrentCultureIgnoreCase))
                return "Scan";
            if (firstSeg.Equals("SA", StringComparison.CurrentCultureIgnoreCase))
                return "Scan";
            return firstSeg;
        }

        internal bool IsJump()
        {
            return Regex.IsMatch(Search, "Jump", RegexOptions.IgnoreCase);
        }

        internal List<TestMethod> GettestMethods()
        {
            var testMethods = new List<TestMethod>();
            var searchs = Search.Split(';');
            foreach (var search in searchs)
            {
                var testMethod = new TestMethod();
                var arr = search.Split(' ');
                if (arr.First().Equals("Jump", StringComparison.CurrentCultureIgnoreCase))
                {
                    testMethod.Name = arr.First();
                    if (arr.Length == 2)
                        testMethod.Arguments = arr.Last();
                }
                else
                    testMethod.Name = arr.First();
                testMethods.Add(testMethod);
            }
            return testMethods;
        }

        internal Pin GetPinElementAt(int index)
        {
            if (string.IsNullOrEmpty(Order))
                return Pins.ElementAt(index);
            else
            {
                var pinList = new List<Pin>();
                foreach (var pin in Order.Split(','))
                    if (Pins.Exists(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                        pinList.Add(Pins.Find(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)));
                if (pinList.Count > index)
                    return pinList.ElementAt(index);
                else
                    return Pins.ElementAt(index);
            }
        }

        #endregion

        #region Constructor

        public AiTestPlanRow()
        {
            Inits = new List<PatternDate>();
            Payloads = new List<PatternDate>();
            Patterns = new List<PatternDate>();
            Pins = new List<Pin>();
        }

        public AiTestPlanRow(string sourceSheetName)
        {
            SheetName = sourceSheetName;
            Inits = new List<PatternDate>();
            Payloads = new List<PatternDate>();
            Patterns = new List<PatternDate>();
            Pins = new List<Pin>();
        }

        #endregion
    }

    internal class TestMethod
    {
        public string Name { get; set; }
        public string Arguments { get; set; }
    }

    public class PatternDate
    {
        private string _name;

        public PatternDate(string pattern)
        {
            _name = pattern;
        }

        public string Name
        {
            get
            {
                if (WithDate)
                    return _name + "_PAT";
                return _name;
            }
            set { _name = value; }
        }

        public bool WithDate
        {
            get { return Regex.IsMatch(_name, @"\d+_\w\d+_\d+$"); }
        }
        public string OriName
        {
            get { return _name; }
            set { _name = value; }
        }
    }

    public enum EnumDataLoggingSettingType
    {
        FC,
        DFC,
        NA
    }

    public class Pin
    {
        public string Name { get; set; }
        public string Start { get; set; }
        public int IndexStart { get; set; }
        public string Stop { get; set; }
        public int IndexStop { get; set; }
        public string Step { get; set; }
        public int IndexStep { get; set; }

        public bool IsSearch
        {
            get { return Start != Stop; }
        }

        public string ShmooName
        {
            get { return "Shmoo_" + Name; }
        }

        public int Steps
        {
            get
            {
                double startValue;
                var start = Start.ConvertNumber();
                if (!double.TryParse(start, out startValue))
                    return 0;

                double stopValue;
                var stop = Stop.ConvertNumber();
                if (!double.TryParse(stop, out stopValue))
                    return 0;

                double stepValue;
                var step = Step.ConvertNumber();
                if (!double.TryParse(step, out stepValue))
                    return 0;

                var value = (int)((stopValue - startValue) / stepValue);
                return value;
            }
        }

        public bool TryParseRangeStep(out int value)
        {
            value = 0;
            double val;
            if (!double.TryParse(Start, out val))
                return false;
            if (!double.TryParse(Stop, out val))
                return false;
            if (!double.TryParse(Step, out val))
                return false;

            var startValue = double.Parse(Start);
            var stopValue = double.Parse(Stop);
            var step = double.Parse(Step);
            value = (int)((stopValue - startValue) / step);
            return true;
        }
    }
}