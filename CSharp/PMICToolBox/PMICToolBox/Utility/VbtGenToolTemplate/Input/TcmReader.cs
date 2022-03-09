using PmicAutomation.Utility.VbtGenToolTemplate.Base;
using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutomation.Utility.VbtGenToolTemplate.Input
{
    public class TcmRow
    {
        public VbtTestPlanRow ConvertVbtTestPlanRow()
        {
            VbtTestPlanRow vbtTestPlanRow = new VbtTestPlanRow();
            if (!string.IsNullOrEmpty(AteDatalogName))
            {
                vbtTestPlanRow.TopList = "Etc";
                vbtTestPlanRow.Command = "DATALOG";
                vbtTestPlanRow.RegisterMacroName = AteDatalogName;
                vbtTestPlanRow.Unit = AteLimitUnit;
                vbtTestPlanRow.HighLimit = AteLimitUpper;
                vbtTestPlanRow.LowLimit = AteLimitLower;
                vbtTestPlanRow.Comment = Comments;
            }

            return vbtTestPlanRow;
        }

        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string TestPlanId { set; get; }
        public string TestPlanName { set; get; }
        public string TcmId { set; get; }
        public string AteDatalogName { set; get; }
        public string TestDescription { set; get; }
        public string AteLimitLower { set; get; }
        public string AteLimitTarget { set; get; }
        public string AteLimitUpper { set; get; }
        public string AteLimitUnit { set; get; }
        public string Lower { set; get; }
        public string Target { set; get; }
        public string Upper { set; get; }
        public string Unit { set; get; }
        public string Priority { set; get; }
        public string TrimTest { set; get; }
        public string ProductionCp { set; get; }
        public string Qual { set; get; }
        public string Char { set; get; }
        public string CoveragePlan { set; get; }
        public string CoverageAct { set; get; }
        public string BenchCorrelationStatus { set; get; }
        public string Comments { set; get; }

        #endregion

        #region Constructor

        public TcmRow()
        {
            TestPlanId = "";
            TestPlanName = "";
            TcmId = "";
            AteDatalogName = "";
            TestDescription = "";
            AteLimitLower = "";
            AteLimitTarget = "";
            AteLimitUpper = "";
            AteLimitUnit = "";
            Lower = "";
            Target = "";
            Upper = "";
            Unit = "";
            Priority = "";
            TrimTest = "";
            ProductionCp = "";
            Qual = "";
            Char = "";
            CoveragePlan = "";
            CoverageAct = "";
            BenchCorrelationStatus = "";
            Comments = "";
        }

        public TcmRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            TestPlanId = "";
            TestPlanName = "";
            TcmId = "";
            AteDatalogName = "";
            TestDescription = "";
            AteLimitLower = "";
            AteLimitTarget = "";
            AteLimitUpper = "";
            AteLimitUnit = "";
            Lower = "";
            Target = "";
            Upper = "";
            Unit = "";
            Priority = "";
            TrimTest = "";
            ProductionCp = "";
            Qual = "";
            Char = "";
            CoveragePlan = "";
            CoverageAct = "";
            BenchCorrelationStatus = "";
            Comments = "";
        }

        #endregion
    }

    public class TcmSheet
    {
        #region Constructor

        public TcmSheet(string name)
        {
            Name = name;
            Rows = new List<TcmRow>();
        }

        #endregion

        private List<VbtTestPlanRow> ConvertRows(IEnumerable<IGrouping<string, TcmRow>> groupRows)
        {
            List<VbtTestPlanRow> vbtTestPlanRows = new List<VbtTestPlanRow>();
            foreach (IGrouping<string, TcmRow> rows in groupRows)
            {
                if (!string.IsNullOrEmpty(rows.First().TestPlanName))
                {
                    vbtTestPlanRows.Add(GetFunctionStartRow(rows));

                    foreach (TcmRow row in rows)
                    {
                        vbtTestPlanRows.Add(row.ConvertVbtTestPlanRow());
                    }

                    vbtTestPlanRows.Add(GetFunctionEndRow());
                }
            }

            if (vbtTestPlanRows.Count != 0)
            {
                vbtTestPlanRows.Add(GetGenCompleteRow());
            }

            return vbtTestPlanRows;
        }

        public List<VbtTestPlanRow> ConvertTestRows()
        {
            IEnumerable<IGrouping<string, TcmRow>> groupRows = Rows.Where(x =>
                x.TrimTest.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                x.TrimTest.Equals("Y", StringComparison.OrdinalIgnoreCase)).GroupBy(x => x.TestPlanName);
            return ConvertRows(groupRows);
        }

        public List<VbtTestPlanRow> ConvertPostRows()
        {
            IEnumerable<IGrouping<string, TcmRow>> groupRows = Rows.Where(x =>
                x.TrimTest.Equals("No", StringComparison.OrdinalIgnoreCase) ||
                x.TrimTest.Equals("N", StringComparison.OrdinalIgnoreCase)).GroupBy(x => x.TestPlanName);
            return ConvertRows(groupRows);
        }

        public List<VbtTestPlanRow> ConvertOtherRows()
        {
            IEnumerable<IGrouping<string, TcmRow>> groupRows = Rows.Where(x =>
                !(x.TrimTest.Equals("No", StringComparison.OrdinalIgnoreCase) ||
                  x.TrimTest.Equals("N", StringComparison.OrdinalIgnoreCase) ||
                  x.TrimTest.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
                  x.TrimTest.Equals("Y", StringComparison.OrdinalIgnoreCase))).GroupBy(x => x.TestPlanName);
            return ConvertRows(groupRows);
        }

        private VbtTestPlanRow GetFunctionStartRow(IGrouping<string, TcmRow> rows)
        {
            VbtTestPlanRow vbtTestPlanRow = new VbtTestPlanRow
            {
                TopList = "Start_End_Setup",
                Command = "start_of_test",
                FunctionName = rows.First().TestPlanName.Replace(" ", "_")
            };
            return vbtTestPlanRow;
        }

        private VbtTestPlanRow GetFunctionEndRow()
        {
            VbtTestPlanRow vbtTestPlanRow = new VbtTestPlanRow {TopList = "Start_End_Setup", Command = "END_OF_TEST"};
            return vbtTestPlanRow;
        }

        private VbtTestPlanRow GetGenCompleteRow()
        {
            VbtTestPlanRow vbtTestPlanRow = new VbtTestPlanRow {TopList = "Start_End_Setup", Command = "Gen_Complete"};
            return vbtTestPlanRow;
        }

        #region Field

        #endregion

        #region Properity

        public string Name { get; set; }
        public List<TcmRow> Rows { get; set; }

        #endregion
    }

    public class TcmReader
    {
        private const string HeaderTestPlanId = "Test Plan Id";
        private const string HeaderTestPlanName = "Test Plan Name";
        private const string HeaderTcmId = "TCM Id";
        private const string HeaderAteDatalogName = "ATE datalog Name (follow Name standard)";
        private const string HeaderTestDescription = "Test Description";
        private const string HeaderAteLimit = "ATE Limits";
        private const string HeaderAteLimitLower = "Lower";
        private const string HeaderAteLimitTarget = "Target";
        private const string HeaderAteLimitUpper = "Upper";
        private const string HeaderAteLimitUnit = "Unit";
        private const string HeaderDesignSpec = "Design Spec";
        private const string HeaderDesignSpecLower = "Lower";
        private const string HeaderDesignSpecTarget = "Target";
        private const string HeaderDesignSpecUpper = "Upper";
        private const string HeaderDesignSpecUnit = "Unit";
        private const string HeaderPriority = "Priority (P1/P2/P3/P4)";
        private const string HeaderTrimTest = "Trim Test";
        private const string HeaderProductionCp = "Production (CP)";
        private const string HeaderQual = "Qual";
        private const string HeaderChar = "Char";
        private const string HeaderCoveragePlan = "Coverage Plan";
        private const string HeaderCoverageAct = "Coverage act";
        private const string HeaderBenchCorrelationStatus = "Bench Correlation status";
        private const string HeaderComments = "Comments";
        private int _ateDatalogNameIndex = -1;
        private int _ateLimitIndex = -1;
        private int _ateLimitLowerIndex = -1;
        private int _ateLimitTargetIndex = -1;
        private int _ateLimitUnitIndex = -1;
        private int _ateLimitUpperIndex = -1;
        private int _benchCorrelationStatusIndex = -1;
        private int _charIndex = -1;
        private int _commentsIndex = -1;
        private TcmSheet _comPinSheet;
        private int _coverageActIndex = -1;
        private int _coveragePlanIndex = -1;
        private int _designSpecIndex = -1;
        private int _designSpecLowerIndex = -1;
        private int _designSpecTargetIndex = -1;
        private int _designSpecUnitIndex = -1;
        private int _designSpecUpperIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _priorityIndex = -1;
        private int _productionCpIndex = -1;
        private int _qualIndex = -1;
        private string _name;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _tcmIdIndex = -1;
        private int _testDescriptionIndex = -1;
        private int _testPlanIdIndex = -1;
        private int _testPlanNameIndex = -1;
        private int _trimTestIndex = -1;

        public TcmSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _comPinSheet = new TcmSheet(_name);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _comPinSheet = ReadSheetData();

            return _comPinSheet;
        }

        private TcmSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                TcmRow row = new TcmRow(_name) {RowNum = i};
                if (_testPlanIdIndex != -1)
                {
                    row.TestPlanId = _excelWorksheet.GetMergeCellValue(i, _testPlanIdIndex).Trim();
                }

                if (_testPlanNameIndex != -1)
                {
                    row.TestPlanName = _excelWorksheet.GetMergeCellValue(i, _testPlanNameIndex).Trim();
                }

                if (_tcmIdIndex != -1)
                {
                    row.TcmId = _excelWorksheet.GetMergeCellValue(i, _tcmIdIndex).Trim();
                }

                if (_ateDatalogNameIndex != -1)
                {
                    row.AteDatalogName = _excelWorksheet.GetMergeCellValue(i, _ateDatalogNameIndex).Trim();
                }

                if (_testDescriptionIndex != -1)
                {
                    row.TestDescription = _excelWorksheet.GetMergeCellValue(i, _testDescriptionIndex).Trim();
                }

                if (_ateLimitLowerIndex != -1)
                {
                    row.AteLimitLower = _excelWorksheet.GetMergeCellValue(i, _ateLimitLowerIndex).Trim();
                }

                if (_ateLimitTargetIndex != -1)
                {
                    row.AteLimitTarget = _excelWorksheet.GetMergeCellValue(i, _ateLimitTargetIndex).Trim();
                }

                if (_ateLimitUpperIndex != -1)
                {
                    row.AteLimitUpper = _excelWorksheet.GetMergeCellValue(i, _ateLimitUpperIndex).Trim();
                }

                if (_ateLimitUnitIndex != -1)
                {
                    row.AteLimitUnit = _excelWorksheet.GetMergeCellValue(i, _ateLimitUnitIndex).Trim();
                }

                if (_designSpecLowerIndex != -1)
                {
                    row.Lower = _excelWorksheet.GetMergeCellValue(i, _designSpecLowerIndex).Trim();
                }

                if (_designSpecTargetIndex != -1)
                {
                    row.Target = _excelWorksheet.GetMergeCellValue(i, _designSpecTargetIndex).Trim();
                }

                if (_designSpecUpperIndex != -1)
                {
                    row.Upper = _excelWorksheet.GetMergeCellValue(i, _designSpecUpperIndex).Trim();
                }

                if (_designSpecUnitIndex != -1)
                {
                    row.Unit = _excelWorksheet.GetMergeCellValue(i, _designSpecUnitIndex).Trim();
                }

                if (_priorityIndex != -1)
                {
                    row.Priority = _excelWorksheet.GetMergeCellValue(i, _priorityIndex).Trim();
                }

                if (_trimTestIndex != -1)
                {
                    row.TrimTest = _excelWorksheet.GetMergeCellValue(i, _trimTestIndex).Trim();
                }

                if (_productionCpIndex != -1)
                {
                    row.ProductionCp = _excelWorksheet.GetMergeCellValue(i, _productionCpIndex).Trim();
                }

                if (_qualIndex != -1)
                {
                    row.Qual = _excelWorksheet.GetMergeCellValue(i, _qualIndex).Trim();
                }

                if (_charIndex != -1)
                {
                    row.Char = _excelWorksheet.GetMergeCellValue(i, _charIndex).Trim();
                }

                if (_coveragePlanIndex != -1)
                {
                    row.CoveragePlan = _excelWorksheet.GetMergeCellValue(i, _coveragePlanIndex).Trim();
                }

                if (_coverageActIndex != -1)
                {
                    row.CoverageAct = _excelWorksheet.GetMergeCellValue(i, _coverageActIndex).Trim();
                }

                if (_benchCorrelationStatusIndex != -1)
                {
                    row.BenchCorrelationStatus =
                        _excelWorksheet.GetMergeCellValue(i, _benchCorrelationStatusIndex).Trim();
                }

                if (_commentsIndex != -1)
                {
                    row.Comments = _excelWorksheet.GetMergeCellValue(i, _commentsIndex).Trim();
                }

                _comPinSheet.Rows.Add(row);
            }

            return _comPinSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderTestPlanId, StringComparison.OrdinalIgnoreCase))
                {
                    _testPlanIdIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTestPlanName, StringComparison.OrdinalIgnoreCase))
                {
                    _testPlanNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTcmId, StringComparison.OrdinalIgnoreCase))
                {
                    _tcmIdIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteDatalogName, StringComparison.OrdinalIgnoreCase))
                {
                    _ateDatalogNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTestDescription, StringComparison.OrdinalIgnoreCase))
                {
                    _testDescriptionIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimit, StringComparison.OrdinalIgnoreCase))
                {
                    _ateLimitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitLower, StringComparison.OrdinalIgnoreCase))
                {
                    _ateLimitLowerIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitTarget, StringComparison.OrdinalIgnoreCase))
                {
                    _ateLimitTargetIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitUpper, StringComparison.OrdinalIgnoreCase))
                {
                    _ateLimitUpperIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitUnit, StringComparison.OrdinalIgnoreCase))
                {
                    _ateLimitUnitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpec, StringComparison.OrdinalIgnoreCase))
                {
                    _designSpecIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecLower, StringComparison.OrdinalIgnoreCase))
                {
                    _designSpecLowerIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecTarget, StringComparison.OrdinalIgnoreCase))
                {
                    _designSpecTargetIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecUpper, StringComparison.OrdinalIgnoreCase))
                {
                    _designSpecUpperIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecUnit, StringComparison.OrdinalIgnoreCase))
                {
                    _designSpecUnitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPriority, StringComparison.OrdinalIgnoreCase))
                {
                    _priorityIndex = i;
                    continue;
                }

                if (lStrHeader.ToUpper().Contains(HeaderTrimTest.ToUpper()))
                {
                    _trimTestIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderProductionCp, StringComparison.OrdinalIgnoreCase))
                {
                    _productionCpIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderQual, StringComparison.OrdinalIgnoreCase))
                {
                    _qualIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderChar, StringComparison.OrdinalIgnoreCase))
                {
                    _charIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCoveragePlan, StringComparison.OrdinalIgnoreCase))
                {
                    _coveragePlanIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCoverageAct, StringComparison.OrdinalIgnoreCase))
                {
                    _coverageActIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderBenchCorrelationStatus, StringComparison.OrdinalIgnoreCase))
                {
                    _benchCorrelationStatusIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderComments, StringComparison.OrdinalIgnoreCase))
                {
                    _commentsIndex = i;
                }
            }

            int nextRowNumber = _startRowNumber + 1;
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(nextRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderTestPlanId, StringComparison.OrdinalIgnoreCase))
                {
                    _testPlanIdIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTestPlanName, StringComparison.OrdinalIgnoreCase))
                {
                    _testPlanNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTcmId, StringComparison.OrdinalIgnoreCase))
                {
                    _tcmIdIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteDatalogName, StringComparison.OrdinalIgnoreCase))
                {
                    _ateDatalogNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderTestDescription, StringComparison.OrdinalIgnoreCase))
                {
                    _testDescriptionIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitLower, StringComparison.OrdinalIgnoreCase) && i >= _ateLimitIndex)
                {
                    _ateLimitLowerIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitTarget, StringComparison.OrdinalIgnoreCase) && i >= _ateLimitIndex)
                {
                    _ateLimitTargetIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitUpper, StringComparison.OrdinalIgnoreCase) && i >= _ateLimitIndex)
                {
                    _ateLimitUpperIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderAteLimitUnit, StringComparison.OrdinalIgnoreCase) && i >= _ateLimitIndex)
                {
                    _ateLimitUnitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecLower, StringComparison.OrdinalIgnoreCase) &&
                    i >= _designSpecIndex)
                {
                    _designSpecLowerIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecTarget, StringComparison.OrdinalIgnoreCase) &&
                    i >= _designSpecIndex)
                {
                    _designSpecTargetIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecUpper, StringComparison.OrdinalIgnoreCase) &&
                    i >= _designSpecIndex)
                {
                    _designSpecUpperIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderDesignSpecUnit, StringComparison.OrdinalIgnoreCase) &&
                    i >= _designSpecIndex)
                {
                    _designSpecUnitIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderPriority, StringComparison.OrdinalIgnoreCase))
                {
                    _priorityIndex = i;
                    continue;
                }

                if (lStrHeader.ToUpper().Contains(HeaderTrimTest.ToUpper()))
                {
                    _trimTestIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderProductionCp, StringComparison.OrdinalIgnoreCase))
                {
                    _productionCpIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderQual, StringComparison.OrdinalIgnoreCase))
                {
                    _qualIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderChar, StringComparison.OrdinalIgnoreCase))
                {
                    _charIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCoveragePlan, StringComparison.OrdinalIgnoreCase))
                {
                    _coveragePlanIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderCoverageAct, StringComparison.OrdinalIgnoreCase))
                {
                    _coverageActIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderBenchCorrelationStatus, StringComparison.OrdinalIgnoreCase))
                {
                    _benchCorrelationStatusIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderComments, StringComparison.OrdinalIgnoreCase))
                {
                    _commentsIndex = i;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            for (int j = 1; j <= colNum; j++)
            {
                if (_excelWorksheet.GetMergeCellValue(i, j).Trim()
                    .Equals(HeaderTestPlanId, StringComparison.OrdinalIgnoreCase))
                {
                    _startRowNumber = i;
                    return true;
                }
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
            _testPlanIdIndex = -1;
            _testPlanNameIndex = -1;
            _tcmIdIndex = -1;
            _ateDatalogNameIndex = -1;
            _testDescriptionIndex = -1;
            _ateLimitLowerIndex = -1;
            _ateLimitTargetIndex = -1;
            _ateLimitUpperIndex = -1;
            _ateLimitUnitIndex = -1;
            _designSpecLowerIndex = -1;
            _designSpecTargetIndex = -1;
            _designSpecUpperIndex = -1;
            _designSpecUnitIndex = -1;
            _priorityIndex = -1;
            _trimTestIndex = -1;
            _productionCpIndex = -1;
            _qualIndex = -1;
            _charIndex = -1;
            _coveragePlanIndex = -1;
            _coverageActIndex = -1;
            _benchCorrelationStatusIndex = -1;
            _commentsIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> dictionary = new List<Dictionary<string, string>>();
            foreach (TcmRow row in _comPinSheet.Rows)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>
                {
                    {"Test Plan Id", row.TestPlanId},
                    {"Test Plan Name", row.TestPlanName},
                    {"TCM Id", row.TcmId},
                    {"ATE datalog Name (follow Name standard)", row.AteDatalogName},
                    {"Test Description", row.TestDescription},
                    {"ATELimitLower", row.AteLimitLower},
                    {"ATELimitTarget", row.AteLimitTarget},
                    {"ATELimitUpper", row.AteLimitUpper},
                    {"ATELimitUnit", row.AteLimitUnit},
                    {"Lower", row.Lower},
                    {"Target", row.Target},
                    {"Upper", row.Upper},
                    {"Unit", row.Unit},
                    {"Priority (P1/P2/P3/P4)", row.Priority},
                    {"Trim Test", row.TrimTest},
                    {"Production (CP)", row.ProductionCp},
                    {"Qual", row.Qual},
                    {"Char", row.Char},
                    {"Coverage Plan", row.CoveragePlan},
                    {"Coverage act", row.CoverageAct},
                    {"Bench Correlation status", row.BenchCorrelationStatus},
                    {"Comments", row.Comments}
                };
                dictionary.Add(dic);
            }

            return dictionary;
        }
    }
}