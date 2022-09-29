using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest
{
    public class TestPlanReader
    {
        protected const string RegRunPattern = @"Run\s*the\s*pattern.*";
        protected const string ErrorMsgMissPat = "Missing pattern in \"Pattern\" Column";
        protected const string ErrorMsgMissPatWithMeasurement = "Missing pattern for measuments";
        protected const string ErrorMsgPatWithMeasurement = "pattern row exist measuments";
        protected int AnalogSetupIndex;
        protected int DescriptionIndex;
        protected int EndRow;
        protected int ForceCharIndex;
        protected int ForceIndex;
        protected Dictionary<string, int> HeaderOrder;
        protected Dictionary<string, string> JobMap;
        protected int MeasIndex;
        protected List<MeasLimit> MeasLimits;
        protected int MiscInfoIndex;
        protected int NoBinOutIndex;
        protected int PatternIndex;
        protected int RegisterIndex;

        protected ExcelWorksheet _excelWorksheet;
        protected int StartRow;
        protected int TestNameIndex;
        protected int TtrIndex;

        public TestPlanReader()
        {
            JobMap = new Dictionary<string, string>
            {
                {"CP1", "CP"},
                {"CP2", "CP"},
                {"FT1", "FT"},
                {"FT2", "FT"},
                {"FT3", "FT"},
                {"QA", "FT"},
                {"HTOL", "FT"}
            };
        }

        public virtual TestPlanSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var planSheet = new TestPlanSheet();
            planSheet.SheetName = excelWorksheet.Name;
            _excelWorksheet = excelWorksheet;
            HeaderOrder = excelWorksheet.GetHeaderOrder();
            StartRow = 2;

            if (excelWorksheet.Dimension == null)
                return planSheet;

            EndRow = excelWorksheet.Dimension.End.Row;
            PatternRow patRow = null;

            //CheckUnRecognizeHeader
            CheckUnRecognizeHeader(HardIpPattern.KnownHeaders);

            //Initial herder index
            GetHeaderIndex();
            SetHardipSheetsHeaderIdx(planSheet);

            //find the forceCondition used by all the patterns in the sheet
            planSheet.ForceIndex = ForceIndex;
            planSheet.MeasIndex = MeasIndex;
            planSheet.ForceStr = GetSheetForceConditions();

            #region read testPlan patterns

            for (var j = StartRow; j <= EndRow; j++)
            {
                var patternName = excelWorksheet.GetCellValue(j, PatternIndex).TrimSpace().Trim(',')
                    .Trim(';').ToLower();
                if (patternName != "")
                {
                    patRow = new PatternRow();
                    ReadPatternRow(excelWorksheet, j, patRow, patternName);
                    //PatternExistMeasurement

                    //patRow.IsMultiplePayload = SearchInfo.CheckMultiPayLoad(patRow);
                    planSheet.PatternRows.Add(patRow);
                    var measStr = excelWorksheet.GetCellValue(j, MeasIndex);
                    if (!string.IsNullOrEmpty(measStr))
                        ErrorManager.AddError(EnumErrorType.PatternExistMeasurement, EnumErrorLevel.Error,
                            excelWorksheet.Name, j, ErrorMsgPatWithMeasurement);
                }
                else
                {
                    CheckRunPattern(excelWorksheet, j);
                    var measStr = excelWorksheet.GetCellValue(j, MeasIndex);

                    //If meas column not blank, but not belong to any pattern, flag error
                    if (patRow == null && !string.IsNullOrEmpty(measStr))
                    {
                        ErrorManager.AddError(EnumErrorType.MisPatternForMeasurement, EnumErrorLevel.Error,
                            excelWorksheet.Name, j, ErrorMsgMissPatWithMeasurement);
                        continue;
                    }

                    //Will ignore the rows that Pattern column, Force Condition column, RegisterAssignment and Meas column are blank
                    if (string.IsNullOrEmpty(measStr))
                    {
                        var registerAssignment = excelWorksheet.GetCellValue(j, RegisterIndex).TrimSpace();
                        //read RegisterAssignment
                        if (patRow != null)
                        {
                            patRow.RegisterAssignment =
                                UpdatePatternItem(patRow.RegisterAssignment, registerAssignment);

                            var forceStr = excelWorksheet.GetCellValue(j, ForceIndex);
                            //Read post pattern force condition
                            patRow.PostPatForceCondition = UpdatePatternItem(patRow.PostPatForceCondition, forceStr);
                        }

                        continue;
                    }

                    ReadNonPatternRowNew(excelWorksheet, ref j, patRow);
                }
            }

            #endregion

            //DividePatternRow(planSheet.PatternRows);
            ParsePlanSheet(planSheet);
            return planSheet;
        }

        private void ParsePlanSheet(TestPlanSheet planSheet)
        {
            var tpPreProcess = new TestPlanSheetPatPreprocess(planSheet);
            tpPreProcess.UpdateSheetPattern();

            var testPlanPatParser = new TestPlanPatParser(planSheet);
            testPlanPatParser.ConvertTpPatterns();
        }

        protected void CheckUnRecognizeHeader(List<string> knownHeaders)
        {
            foreach (var header in HeaderOrder.Keys)
            {
                var isKnown = false;
                foreach (var knowHeader in knownHeaders)
                    if (Regex.IsMatch(header, knowHeader, RegexOptions.IgnoreCase))
                    {
                        isKnown = true;
                        break;
                    }

                if (!isKnown)
                {
                    var outString = "UnRecognized header : " + header;
                    ErrorManager.AddError(EnumErrorType.UnrecognisedHeader, EnumErrorLevel.Warning, _excelWorksheet.Name, 1,
                        outString);
                }
            }
        }

        protected void GetHeaderIndex()
        {
            MeasLimits = new List<MeasLimit>();
            var jobList = JobMap.Keys.ToList();
            foreach (var job in jobList)
            {
                var limit = new MeasLimit(job);
                if (JobMap.Keys.Contains(job) && !JobMap[job].Equals(job, StringComparison.OrdinalIgnoreCase))
                {
                    limit.LoHeaderIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, job + HardIpPattern.LoHeader, false);
                    if (limit.LoHeaderIndex == 1)
                        limit.LoHeaderIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder,
                            JobMap[job] + HardIpPattern.LoHeader);
                    limit.HiHeaderIndex =
                        GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, job + HardIpPattern.HiHeader, false);
                    if (limit.HiHeaderIndex == 1)
                        limit.HiHeaderIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder,
                            JobMap[job] + HardIpPattern.HiHeader);
                }
                else
                {
                    limit.LoHeaderIndex =
                        GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, job + HardIpPattern.LoHeader);
                    limit.HiHeaderIndex =
                        GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, job + HardIpPattern.HiHeader);
                }

                MeasLimits.Add(limit);
            }

            TtrIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.TtrHeader);
            var headerIndex = HeaderOrder.FirstOrDefault(a =>
                Regex.IsMatch(a.Key, HardIpPattern.NoBinOutHeader, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(a.Key, "Pattern Release Status")).Value;
            if (headerIndex > 0)
                NoBinOutIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.NoBinOutHeader);
            DescriptionIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.DescriptionHeader);
            PatternIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.PatternHeader);
            ForceIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.ForceConditionHeader);
            ForceCharIndex =
                GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.ForceConditionCharHeader, false);
            AnalogSetupIndex =
                GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.AnalogSetupHeader, false);
            TestNameIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.TestNameHeader, false);
            RegisterIndex =
                GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.RegisterAssignmentHeader);
            MiscInfoIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.MiscInfoHeader);
            MeasIndex = GetHeaderIndex(_excelWorksheet.Name, HeaderOrder, HardIpPattern.MeasHeader);
        }

        protected string GetSheetForceConditions()
        {
            var forceStr = "";
            for (var j = StartRow; j <= EndRow; j++)
            {
                var patternName = _excelWorksheet.GetCellValue(j, PatternIndex).TrimSpace();
                if (patternName != "")
                    break;
                var forceCondition = _excelWorksheet.GetCellValue(j, ForceIndex).TrimSpace();
                if (forceCondition != "")
                    forceStr += forceCondition + ";";
                StartRow++;
            }

            return forceStr.Trim(';');
        }

        protected virtual TestPlanRow ReadTestPlanRow(int rowNum)
        {
            var testPlanRow = new TestPlanRow();
            testPlanRow.RowNum = rowNum;
            testPlanRow.Description =
                _excelWorksheet.GetCellValue(rowNum, DescriptionIndex).TrimSpace();
            testPlanRow.ForceCondition =
                _excelWorksheet.GetMergedCellValue(rowNum, ForceIndex).TrimSpace();
            testPlanRow.ForceConditionChar =
                _excelWorksheet.GetMergedCellValue(rowNum, ForceCharIndex).TrimSpace();
            testPlanRow.MiscInfo =
                _excelWorksheet.GetMergedCellValue(rowNum, MiscInfoIndex).TrimSpace();
            testPlanRow.RegisterAssignment =
                _excelWorksheet.GetMergedCellValue(rowNum, RegisterIndex).TrimSpace();
            if (TestNameIndex != 1)
                testPlanRow.TestName = _excelWorksheet.GetCellValue(rowNum, TestNameIndex).TrimSpace();
            if (_excelWorksheet.MergedCells[rowNum, MeasIndex] != null)
            {
                testPlanRow.MergeRowNumForMeas = new ExcelAddress(_excelWorksheet.MergedCells[rowNum, MeasIndex]).Start.Row;
                testPlanRow.Meas = _excelWorksheet.GetCellValue(testPlanRow.MergeRowNumForMeas, MeasIndex)
                    .Replace("\n", "").Replace("\t", "");
            }
            else
            {
                testPlanRow.MergeRowNumForMeas = 0;
                testPlanRow.Meas = _excelWorksheet.GetCellValue(rowNum, MeasIndex).Replace("\n", "")
                    .Replace("\t", "");
            }

            foreach (var limit in MeasLimits)
            {
                var newLimit = new MeasLimit(limit.JobName);
                newLimit.LoLimit = _excelWorksheet.GetCellValue(rowNum, limit.LoHeaderIndex).Replace(" ", "");
                newLimit.HiLimit = _excelWorksheet.GetCellValue(rowNum, limit.HiHeaderIndex).Replace(" ", "");
                newLimit.LoHeaderIndex = limit.LoHeaderIndex;
                newLimit.HiHeaderIndex = limit.HiHeaderIndex;
                testPlanRow.Limits.Add(newLimit);
            }

            return testPlanRow;
        }

        public void SetHardipSheetsHeaderIdx(TestPlanSheet testPlanSheet)
        {
            var idxDic = testPlanSheet.PlanHeaderIdx;
            idxDic.Add("ttrIndex", TtrIndex);
            idxDic.Add("noBinOutIndex", TtrIndex);
            idxDic.Add("descriptionIndex", DescriptionIndex);
            idxDic.Add("patternIndex", PatternIndex);
            idxDic.Add("forceIndex", ForceIndex);
            idxDic.Add("forceCharIndex", ForceCharIndex);
            idxDic.Add("analogSetupIndex", AnalogSetupIndex);
            idxDic.Add("testNameIndex", TestNameIndex);
            idxDic.Add("registerIndex", RegisterIndex);
            idxDic.Add("miscInfoIndex", MiscInfoIndex);
            idxDic.Add("measIndex", MeasIndex);

            //if (HardIpDataMain.TestPlanData != null)
            //    HardIpDataMain.TestPlanData.PlanHeaderIdx.Add(testPlanSheet.SheetName, idxDic);
        }

        protected void ReadPatternRow(ExcelWorksheet excelWorksheet, int rowIndex, PatternRow patRow, string patternName)
        {
            patRow.PatternColumnNum = PatternIndex;
            patRow.SheetName = excelWorksheet.Name;
            patRow.RowNum = rowIndex;
            patRow.TtrStr = excelWorksheet.GetMergedCellValue(rowIndex, TtrIndex).TrimSpace();
            patRow.NoBinOutStr = excelWorksheet.GetMergedCellValue(rowIndex, NoBinOutIndex).TrimSpace();
            patRow.Description = excelWorksheet.GetCellValue(rowIndex, DescriptionIndex);
            patRow.Pattern = new PatternClass(patternName);

            patRow.ForceCondition.ForceCondition =
                excelWorksheet.GetCellValue(rowIndex, ForceIndex).TrimSpace();
            if (ForceCharIndex != 1)
                patRow.ForceConditionChar =
                    excelWorksheet.GetCellValue(rowIndex, ForceCharIndex).TrimSpace();
            if (AnalogSetupIndex != 1)
                patRow.AnalogSetup =
                    excelWorksheet.GetCellValue(rowIndex, AnalogSetupIndex).TrimSpace();
            if (TestNameIndex != 1)
                patRow.SpecifyTestName =
                    excelWorksheet.GetCellValue(rowIndex, TestNameIndex).TrimSpace();
            patRow.RegisterAssignment =
                excelWorksheet.GetCellValue(rowIndex, RegisterIndex).TrimSpace();
            patRow.MiscInfo = excelWorksheet.GetCellValue(rowIndex, MiscInfoIndex).TrimSpace();
            //patRow.IsMultiplePayload = SearchInfo.CheckMultiPayLoad(patRow);
        }

        protected virtual void ReadNonPatternRowNew(ExcelWorksheet sheet, ref int rowIndex, PatternRow patRow)
        {
            try
            {
                //find whether the force condition column is merged or not
                var isMerged = sheet.MergedCells[rowIndex, ForceIndex] != null;

                var patChildRow = new PatSubChildRow();
                patChildRow.IsMerged = isMerged;
                if (!isMerged)
                {
                    patChildRow.TpRows.Add(ReadTestPlanRow(rowIndex));
                }
                else
                {
                    for (var mergedRow = rowIndex;
                         mergedRow <= new ExcelAddress(sheet.MergedCells[rowIndex, ForceIndex]).End.Row;
                         mergedRow++)
                        patChildRow.TpRows.Add(ReadTestPlanRow(mergedRow));
                    rowIndex = new ExcelAddress(sheet.MergedCells[rowIndex, ForceIndex]).End.Row;
                }

                patRow.PatChildRows.Add(patChildRow);
            }
            catch (Exception e)
            {
                var message = e.Message;
                Response.Report(message, EnumMessageLevel.Error, 100);
            }
        }

        protected void CheckRunPattern(ExcelWorksheet excelWorksheet, int rowIndex)
        {
            var description = excelWorksheet.GetCellValue(rowIndex, DescriptionIndex).TrimSpace();
            //If "Run the pattern" exist Description column but pattern column is blank, flag error
            if (Regex.IsMatch(description, RegRunPattern, RegexOptions.IgnoreCase))
                ErrorManager.AddError(EnumErrorType.MissingPatternInTestPlan, EnumErrorLevel.Error, excelWorksheet.Name,
                    rowIndex, ErrorMsgMissPat);
        }

        protected string UpdatePatternItem(string origin, string value)
        {
            var result = origin;
            if (!string.IsNullOrEmpty(value))
                result += ";" + value;
            return result.Trim(';');
        }

        public int GetHeaderIndex(string sheetName, Dictionary<string, int> headerOrder, string header,
            bool optionalFlag = true)
        {
            var headerIndex = headerOrder.FirstOrDefault(a =>
                Regex.IsMatch(a.Key, header, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(a.Key, "Pattern Release Status")).Value;
            if (headerIndex > 0)
                return headerIndex;
            if (optionalFlag)
            {
                header = header.Replace(@"\s*", " ").Replace(@"\s", " ").Replace(@".*", "");
                var errorMessage = "Missing header " + header + " in sheet " + sheetName;
                ErrorManager.AddError(EnumErrorType.MissingHeader, EnumErrorLevel.Error, sheetName, 1, errorMessage,
                    header);
            }

            return 1;
        }
    }
}