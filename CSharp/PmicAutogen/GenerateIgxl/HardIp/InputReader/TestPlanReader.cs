using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using AutomationCommon.EpplusErrorReport;
using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.ExcelUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using PmicAutogen.InputPackages;

namespace PmicAutogen.GenerateIgxl.HardIp.InputReader
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

        protected ExcelWorksheet Sheet;
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

        public virtual TestPlanSheet ReadSheet(ExcelWorksheet sheet)
        {
            var planSheet = new TestPlanSheet();
            planSheet.SheetName = sheet.Name;
            Sheet = sheet;
            HeaderOrder = EpplusOperation.GetHeaderOrder(sheet);
            StartRow = 2;

            if (sheet.Dimension == null)
                return planSheet;

            EndRow = sheet.Dimension.End.Row;
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
                var patternName = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, j, PatternIndex)).Trim(',')
                    .Trim(';').ToLower();
                if (patternName != "")
                {
                    patRow = new PatternRow();
                    ReadPatternRow(sheet, j, patRow, patternName);
                    //PatternExistMeasurement

                    //patRow.IsMultiplePayload = SearchInfo.CheckMultiPayLoad(patRow);
                    planSheet.PatternRows.Add(patRow);
                    var measStr = EpplusOperation.GetCellValue(sheet, j, MeasIndex);
                    if (!string.IsNullOrEmpty(measStr))
                        EpplusErrorManager.AddError(HardIpErrorType.PatternExistMeasurement, ErrorLevel.Error,
                            sheet.Name, j, ErrorMsgPatWithMeasurement);
                }
                else
                {
                    CheckRunPattern(sheet, j);
                    var measStr = EpplusOperation.GetCellValue(sheet, j, MeasIndex);

                    //If meas column not blank, but not belong to any pattern, flag error
                    if (patRow == null && !string.IsNullOrEmpty(measStr))
                    {
                        EpplusErrorManager.AddError(HardIpErrorType.MisPatternForMeasurement, ErrorLevel.Error, sheet.Name, j, ErrorMsgMissPatWithMeasurement);
                        continue;
                    }

                    //Will ignore the rows that Pattern column, Force Condition column, RegisterAssignment and Meas column are blank
                    if (string.IsNullOrEmpty(measStr))
                    {
                        var registerAssignment = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, j, RegisterIndex));
                        //read RegisterAssignment
                        if (patRow != null)
                        {
                            patRow.RegisterAssignment =
                                UpdatePatternItem(patRow.RegisterAssignment, registerAssignment);

                            var forceStr = EpplusOperation.GetCellValue(sheet, j, ForceIndex);
                            //Read post pattern force condition
                            patRow.PostPatForceCondition = UpdatePatternItem(patRow.PostPatForceCondition, forceStr);
                        }
                        continue;
                    }

                    ReadNonPatternRowNew(sheet, ref j, patRow);
                }
            }

            #endregion

            //DividePatternRow(planSheet.PatternRows);
            return planSheet;
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
                    EpplusErrorManager.AddError(HardIpErrorType.UnrecognisedHeader, ErrorLevel.Warning, Sheet.Name, 1,
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
                    limit.LoHeaderIndex =
                        ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, job + HardIpPattern.LoHeader, false);
                    if (limit.LoHeaderIndex == 1)
                        limit.LoHeaderIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder,
                            JobMap[job] + HardIpPattern.LoHeader);
                    limit.HiHeaderIndex =
                        ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, job + HardIpPattern.HiHeader, false);
                    if (limit.HiHeaderIndex == 1)
                        limit.HiHeaderIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder,
                            JobMap[job] + HardIpPattern.HiHeader);
                }
                else
                {
                    limit.LoHeaderIndex =
                        ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, job + HardIpPattern.LoHeader);
                    limit.HiHeaderIndex =
                        ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, job + HardIpPattern.HiHeader);
                }

                MeasLimits.Add(limit);
            }

            TtrIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.TtrHeader);
            var headerIndex = HeaderOrder.FirstOrDefault(a =>
                Regex.IsMatch(a.Key, HardIpPattern.NoBinOutHeader, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(a.Key, "Pattern Release Status")).Value;
            if (headerIndex > 0)
                NoBinOutIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.NoBinOutHeader);
            DescriptionIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.DescriptionHeader);
            PatternIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.PatternHeader);
            ForceIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.ForceConditionHeader);
            ForceCharIndex =
                ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.ForceConditionCharHeader, false);
            AnalogSetupIndex =
                ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.AnalogSetupHeader, false);
            TestNameIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.TestNameHeader, false);
            RegisterIndex =
                ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.RegisterAssignmentHeader);
            MiscInfoIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.MiscInfoHeader);
            MeasIndex = ExcelUtility.GetHeaderIndex(Sheet.Name, HeaderOrder, HardIpPattern.MeasHeader);
        }

        protected string GetSheetForceConditions()
        {
            var forceStr = "";
            for (var j = StartRow; j <= EndRow; j++)
            {
                var patternName = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(Sheet, j, PatternIndex));
                if (patternName != "")
                    break;
                var forceCondition = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(Sheet, j, ForceIndex));
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
                SearchInfo.TrimSpace(EpplusOperation.GetCellValue(Sheet, rowNum, DescriptionIndex));
            testPlanRow.ForceCondition =
                SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(Sheet, rowNum, ForceIndex));
            testPlanRow.ForceConditionChar =
                SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(Sheet, rowNum, ForceCharIndex));
            testPlanRow.MiscInfo =
                SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(Sheet, rowNum, MiscInfoIndex));
            testPlanRow.RegisterAssignment =
                SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(Sheet, rowNum, RegisterIndex));
            if (TestNameIndex != 1)
                testPlanRow.TestName = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(Sheet, rowNum, TestNameIndex));
            if (Sheet.MergedCells[rowNum, MeasIndex] != null)
            {
                testPlanRow.MergeRowNumForMeas = new ExcelAddress(Sheet.MergedCells[rowNum, MeasIndex]).Start.Row;
                testPlanRow.Meas = EpplusOperation.GetCellValue(Sheet, testPlanRow.MergeRowNumForMeas, MeasIndex)
                    .Replace("\n", "").Replace("\t", "");
            }
            else
            {
                testPlanRow.MergeRowNumForMeas = 0;
                testPlanRow.Meas = EpplusOperation.GetCellValue(Sheet, rowNum, MeasIndex).Replace("\n", "")
                    .Replace("\t", "");
            }

            foreach (var limit in MeasLimits)
            {
                var newLimit = new MeasLimit(limit.JobName);
                newLimit.LoLimit = EpplusOperation.GetCellValue(Sheet, rowNum, limit.LoHeaderIndex).Replace(" ", "");
                newLimit.HiLimit = EpplusOperation.GetCellValue(Sheet, rowNum, limit.HiHeaderIndex).Replace(" ", "");
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

            if (HardIpDataMain.TestPlanData != null)
                HardIpDataMain.TestPlanData.PlanHeaderIdx.Add(testPlanSheet.SheetName, idxDic);
        }

        protected void ReadPatternRow(ExcelWorksheet sheet, int rowIndex, PatternRow patRow, string patternName)
        {
            patRow.PatternColumnNum = PatternIndex;
            patRow.SheetName = sheet.Name;
            patRow.RowNum = rowIndex;
            patRow.TtrStr = SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(sheet, rowIndex, TtrIndex));
            patRow.NoBinOutStr =
                SearchInfo.TrimSpace(EpplusOperation.GetMergedCellValue(sheet, rowIndex, NoBinOutIndex));
            patRow.Description = EpplusOperation.GetCellValue(sheet, rowIndex, DescriptionIndex);
            patRow.Pattern = new PatternClass(patternName);

            patRow.ForceCondition.ForceCondition =
                SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, ForceIndex));
            if (ForceCharIndex != 1)
                patRow.ForceConditionChar =
                    SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, ForceCharIndex));
            if (AnalogSetupIndex != 1)
                patRow.AnalogSetup =
                    SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, AnalogSetupIndex));
            if (TestNameIndex != 1)
                patRow.SpecifyTestName =
                    SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, TestNameIndex));
            patRow.RegisterAssignment =
                SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, RegisterIndex));
            patRow.MiscInfo = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, MiscInfoIndex));
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
                Response.Report(message, MessageLevel.Error, 100);
            }
        }

        protected void CheckRunPattern(ExcelWorksheet sheet, int rowIndex)
        {
            var description = SearchInfo.TrimSpace(EpplusOperation.GetCellValue(sheet, rowIndex, DescriptionIndex));
            //If "Run the pattern" exist Description column but pattern column is blank, flag error
            if (Regex.IsMatch(description, RegRunPattern, RegexOptions.IgnoreCase))
                EpplusErrorManager.AddError(HardIpErrorType.MissingPatternInTestPlan, ErrorLevel.Error, sheet.Name,
                    rowIndex, ErrorMsgMissPat);
        }

        protected string UpdatePatternItem(string origin, string value)
        {
            var result = origin;
            if (!string.IsNullOrEmpty(value))
                result += ";" + value;
            return result.Trim(';');
        }
    }
}