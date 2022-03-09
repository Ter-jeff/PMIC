using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PmicAutomation.Utility.TCMID.DataStructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CompareStatus = PmicAutomation.Utility.TCMIDComparator.DataStructure.EnumStore.CompareStatus;
using LIMIT_FILE_TYPE = PmicAutomation.Utility.TCMID.DataStructure.EnumStore.LIMIT_FILE_TYPE;

namespace PmicAutomation.Utility.TCMID.Business
{
    public class TcmIDGenBase
    {
        protected System.Data.DataTable _limitDT;
        protected Dictionary<string, int> _dicGroupIndex;
        //protected Dictionary<string, Dictionary<string, int>> _dicTNameKeyValue;
        protected string _inputFile;
        protected string _outputPath;
        protected string _tpVersion;
        protected int _idxFlowtable;
        protected int _idxTestname;
        protected int _idxPETname;
        protected int _idxScale;
        protected int _idxUnits;
        protected int _idxLowlim;
        protected int _idxHilim;
        protected LIMIT_FILE_TYPE _limit_file_type;
        protected List<TcmIdEntry> _tcmIdList;
        private List<string> _skipLines;

        private int _iMegaCell = 1;
        private int _iTcmLabel = 2;
        private int _iInstance = 3;
        private int _iTcmID = 4;
        private int _iParameter = 5;
        private int _iVbatt = 6;
        private int _iTpRev = 7;
        private int _iTestType = 8;
        private int _iTestname = 9;
        private int _iTestDescription = 10;
        private int _iParameterType = 11;
        private int _iUnits = 12;
        private int _iDesignLowLimit = 13;
        private int _iDesignUpLimit = 14;
        private int _iATELowLimit = 15;
        private int _iATEUpLimit = 16;

        protected TcmIDGenBase()
        {
            _dicGroupIndex = new Dictionary<string, int>();
            //_dicTNameKeyValue = new Dictionary<string, Dictionary<string, int>>();
        }

        public Dictionary<string, int> DicGroupIndex
        {
            get { return _dicGroupIndex; }
            set { _dicGroupIndex = value; }
        }

        public List<string> skipLines
        {
            get { return _skipLines; }
            set { _skipLines = value; }
        }

        //public Dictionary<string, Dictionary<string, int>> DicTNameKeyValue
        //{
        //    get { return _dicTNameKeyValue; }
        //    set { _dicTNameKeyValue = value; }
        //}

        public List<TcmIdEntry> TcmIdList
        {
            get { return _tcmIdList; }
            set { _tcmIdList = value; }
        }

        public void SetParameter(LIMIT_FILE_TYPE limit_file_type, System.Data.DataTable limitDT, string inputFile, string outputPath, Dictionary<string, int> dicHeaderIndex, string tpVersion, List<string> skipLines)
        {
            _limit_file_type = limit_file_type;
            _limitDT = limitDT;
            _inputFile = inputFile;
            _outputPath = outputPath;
            _tpVersion = tpVersion;
            _idxFlowtable = dicHeaderIndex["idxFlowtable"];
            _idxTestname = dicHeaderIndex["idxTestname"];
            _idxPETname = dicHeaderIndex["idxPETname"];
            _idxScale = dicHeaderIndex["idxScale"];
            _idxUnits = dicHeaderIndex["idxUnits"];
            _idxLowlim = dicHeaderIndex["idxLowlim"];
            _idxHilim = dicHeaderIndex["idxHilim"];
            _skipLines = skipLines;
        }

        public virtual void Gen(bool bCompare = false, bool bGenFlag = true)
        {
            List<DataRow> targetList = SortAndFilter();
            _tcmIdList = GenTCMID(targetList);
            if (bGenFlag)
            {
                GenCompareReport(_tcmIdList, bCompare);
                GenLimitSheet(_tcmIdList, bCompare);
            }
        }

        private string GetFirstTwoToken(string name)
        {
            var tokens = name.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Count() == 1)
                return tokens.First();
            else
            {
                string retStr = string.Join("_", tokens[0], tokens[1]);
                if (Regex.IsMatch(retStr, @"_LimitSheet", RegexOptions.IgnoreCase))
                    return tokens[0];
                else
                    return retStr;
            }
        }

        public void GenLimitSheet(List<TcmIdEntry> tcmIdList, bool bCompare)
        {
            string sheetName = Path.GetFileNameWithoutExtension(_inputFile);
            string newOutputPath = Path.Combine(_outputPath, GetFirstTwoToken(sheetName));
            string xlsxFile = Path.Combine(newOutputPath, sheetName + "_new.xlsx");
            string txtFile = Path.Combine(newOutputPath, sheetName + "_new.txt");
            if (!Directory.Exists(newOutputPath))
                Directory.CreateDirectory(newOutputPath);
            if (File.Exists(xlsxFile))
                File.Delete(xlsxFile);
            if (File.Exists(txtFile))
                File.Delete(txtFile);

            List<DataRow> dtList = _limitDT.Rows.Cast<DataRow>().ToList();

            // remove PETname old data
            dtList.FindAll(s => !s[_idxPETname].ToString().ToUpper().Equals("PETNAME")).ForEach(p => p[_idxPETname] = string.Empty);

            foreach (var item in tcmIdList)
            {
                var target = dtList.First(p => p[_idxFlowtable].ToString().Equals(item.Flowtable) && p[_idxTestname].ToString().Equals(item.Testname));
                target[_idxPETname] = GetPETname(target[_idxTestname].ToString(), item.TcmId);
            }

            // used for txt
            Application lExcelApp = new Application();
            lExcelApp.DisplayAlerts = false;
            Workbooks lWks = lExcelApp.Workbooks;
            Workbook lWk = null;

            try
            {
                var excel = new ExcelPackage(new FileInfo(xlsxFile));
                var workbook = excel.Workbook;
                ExcelWorksheet workSheet = workbook.Worksheets.Add(sheetName);
                workSheet.Cells["A1"].LoadFromDataTable(_limitDT, false);

                workSheet.Cells.AutoFitColumns();
                excel.Save();
                excel.Dispose();

                lWk = lWks.Open(xlsxFile);
                lWk.SaveAs(txtFile, XlFileFormat.xlUnicodeText);
            }
            catch (Exception e)
            {
                throw new Exception("Write report failed: " + e.Message);
            }
            finally
            {
                if (lWk != null)
                {
                    lWk.Close(false);
                    Marshal.ReleaseComObject(lWk);
                }
                lWk = null;
                Marshal.ReleaseComObject(lWks);
                lWks = null;
                lExcelApp.Quit();
                Marshal.ReleaseComObject(lExcelApp);
                lExcelApp = null;
                GC.Collect();
            }
        }

        private string GetPETname(string source, string target)
        {
            if (string.IsNullOrEmpty(target))
                //return source;
                return string.Empty;
            string[] result = source.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (result.Length < 14)
                throw new Exception("Testname length could not less than 14: " + source);
            result[12] = target;
            return string.Join("_", result);
        }

        private void WriteHeader(ExcelWorksheet workSheet)
        {
            for (int c = 1; c <= 12; ++c)
            {
                workSheet.Cells[1, c, 2, c].Merge = true;
                workSheet.Cells[1, c].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[1, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            for (int c = 13; c <= 21; c += 2)
            {
                workSheet.Cells[1, c, 1, c + 1].Merge = true;
                workSheet.Cells[1, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            for (int c = 23; c <= 29; c += 3)
            {
                workSheet.Cells[1, c, 1, c + 2].Merge = true;
                workSheet.Cells[1, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            workSheet.Cells[1, 32, 1, 38].Merge = true;
            workSheet.Cells[1, 32].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells[1, 39, 1, 42].Merge = true;
            workSheet.Cells[1, 39].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            workSheet.Cells[1, _iMegaCell].Value = "MEGA CELL";
            workSheet.Cells[1, _iTcmLabel].Value = "TCM Label";
            workSheet.Cells[1, _iInstance].Value = "INSTANCE";
            workSheet.Cells[1, _iTcmID].Value = "TCMID";
            workSheet.Cells[1, _iParameter].Value = "Parameter";
            workSheet.Cells[1, _iVbatt].Value = "Vbatt";
            workSheet.Cells[1, _iTpRev].Value = "TP_REV";
            workSheet.Cells[1, _iTestType].Value = "Test Type";
            workSheet.Cells[1, _iTestname].Value = "Test Name";
            workSheet.Cells[1, _iTestDescription].Value = "Test Description";
            workSheet.Cells[1, _iParameterType].Value = "Parameter Type";
            workSheet.Cells[1, _iUnits].Value = "Units";

            workSheet.Cells[1, 13].Value = "Design Limits";
            workSheet.Cells[2, _iDesignLowLimit].Value = "Lower Limit";
            workSheet.Cells[2, _iDesignUpLimit].Value = "Upper Limit";

            workSheet.Cells[1, 15].Value = "ATE Limits";
            workSheet.Cells[2, _iATELowLimit].Value = "Lower Limit";
            workSheet.Cells[2, _iATEUpLimit].Value = "Upper Limit";

            workSheet.Cells[1, 17].Value = "Production Limits";
            workSheet.Cells[2, 17].Value = "Lower Limit";
            workSheet.Cells[2, 18].Value = "Upper Limit";

            workSheet.Cells[1, 19].Value = "CHAR Limits";
            workSheet.Cells[2, 19].Value = "Lower Limit";
            workSheet.Cells[2, 20].Value = "Upper Limit";

            workSheet.Cells[1, 21].Value = "QUAL Limits";
            workSheet.Cells[2, 21].Value = "Lower Limit";
            workSheet.Cells[2, 22].Value = "Upper Limit";

            workSheet.Cells[1, 23].Value = "HTOL_GB";
            workSheet.Cells[2, 23].Value = "Lower";
            workSheet.Cells[2, 24].Value = "Upper";
            workSheet.Cells[2, 25].Value = "CRTL_FLG";

            workSheet.Cells[1, 26].Value = "TEMP_GB";
            workSheet.Cells[2, 26].Value = "Lower";
            workSheet.Cells[2, 27].Value = "Upper";
            workSheet.Cells[2, 28].Value = "CRTL_FLG";

            workSheet.Cells[1, 29].Value = "GRR_GB";
            workSheet.Cells[2, 29].Value = "Lower";
            workSheet.Cells[2, 30].Value = "Upper";
            workSheet.Cells[2, 31].Value = "CRTL_FLG";

            workSheet.Cells[1, 32].Value = "Measurements";
            workSheet.Cells[2, 32].Value = "Mean";
            workSheet.Cells[2, 33].Value = "Median";
            workSheet.Cells[2, 34].Value = "Max";
            workSheet.Cells[2, 35].Value = "Min";
            workSheet.Cells[2, 36].Value = "Std Dev";
            workSheet.Cells[2, 37].Value = "Qhi";
            workSheet.Cells[2, 38].Value = "Qlo";

            workSheet.Cells[1, 39].Value = "Process Capability";
            workSheet.Cells[2, 39].Value = "Cpu";
            workSheet.Cells[2, 40].Value = "Cpl";
            workSheet.Cells[2, 41].Value = "Cpk";
            workSheet.Cells[2, 42].Value = "Cpk_ctrl";
        }

        private string GetInstanceName(string testname)
        {
            string[] tokens = testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            string group2 = tokens[1].ToUpper().Trim();
            string group1 = tokens[0].ToUpper().Trim();
            if (group1.Equals("LDO"))
                return "IQ";
            else if (group1.Equals("LDO0"))
                return "LDOINT";
            else if (group1.Equals("BANDGAP") && group2.Equals("FBG-CLK"))
                return "RTC";
            else if (group1.Equals("BANDGAP") && group2.Equals("IABS"))
                return "IBAT";
            else if (group1.Equals("BANDGAP"))
                return "VrefIref";
            else if (group1.Equals("COMP"))
                return "VoltCompare";
            else if (group1.Equals("ADC"))
                return "GPADC";
            else if (group1.StartsWith("BUCKSW"))
            {
                string digit = Regex.Match(group1, @"(?<val>\d+)").Groups["val"].ToString();
                return "SWITCH" + digit;
            }
            else if (group1.Contains("-"))
                group1 = group1.Substring(0, group1.IndexOf("-"));
            return group1;
        }

        private string GetVbatt(string testname)
        {
            string[] tokens = testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens[8].ToUpper().Contains("VDD"))
                return tokens[8].ToUpper();
            else
                return string.Empty;
        }

        protected virtual string GetTestType(string testname)
        {
            string[] tokens = testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length <= 6)
                return testname;
            else
                return tokens[6];
        }

        private void WriteBody(ExcelWorksheet workSheet, TcmIdEntry item, int rowIndex)
        {
            workSheet.Cells[rowIndex, _iTcmLabel].Value = "CP1";
            workSheet.Cells[rowIndex, _iInstance].Value = GetInstanceName(item.Testname);
            workSheet.Cells[rowIndex, _iTcmID].Value = item.TcmId;
            workSheet.Cells[rowIndex, _iVbatt].Value = GetVbatt(item.Testname);
            workSheet.Cells[rowIndex, _iTpRev].Value = _tpVersion;
            workSheet.Cells[rowIndex, _iTestType].Value = GetTestType(item.Testname);
            workSheet.Cells[rowIndex, _iTestname].Value = item.Testname;
            workSheet.Cells[rowIndex, _iUnits].Value = string.Concat(item.Scale.Equals("9999") ? "" : item.Scale, item.Units);
            workSheet.Cells[rowIndex, _iATELowLimit].Formula = item.LowLim.Replace("=", "");
            workSheet.Cells[rowIndex, _iATELowLimit].Calculate();
            workSheet.Cells[rowIndex, _iATEUpLimit].Formula = item.HiLim.Replace("=", "");
            workSheet.Cells[rowIndex, _iATEUpLimit].Calculate();
        }

        public void GenDiffReport(List<Tuple<TcmIdEntry, TcmIdEntry>> reportList)
        {
            const int iTestname = 1;
            const int iStatus = 2;
            const int iTcmId = 3;
            const int iLolim = 4;
            const int iHilim = 5;

            string sheetName = Path.GetFileNameWithoutExtension(_inputFile);
            string newOutputPath = Path.Combine(_outputPath, GetFirstTwoToken(sheetName));
            string file = Path.Combine(newOutputPath, sheetName + "_Difference_Report.xlsx");
            if (!Directory.Exists(newOutputPath))
                Directory.CreateDirectory(newOutputPath);
            if (File.Exists(file))
                File.Delete(file);

            try
            {
                var excel = new ExcelPackage(new FileInfo(file));
                var workbook = excel.Workbook;
                ExcelWorksheet workSheet = workbook.Worksheets.Add("DiffReport");

                for (int c = 1; c <= 5; ++c)
                {
                    workSheet.Cells[1, c].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    workSheet.Cells[1, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                workSheet.Cells[1, iTestname].Value = "testname";
                workSheet.Cells[1, iStatus].Value = "status";
                workSheet.Cells[1, iTcmId].Value = "TCMID";
                workSheet.Cells[1, iLolim].Value = "lolim";
                workSheet.Cells[1, iHilim].Value = "hilim";

                int iRow = 2;
                foreach (var obj in reportList)
                {
                    // item may be null
                    TcmIdEntry baseEntry = obj.Item1;
                    TcmIdEntry newEntry = obj.Item2;
                    TcmIdEntry displayEntry = null;

                    if (newEntry == null) // status: remove
                        displayEntry = baseEntry;
                    else
                        displayEntry = newEntry;

                    if (!displayEntry.Status.Equals(CompareStatus.TCMID_REMOVE) && string.IsNullOrEmpty(displayEntry.TcmId))
                        continue;

                    workSheet.Cells[iRow, iTestname].Value = displayEntry.Testname;
                    workSheet.Cells[iRow, iStatus].Value = displayEntry.Status;

                    workSheet.Cells[iRow, iTcmId].Value = displayEntry.OriginalTcmId;
                    if (newEntry != null && baseEntry != null && !newEntry.OriginalTcmId.ToUpper().Equals(baseEntry.TcmId.ToUpper()))
                    {
                        workSheet.Cells[iRow, iTcmId].AddComment(baseEntry.TcmId, "Teradyne");
                        SetCellColor(workSheet, iRow, iTcmId);
                    }

                    //workSheet.Cells[iRow, iLolim].Value = displayEntry.LowLim;
                    workSheet.Cells[iRow, iLolim].Formula = displayEntry.LowLim.Replace("=", "");
                    workSheet.Cells[iRow, iLolim].Calculate();
                    if (newEntry != null && baseEntry != null && !newEntry.LowLim.Equals(baseEntry.LowLim))
                    {
                        workSheet.Cells[iRow, iLolim].AddComment(baseEntry.LowLim, "Teradyne");
                        SetCellColor(workSheet, iRow, iLolim);
                    }

                    //workSheet.Cells[iRow, iHilim].Value = displayEntry.HiLim;
                    workSheet.Cells[iRow, iHilim].Formula = displayEntry.HiLim.Replace("=", "");
                    workSheet.Cells[iRow, iHilim].Calculate();
                    if (newEntry != null && baseEntry != null && !newEntry.HiLim.Equals(baseEntry.HiLim))
                    {
                        workSheet.Cells[iRow, iHilim].AddComment(baseEntry.HiLim, "Teradyne");
                        SetCellColor(workSheet, iRow, iHilim);
                    }

                    ++iRow;
                }

                workSheet.Cells.AutoFitColumns();
                excel.Save();
                excel.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Write difference report failed: " + e.Message);
            }
        }

        private void SetCellColor(ExcelWorksheet workSheet, int row, int col, bool bBold = false)
        {
            workSheet.Cells[row, col].Style.Font.Bold = bBold;
            workSheet.Cells[row, col].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }

        public string GenCompareReport(List<TcmIdEntry> tcmIdList, bool bCompare)
        {
            string file = string.Empty;
            string sheetName = Path.GetFileNameWithoutExtension(_inputFile);
            string newOutputPath = Path.Combine(_outputPath, GetFirstTwoToken(sheetName));
            if (bCompare)
                file = Path.Combine(newOutputPath, sheetName + "_TCMID_Report.xlsx");
            else
                file = Path.Combine(newOutputPath, sheetName + "_Report.xlsx");
            if (!Directory.Exists(newOutputPath))
                Directory.CreateDirectory(newOutputPath);
            if (File.Exists(file))
                File.Delete(file);

            try
            {
                var excel = new ExcelPackage(new FileInfo(file));
                var workbook = excel.Workbook;

                // report sheet name based on testname first token
                List<string> blockList = tcmIdList.Select(p => p.Testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries).First().ToUpper()).Distinct().ToList();

                foreach (string block in blockList)
                {
                    ExcelWorksheet workSheet = workbook.Worksheets.Add(block);
                    WriteHeader(workSheet);
                    int rowIndex = 3;
                    foreach (TcmIdEntry item in tcmIdList.FindAll(p => p.Testname.StartsWith(block + "_", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (string.IsNullOrEmpty(item.TcmId))
                            continue;
                        WriteBody(workSheet, item, rowIndex++);
                    }
                    workSheet.Cells.AutoFitColumns();
                }

                excel.Save();
                excel.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Write report failed: " + e.Message);
            }

            return file;
        }

        protected List<TcmIdEntry> GenTCMID(List<DataRow> targetList)
        {
            List<TcmIdEntry> resultList = new List<TcmIdEntry>();
            foreach (var item in targetList)
            {
                string flowtable = item[_idxFlowtable].ToString();
                string testname = item[_idxTestname].ToString();
                string scale = item[_idxScale].ToString();
                string units = item[_idxUnits].ToString();
                string lowlim = item[_idxLowlim].ToString();
                string hilim = item[_idxHilim].ToString();
                string tcmId = string.Empty;
                if (!_skipLines.Exists(p => testname.IndexOf(p, StringComparison.CurrentCultureIgnoreCase) != -1))
                    tcmId = FetchTcmID(testname);
                //if (string.IsNullOrEmpty(tcmId))
                //    continue;
                resultList.Add(new TcmIdEntry(flowtable, testname, tcmId, scale, units, lowlim, hilim));
            }
            return resultList;
        }

        // used for Others, format: PULL.00.P0001
        protected virtual string FetchTcmID(string testname)
        {
            if (string.IsNullOrEmpty(testname.Trim()))
                return "N/A";

            string id = string.Empty;
            string digit = string.Empty;
            string[] tokens = testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length < 8)
                return "Error";

            string group3 = tokens[2].Trim().ToUpper();
            string group2 = tokens[1].Trim().ToUpper();
            string group1 = tokens[0].Trim().ToUpper();
            if (Regex.IsMatch(group1, @"\d+", RegexOptions.IgnoreCase))
            {
                digit = Regex.Match(group1, @"(?<val>\d+)").Groups["val"].ToString();
                id = group1.Replace(digit, "") + "." + string.Format("{0:D2}", Convert.ToInt32(digit));
            }
            else
            {
                digit = "0";
                id = group1 + ".00";
            }

            if (group1.Equals("LDO"))
                id = "IQ.00";
            else if (group1.Equals("LDO0"))
                id = "LDO.INT";
            else if (group1.Equals("BANDGAP") && group2.Equals("FBG-CLK"))
                id = "RTC.00";
            else if (group1.Equals("BANDGAP") && group2.Equals("IABS"))
                id = "IBAT.00";
            else if (group1.Equals("BANDGAP"))
                id = "VrefIref.00";
            else if (group1.Equals("COMP"))
                id = "VoltCompare.00";
            else if (group1.Equals("ADC"))
                id = "GPADC.00";
            else if (group1.StartsWith("BUCKSW"))
                id = "SWITCH"+ "." + string.Format("{0:D2}", Convert.ToInt32(digit));

            string group9 = tokens[8].Trim().ToUpper();
            string group8 = tokens[7].Trim().ToUpper();
            string group7 = tokens[6].Trim().ToUpper();
            if (group8.Equals("X"))
            {
                if (group7.Equals("P"))
                    group8 = "P";
                else if (group7.Equals("SWP5"))
                    group8 = "P";
                else
                    group8 = string.Empty;
            }
            else if (group8.Equals("C") || group8.Equals("T"))
            {
                if (group7.Equals("X"))
                    group8 = string.Empty;
                else if (group7.Equals("DIFFTRIMCODE"))
                    group8 = string.Empty;
                else if (Regex.IsMatch(group7, @"FINALTRIMCODE|PRETRIMCODE|FIRSTTRIMCODE|POSTBURNCODE"))
                    group8 = "C";
                else if (Regex.IsMatch(group7, @"FINALTRIM|PRETRIM|FIRSTTRIM|POSTBURN"))
                    group8 = "T";
                else
                    group8 = string.Empty;
            }

            if (group7.Equals("P") && group9.Equals("VDDXV"))
                group8 = "P";

            if (string.IsNullOrEmpty(group8) && Regex.IsMatch(_inputFile, @"BUCK2p3p5p", RegexOptions.IgnoreCase))
            {
                string testType = GetTestType(testname);
                if (Regex.IsMatch(testType, @"GNG$", RegexOptions.IgnoreCase))
                {
                    if (Regex.IsMatch(testname, @"_SlopeComp_|_REF_Ton", RegexOptions.IgnoreCase))
                        group8 = "T";
                }
            }

            if (string.IsNullOrEmpty(group8))
                return string.Empty;

            PutEntryInDic(group8 + digit);
            id += "." + group8 + string.Format("{0:D4}", _dicGroupIndex[group8 + digit]);

            return id.ToUpper();
        }

        // used for Conti/IDS/Leakage
        protected string FetchTcmID(string testname, string group8)
        {
            if (string.IsNullOrEmpty(testname.Trim()))
                return "N/A";

            string id = string.Empty;
            string digit = string.Empty;
            string[] tokens = testname.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length < 8)
                return "Error";

            if (tokens[0].Contains("-"))
                tokens[0] = tokens[0].Substring(0, tokens[0].IndexOf("-"));

            if (Regex.IsMatch(tokens[0], @"\d+", RegexOptions.IgnoreCase))
            {
                digit = Regex.Match(tokens[0], @"(?<val>\d+)").Groups["val"].ToString();
                id = tokens[0].Replace(digit, "") + "." + string.Format("{0:D2}", Convert.ToInt32(digit));
            }
            else
            {
                id = tokens[0].Trim() + ".00";
            }

            PutEntryInDic(tokens[0] + digit);
            id += "." + group8 + string.Format("{0:D4}", _dicGroupIndex[tokens[0] + digit]);

            return id.ToUpper();
        }

        protected void PutEntryInDic(string key)
        {
            if (!_dicGroupIndex.ContainsKey(key))
                _dicGroupIndex.Add(key, 1);
            else
            {
                int tmp = _dicGroupIndex[key];
                _dicGroupIndex.Remove(key);
                _dicGroupIndex.Add(key, ++tmp);
            }
        }

        // sorting based on column flowtable by _TRIM & _PostBurn
        protected List<DataRow> SortAndFilter()
        {
            IEnumerable<DataRow> collection = _limitDT.Rows.Cast<DataRow>();
            List<DataRow> targetList = collection.ToList().OrderBy(s => s[_idxFlowtable].ToString(), StringComparer.OrdinalIgnoreCase)
                .Where(s => !(Regex.IsMatch(s[_idxLowlim].ToString(), @"N/A", RegexOptions.IgnoreCase) || Regex.IsMatch(s[_idxHilim].ToString(), @"N/A", RegexOptions.IgnoreCase)))
                .Where(s => !Regex.IsMatch(s[_idxFlowtable].ToString(), @"_HV$|_LV$|_ULV$|_UHV$", RegexOptions.IgnoreCase))
                .Where(s => Regex.IsMatch(s[_idxTestname].ToString(), @"NV_|xV_|3\.80V_", RegexOptions.IgnoreCase)).ToList();
            return targetList;
        }
    }
}
