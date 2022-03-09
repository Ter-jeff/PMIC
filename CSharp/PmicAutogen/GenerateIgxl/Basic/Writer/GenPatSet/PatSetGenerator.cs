using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others.PatternListCsvFile;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenPatSet
{
    public class PatSetGenerator
    {
        private const string DefaultBurst = "NO";
        private const string UsedType = "USE";
        private const string PatPath = @".\Pattern\";
        private const string PatSuffix = ".PAT.GZ";
        private readonly Dictionary<string, string> _fileListInAllFolder = new Dictionary<string, string>();
        private Dictionary<string, SubPatInfo> _hardIpInfoAllDic = new Dictionary<string, SubPatInfo>();
        private PatSetSheet _patSetSheetAllNoPat;
        private PatSetSubSheet _patSubSheetAllNoPat;
        private string _patternPath;
        public Dictionary<string, PatSetStatus> PatSetStatus = new Dictionary<string, PatSetStatus>();
        public PatSetSheet PatSetSheetAll { get; set; }
        public PatSetSubSheet PatSubSheetAll { get; set; }

        #region Convertor Flow

        public void GenerateFlow(List<PatternListCsvRow> patternListCsvRows, string patternPath)
        {
            _patternPath = patternPath;
            GetPatFileDirAndPatInfo();
            PatSetSheetAll = new PatSetSheet(PmicConst.PatSetsAll);
            PatSubSheetAll = new PatSetSubSheet(PmicConst.PatternSubRoutine);
            _patSetSheetAllNoPat = new PatSetSheet("patSetSheetAllNoPat"); //temporary storage
            _patSubSheetAllNoPat = new PatSetSubSheet("patSubSheetAllNoPat"); //temporary storage
            foreach (var patternListCsvRow in patternListCsvRows)
                if (patternListCsvRow.Use.ToUpper().Equals(UsedType))
                {
                    GenPatSetData(patternListCsvRow, false);
                }
                else
                {
                    var patSetStatus = new PatSetStatus {Used = "Don't Use"};
                    if (!PatSetStatus.ContainsKey(patternListCsvRow.PatternName))
                        PatSetStatus.Add(patternListCsvRow.PatternName, patSetStatus);
                }

            var dummyPat = new PatSet {PatSetName = ""};
            dummyPat.AddRow(new PatSetRow {File = ""});
            PatSetSheetAll.AddPatSet(dummyPat);
            PatSetSheetAll.AddPatSet(dummyPat);

            foreach (var item in _patSetSheetAllNoPat.PatSetRows)
                PatSetSheetAll.AddPatSet(item);

            if (_patSubSheetAllNoPat.PatSetSubData.Count > 0)
            {
                var dummyPatSub = new PatSetSubRow {PatternFileName = ""};
                PatSubSheetAll.AddRow(dummyPatSub);
                PatSubSheetAll.AddRow(dummyPatSub);
                foreach (var item in _patSubSheetAllNoPat.PatSetSubData) PatSubSheetAll.AddRow(item);
            }

            if (PatSubSheetAll.PatSetSubData.Count == 0)
                PatSubSheetAll = null;
        }

        private void GenPatSetData(PatternListCsvRow patternListCsvRow, bool skipPatternExistCheck)
        {
            string baseFilePath;
            var fileValue = CreateFileValue(patternListCsvRow.FileVersion, out baseFilePath);
            var patSetStatus = new PatSetStatus();
            patSetStatus.Used = "Used";
            baseFilePath = baseFilePath.Trim().Replace('/', '\\').TrimEnd('\\');

            // if file version equals to "NA" then do not generate it.
            if (patternListCsvRow.TimeSetVersion.Trim().ToUpper() == "NA" ||
                patternListCsvRow.FileVersion.Trim().ToUpper() == "NA" ||
                patternListCsvRow.TimeSetVersion.Trim().ToUpper() == "N/A" ||
                patternListCsvRow.FileVersion.Trim().ToUpper() == "N/A")
            {
                patSetStatus.ValidTs = "NoValidTs";
                PatSetStatus.Add(patternListCsvRow.PatternName, patSetStatus);
                return;
            }

            patSetStatus.ValidTs = "ValidTs";

            // Check if file exist
            var patExisted = false;
            var patternName = baseFilePath.ToUpper().Replace(".PAT.GZ", "");
            var containSub = false;
            var subNoVm = false;
            var patSetSubFileValues = new List<string>();
            if (!skipPatternExistCheck)
            {
                if (_fileListInAllFolder.ContainsKey(baseFilePath.ToUpper()))
                {
                    patExisted = true;
                    fileValue = _fileListInAllFolder[baseFilePath.ToUpper()];
                }
                else
                {
                    // create fake directory 
                    var patternObj = new PatternNameInfo(baseFilePath);
                    fileValue = @".\Pattern\" + patternObj.TpCategory + @"\" + fileValue;
                }

                // check if contain subroutine
                if (_hardIpInfoAllDic.ContainsKey(patternName))
                {
                    if (_hardIpInfoAllDic[patternName].Subroutine.Any())
                    {
                        containSub = true;
                        patSetStatus.ContainSub = "ContainSub";
                        foreach (var sub in _hardIpInfoAllDic[patternName].Subroutine)
                        {
                            var patSetSubFileValue = fileValue + ":" + sub;
                            patSetSubFileValue = patSetSubFileValue.Replace(_patternPath, PatPath).Replace(@"\\", @"\")
                                .Replace("/", "\\");
                            patSetSubFileValues.Add(patSetSubFileValue);
                        }
                    }

                    var vmVector = _hardIpInfoAllDic[patternName].VmVector;
                    if (!string.IsNullOrEmpty(vmVector))
                        fileValue = fileValue + ":" + vmVector;
                    else if (containSub) subNoVm = true;
                }

                if (!Regex.IsMatch(fileValue, "gz:|pat:", RegexOptions.IgnoreCase))
                    fileValue = fileValue + ":" +
                                fileValue.Split('\\').Last().ToUpper().Replace(".PAT", "")
                                    .Replace(".GZ",
                                        ""); // always put vm_vector name no matter subroutine existed or not. 2017/05/05 Osprey Team.
                fileValue = fileValue.Replace(_patternPath, PatPath).Replace(@"\\", @"\").Replace("/", "\\");
                if (!containSub) patSetStatus.ContainSub = "NonContainSub";

                patSetStatus.Existed = !patExisted ? "NonExisted" : "Existed";
            }
            else
            {
                patSetStatus.Existed = "Skipped";
            }

            PatSetStatus.Add(patternListCsvRow.PatternName, patSetStatus);

            // remove PatSet .gz
            var patSet = new PatSetRow
            {
                //PatternSet = pattern.PatternName.ToUpper(),
                Burst = DefaultBurst.ToUpper(), File = fileValue.ToUpper().Replace(".GZ", "")
            };

            var patSetSubRowList = new List<PatSetSubRow>();
            foreach (var patSetSubFileValue in patSetSubFileValues)
            {
                var patSetSubRow = new PatSetSubRow {PatternFileName = patSetSubFileValue.ToUpper().Replace(".GZ", "")};
                patSetSubRowList.Add(patSetSubRow);
            }

            if (PatSetSheetAll.GetPatSetCnt() == 0) patSet.Comment = LocalSpecs.PatListCsvFile.Split('\\').Last();
            if (PatSetSheetAll.GetPatSetCnt() == 1) patSet.Comment = LocalSpecs.ScghFileName.Split('\\').Last();
            if (PatSetSheetAll.GetPatSetCnt() == 2) patSet.Comment = LocalSpecs.TestPlanFileName.Split('\\').Last();

            var patSetItem = new PatSet();
            patSetItem.PatSetName = patternListCsvRow.PatternName.ToUpper();
            if (patSetStatus.Existed == "NonExisted")
                patSet.Comment = "NonExisted";

            patSetItem.AddRow(patSet);
            if (!subNoVm)
            {
                if (patSetStatus.Existed == "Existed" || patSetStatus.Existed == "Skipped")
                {
                    if (patternListCsvRow.Check.ToUpper() == "FAIL")
                    {
                        patSetItem.PatSetRows[0].Comment = patternListCsvRow.CheckComment;
                        _patSetSheetAllNoPat.AddPatSet(patSetItem);
                    }
                    else
                    {
                        PatSetSheetAll.AddPatSet(patSetItem);
                    }
                }
                else
                {
                    _patSetSheetAllNoPat.AddPatSet(patSetItem);
                }
            }


            if (!containSub) return;
            foreach (var patSetSubRow in patSetSubRowList)
            {
                if (patSetStatus.Existed == "NonExisted")
                    patSetSubRow.Comment = "NonExisted";

                if (PatSubSheetAll.GetPatSetSubCnt() == 1)
                    patSetSubRow.Comment = LocalSpecs.PatListCsvFile.Split('\\').Last();
                if (PatSubSheetAll.GetPatSetSubCnt() == 2)
                    patSetSubRow.Comment = LocalSpecs.ScghFileName.Split('\\').Last();
                if (PatSubSheetAll.GetPatSetSubCnt() == 3)
                    patSetSubRow.Comment = LocalSpecs.TestPlanFileName.Split('\\').Last();

                if (patSetStatus.Existed == "Existed" || patSetStatus.Existed == "Skipped")
                {
                    if (patternListCsvRow.Check.ToUpper() == "FAIL")
                    {
                        patSetSubRow.Comment = patternListCsvRow.CheckComment;
                        _patSubSheetAllNoPat.AddRow(patSetSubRow);
                    }
                    else
                    {
                        PatSubSheetAll.AddRow(patSetSubRow);
                    }
                }

                else
                {
                    _patSubSheetAllNoPat.AddRow(patSetSubRow);
                }
            }
        }

        private void GetPatFileDirAndPatInfo()
        {
            if (_fileListInAllFolder.Count == 0)
            {
                var files = Directory.GetFiles(_patternPath, "*.PAT.*", SearchOption.AllDirectories);
                foreach (var file in files)
                {
                    var fileName = file.Split('\\').Last().ToUpper();
                    if (!_fileListInAllFolder.ContainsKey(fileName))
                        _fileListInAllFolder.Add(fileName, file);
                }
            }

            if (_hardIpInfoAllDic.Count == 0)
            {
                var hardipFolder = _patternPath + @"\";
                _hardIpInfoAllDic = ReadHardIpInfoAll(Path.Combine(hardipFolder, "HardIP_AutoGen_Info_All.txt"));
            }
        }

        private Dictionary<string, SubPatInfo> ReadHardIpInfoAll(string hardipInfoFile)
        {
            var patternInfoAll = new Dictionary<string, SubPatInfo>();

            if (File.Exists(hardipInfoFile))
            {
                string line;
                var genericPatName = "";
                var version = "";
                var subPatInfo = new SubPatInfo();
                var file = new StreamReader(hardipInfoFile);
                while ((line = file.ReadLine()) != null)
                    if (line.IndexOf("GenericPat:=", StringComparison.Ordinal) != -1)
                    {
                        genericPatName = line.Substring(line.IndexOf(":=", StringComparison.Ordinal) + 2);
                    }
                    else if (line.IndexOf("Version:=", StringComparison.Ordinal) != -1)
                    {
                        version = line.Substring(line.IndexOf(":=", StringComparison.Ordinal) + 2);
                    }
                    else if (line.IndexOf("<HardIP_Info_Token>", StringComparison.Ordinal) != -1)
                    {
                        var fullPattern = genericPatName + "_" + version;
                        if (!patternInfoAll.ContainsKey(fullPattern.ToUpper()))
                            patternInfoAll.Add(fullPattern.ToUpper(), subPatInfo);
                        genericPatName = "";
                        version = "";
                        subPatInfo = new SubPatInfo();
                    }
                    else if (line != "")
                    {
                        if (line.IndexOf("Subr:", StringComparison.Ordinal) != -1)
                        {
                            var subList = line.Split(':')[1].Trim().Split(',');
                            foreach (var sub in subList)
                                subPatInfo.Subroutine.Add(sub);
                        }

                        if (line.IndexOf("VM_Vector:", StringComparison.Ordinal) != -1)
                        {
                            // "VM_Vector: A-Z | a-z | 0-9 | _
                            var m1 = new Regex(@"(VM_Vector:)((\s)+)(?<VM_vector>(([A-Z]|[a-z]|[0-9]|[_])+))");
                            var match = m1.Match(line);
                            if (match.Success) subPatInfo.VmVector = match.Groups["VM_vector"].Value;
                        }
                    }

                file.Close();
            }

            return patternInfoAll;
        }

        private string CreateFileValue(string fileVersion, out string baseFilePath)
        {
            var fileName = GetFileName(fileVersion);
            baseFilePath = fileName;
            return fileName;
        }

        private string GetFileName(string fileVersion)
        {
            fileVersion = fileVersion.Replace('/', '\\');
            var startIndex = fileVersion.LastIndexOf('\\');
            var fileName = fileVersion;
            if (startIndex > 0) fileName = fileVersion.Substring(startIndex + 1);

            var endIndex = fileName.IndexOf('.');
            if (endIndex > 0)
            {
                fileName = fileName.Substring(0, endIndex);
                fileName = fileName + PatSuffix;
            }

            if (fileName.Trim().ToUpper() == "NA" || fileName.Trim().ToUpper() == "N/A") fileName = string.Empty;

            return fileName;
        }

        #endregion
    }
}