using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class PatSetSheet : IgxlSheet
    {
        private const string SheetType = "DTPatternSetSheet";

        #region Field
        private List<PatSet> _patSets;
        private Dictionary<string, int> _patSetNameDictionary;
        #endregion

        #region Properity
        public List<PatSet> PatSetRows
        {
            set { _patSets = value; }
            get
            {
                return _patSets ?? (_patSets = new List<PatSet>());
            }
        }
        #endregion

        #region Constructor
        public PatSetSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _patSets = new List<PatSet>();
            IgxlSheetName = IgxlSheetNameList.PatternSet;
            _patSetNameDictionary = new Dictionary<string, int>();

        }

        public PatSetSheet(string sheetName)
            : base(sheetName)
        {
            _patSets = new List<PatSet>();
            IgxlSheetName = IgxlSheetNameList.PatternSet;
            _patSetNameDictionary = new Dictionary<string, int>();
        }
        #endregion

        #region Member Function
        public string GetPattenSetNameWithSeq(string patSetName)
        {
            if (_patSetNameDictionary.ContainsKey(patSetName))
            {
                _patSetNameDictionary[patSetName] = _patSetNameDictionary[patSetName] + 1;
                return patSetName + "_" + _patSetNameDictionary[patSetName];
            }
            return patSetName;
        }

        public bool IsExistTheSamePatSet(PatSet igxlItem)
        {
            if (!_patSets.Exists(x => x.PatSetName.Equals(igxlItem.PatSetName, StringComparison.OrdinalIgnoreCase)))
                return false;

            var row = _patSets.Find(x => x.PatSetName.Equals(igxlItem.PatSetName, StringComparison.OrdinalIgnoreCase));
            if (row.PatSetRows.Count != igxlItem.PatSetRows.Count)
                return false;
            for (var i = 0; i < igxlItem.PatSetRows.Count; i++)
            {
                if (!igxlItem.PatSetRows[i].File.Equals(row.PatSetRows[i].File, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        public void AddPatSet(PatSet igxlItem)
        {
            _patSets.Add(igxlItem);
            if (igxlItem.PatSetName == null) return;
            if (!_patSetNameDictionary.ContainsKey(igxlItem.PatSetName))
                _patSetNameDictionary.Add(igxlItem.PatSetName, 0);
        }

        public bool IsExist(PatSet patSet)
        {
            return PatSetRows.Exists(x => x.PatSetName.Equals(patSet.PatSetName, StringComparison.CurrentCultureIgnoreCase));
        }

        public bool IsExist(string patSetName)
        {
            return PatSetRows.Exists(x => x.PatSetName.Equals(patSetName, StringComparison.CurrentCultureIgnoreCase));
        }

        public void AddPatSets(List<PatSet> igxlItems)
        {
            foreach (var igxlItem in igxlItems)
            {
                _patSets.Add(igxlItem);
                if (igxlItem.PatSetName == null) continue;
                if (!_patSetNameDictionary.ContainsKey(igxlItem.PatSetName))
                    _patSetNameDictionary.Add(igxlItem.PatSetName, 0);
            }
        }

        public long GetPatSetCnt()
        {
            return _patSets.Count;
        }

        public override void Write(string fileName, string version = "2.2")
        {
            //var versionDouble = Double.Parse(version);
            //if (versionDouble > 2.2)
            //    versionDouble = 2.2;
            //var validate = new Action<string>((a) => { });
            //var patSetGen = new GenPatternSetSheet(fileName, validate, "", true, versionDouble);
            //var backuplist = _patSets.Where(x => x.IsBackup).ToList();
            //var mainlist = _patSets.Where(x => !x.IsBackup).ToList();

            //foreach (var patSet in mainlist)
            //{
            //    foreach (var patSetRow in patSet.PatSetRows)
            //    {
            //        patSetGen.AddRow(patSet.PatSetName, patSetRow.TdGroup, patSetRow.TimeDomain, patSetRow.Enable, patSetRow.File, patSetRow.Burst, patSetRow.StartLabel, patSetRow.StopLabel, patSetRow.Comment);
            //    }
            //}
            //if (backuplist.Count != 0)
            //    patSetGen.AddBlankLine();
            //foreach (var patSet in backuplist)
            //{
            //    foreach (var patSetRow in patSet.PatSetRows)
            //    {
            //        patSetGen.AddRow(patSet.PatSetName, patSetRow.TdGroup, patSetRow.TimeDomain, patSetRow.Enable, patSetRow.File, patSetRow.Burst, patSetRow.StartLabel, patSetRow.StopLabel, patSetRow.Comment);
            //    }
            //}
            //patSetGen.WriteSheet();

            //Support 2.2 & 2.3
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("2.3"))
                {
                    var igxlSheetsVersion = dic["2.3"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (PatSetRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var patternSetIndex = GetIndexFrom(igxlSheetsVersion, "Pattern Set");
                var tDGroupSetIndex = GetIndexFrom(igxlSheetsVersion, "TD Group");
                var timeDomainIndex = GetIndexFrom(igxlSheetsVersion, "Time Domain");
                var enableIndex = GetIndexFrom(igxlSheetsVersion, "Enable");
                var fileGroupNameIndex = GetIndexFrom(igxlSheetsVersion, "File/Group Name");
                var burstIndex = GetIndexFrom(igxlSheetsVersion, "Burst");
                var startLabelIndex = GetIndexFrom(igxlSheetsVersion, "Start Label");
                var stopLabelIndex = GetIndexFrom(igxlSheetsVersion, "Stop Label");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                var mainList = _patSets.Where(x => !x.PatSetRows.Any(y => y.IsBackup)).ToList();
                var backupList = _patSets.Where(x => x.PatSetRows.Any(y => y.IsBackup)).ToList();
                if (backupList.Any())
                {
                    var empty = new PatSet();
                    empty.PatSetRows.Add(new PatSetRow());
                    mainList.Add(empty);
                    mainList.AddRange(backupList);
                }

                for (var index = 0; index < mainList.Count; index++)
                {
                    var patSet = mainList[index];
                    for (int i = 0; i < patSet.PatSetRows.Count; i++)
                    {
                        var row = patSet.PatSetRows[i];
                        var arr = Enumerable.Repeat("", maxCount).ToArray();
                        arr[0] = row.ColumnA;
                        arr[patternSetIndex] = patSet.PatSetName;
                        if (tDGroupSetIndex != -1)
                            arr[tDGroupSetIndex] = row.TdGroup;
                        arr[timeDomainIndex] = row.TimeDomain;
                        arr[enableIndex] = row.Enable;
                        arr[fileGroupNameIndex] = row.File;
                        arr[burstIndex] = row.Burst;
                        arr[startLabelIndex] = row.StartLabel;
                        arr[stopLabelIndex] = row.StopLabel;
                        arr[commentIndex] = row.Comment;
                        sw.WriteLine(string.Join("\t", arr));
                    }
                }
                #endregion
            }
        }

        public void Append(string oldFile, string newFile, List<PatSet> patSets, string version = "2.2")
        {
            File.Copy(oldFile, newFile, true);
            using (var sw = File.AppendText(newFile))
            {
                foreach (var patSet in patSets)
                {
                    foreach (var patSetRow in patSet.PatSetRows)
                    {
                        var columnA = patSetRow.ColumnA ?? "";
                        sw.WriteLine(columnA + "\t" + patSet.PatSetName + "\t" + patSetRow.TdGroup + "\t" +
                            patSetRow.TimeDomain + "\t" + patSetRow.Enable + "\t" + patSetRow.File +
                            "\t" + patSetRow.Burst + "\t" + patSetRow.StartLabel + "\t" + patSetRow.StopLabel +
                            "\t" + patSetRow.Comment + "\t");
                    }
                }
            }
        }
        #endregion
    }
}