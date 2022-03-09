using System.Linq;
using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Teradyne.Oasis.IGLinkBase.ProgramGeneration;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class PatSetSheet : IgxlSheet
    {
        #region Field
        private const string SheetType = "DTPatternSetSheet";
        private List<PatSet> _patSetData;
        private readonly Dictionary<string, int> _patSetNameDictionary;
        #endregion

        #region Property
        public List<PatSet> PatSetRows
        {
            set { _patSetData = value; }
            get
            {
                return _patSetData ?? (_patSetData = new List<PatSet>());
            }
        }
        #endregion

        #region Constructor
        public PatSetSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _patSetData = new List<PatSet>();
            IgxlSheetName = IgxlSheetNameList.PatternSet;
            _patSetNameDictionary = new Dictionary<string, int>();

        }

        public PatSetSheet(string sheetName)
            : base(sheetName)
        {
            _patSetData = new List<PatSet>();
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
            if (!_patSetData.Exists(x => x.PatSetName.Equals(igxlItem.PatSetName, StringComparison.OrdinalIgnoreCase)))
                return false;

            var row = _patSetData.Find(x => x.PatSetName.Equals(igxlItem.PatSetName, StringComparison.OrdinalIgnoreCase));
            if (row.PatSetRows.Count != igxlItem.PatSetRows.Count)
                return false;
            for (int i = 0; i < igxlItem.PatSetRows.Count; i++)
            {
                if (!igxlItem.PatSetRows[i].File.Equals(row.PatSetRows[i].File, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        public void AddPatSet(PatSet igxlItem)
        {
            _patSetData.Add(igxlItem);
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
                _patSetData.Add(igxlItem);
                if (igxlItem.PatSetName == null) continue;
                if (!_patSetNameDictionary.ContainsKey(igxlItem.PatSetName))
                    _patSetNameDictionary.Add(igxlItem.PatSetName, 0);
            }
        }

        public long GetPatSetCnt()
        {
            return _patSetData.Count;
        }

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
        }

        public override void Write(string fileName, string version = "2.2")
        {
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version=="2.2")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (version=="2.3")
                {
                    var igxlSheetsVersion = dic["2.3"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The Pattern Set sheet version:{0} is not supported!", version));
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
                var mainList = _patSetData.Where(x => !x.IsBackup).ToList();
                var backupList = _patSetData.Where(x => x.IsBackup).ToList();
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

        private void WritePatternSet2_3(string filePath)
        {
            using (var sw = new StreamWriter(filePath, false))
            {
                sw.WriteLine("DTPatternSetSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPattern Sets");
                sw.WriteLine("\t\t\t\t\t\t\t\t\t");
                sw.WriteLine("\tPattern Set\tTime Domain\tEnable\tFile/Group Name\tBurst\tStart Label\tStop Label\tComment\t");

                var backupList = _patSetData.Where(x => x.IsBackup).ToList();
                var mainList = _patSetData.Where(x => !x.IsBackup).ToList();

                foreach (PatSet patSet in mainList)
                {
                    foreach (var patSetRow in patSet.PatSetRows)
                    {
                        sw.WriteLine("{0}{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}","\t",patSet.PatSetName, patSetRow.TimeDomain, patSetRow.Enable, patSetRow.File, patSetRow.Burst, patSetRow.StartLabel, patSetRow.StopLabel, patSetRow.Comment);
                    }
                }
                if (backupList.Count != 0)
                    sw.WriteLine("\t\t\t\t\t\t\t\t\t");
                foreach (PatSet patSet in backupList)
                {
                    foreach (var patSetRow in patSet.PatSetRows)
                    {
                        sw.WriteLine("{0}{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}", "\t",patSet.PatSetName, patSetRow.TimeDomain, patSetRow.Enable, patSetRow.File, patSetRow.Burst, patSetRow.StartLabel, patSetRow.StopLabel, patSetRow.Comment);
                    }
                }
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
                        string columnA = patSetRow.ColumnA ?? "";
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
