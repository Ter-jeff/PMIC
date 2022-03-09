using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    public class InstanceSheet : IgxlSheet
    {
        private const string SheetType = "DTTestInstancesSheet";

        #region Field
        private List<InstanceRow> _instanceRows;
        #endregion

        #region Property

        public List<InstanceRow> InstanceRows
        {
            get { return _instanceRows; }
            set { _instanceRows = value; }
        }

        #endregion

        #region Constructor

        public InstanceSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _instanceRows = new List<InstanceRow>();
            IgxlSheetName = IgxlSheetNameList.TestInstance;
        }

        public InstanceSheet(string sheetName)
            : base(sheetName)
        {
            _instanceRows = new List<InstanceRow>();
            IgxlSheetName = IgxlSheetNameList.TestInstance;
        }
        #endregion

        #region Member Function

        public void AddRow(InstanceRow igxlItem)
        {
            _instanceRows.Add(igxlItem);
        }

        public void AddRows(List<InstanceRow> igxlItems)
        {
            _instanceRows.AddRange(igxlItems);
        }

        public void WriteNew(string fileName, string version = "2.4")
        {
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                return;
            using (var sw = new StreamWriter(fileName, false))
            {
                sw.WriteLine("DTTestInstancesSheet,version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTest Instances\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
                sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
                sw.WriteLine("\t\tTest Procedure\t\t\tDC Specs\t\tAC Specs\t\tSheet Parameters\t\t\t\t\tOther Parameters\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
                sw.WriteLine("\tTest Name\tType\tName\tCalled As\tCategory\tSelector\tCategory\tSelector\tTime Sets\tEdge Sets\tPin Levels\tMixed Signal Timing\tOverlay\tArg0\tArg1\tArg2\tArg3\tArg4\tArg5\tArg6\tArg7\tArg8\tArg9\tArg10\tArg11\tArg12\tArg13\tArg14\tArg15\tArg16\tArg17\tArg18\tArg19\tArg20\tArg21\tArg22\tArg23\tArg24\tArg25\tArg26\tArg27\tArg28\tArg29\tArg30\tArg31\tArg32\tArg33\tArg34\tArg35\tArg36\tArg37\tArg38\tArg39\tArg40\tArg41\tArg42\tArg43\tArg44\tArg45\tArg46\tArg47\tArg48\tArg49\tArg50\tArg51\tArg52\tArg53\tArg54\tArg55\tArg56\tArg57\tArg58\tArg59\tArg60\tArg61\tArg62\tArg63\tArg64\tArg65\tArg66\tArg67\tArg68\tArg69\tArg70\tArg71\tArg72\tArg73\tArg74\tArg75\tArg76\tArg77\tArg78\tArg79\tArg80\tArg81\tArg82\tArg83\tArg84\tArg85\tArg86\tArg87\tArg88\tArg89\tArg90\tArg91\tArg92\tArg93\tArg94\tArg95\tArg96\tArg97\tArg98\tArg99\tArg100\tArg101\tArg102\tArg103\tArg104\tArg105\tArg106\tArg107\tArg108\tArg109\tArg110\tArg111\tArg112\tArg113\tArg114\tArg115\tArg116\tArg117\tArg118\tArg119\tArg120\tArg121\tArg122\tArg123\tArg124\tArg125\tArg126\tArg127\tArg128\tArg129\tComment\t");
                //sw.WriteLine("\tTest Name\tType\tName\tCalled As\tCategory\tSelector\tCategory\tSelector\tTime Sets\tEdge Sets\tPin Levels\tMixed Signal Timing\tArgList\tArg1\tArg2\tArg3\tArg4\tArg5\tArg6\tArg7\tArg8\tArg9\tArg10\tArg11\tArg12\tArg13\tArg14\tArg15\tArg16\tArg17\tArg18\tArg19\tArg20\tArg21\tArg22\tArg23\tArg24\tArg25\tArg26\tArg27\tArg28\tArg29\tArg30\tArg31\tArg32\tArg33\tArg34\tArg35\tArg36\tArg37\tArg38\tArg39\tArg40\tArg41\tArg42\tArg43\tArg44\tArg45\tArg46\tArg47\tArg48\tArg49\tArg50\tArg51\tArg52\tArg53\tArg54\tArg55\tArg56\tArg57\tArg58\tArg59\tArg60\tArg61\tArg62\tArg63\tArg64\tArg65\tArg66\tArg67\tArg68\tArg69\tArg70\tArg71\tArg72\tArg73\tArg74\tArg75\tArg76\tArg77\tArg78\tArg79\tArg80\tArg81\tArg82\tArg83\tArg84\tArg85\tArg86\tArg87\tArg88\tArg89\tArg90\tArg91\tArg92\tArg93\tArg94\tArg95\tArg96\tArg97\tArg98\tArg99\tArg100\tArg101\tArg102\tArg103\tArg104\tArg105\tArg106\tArg107\tArg108\tArg109\tArg110\tArg111\tArg112\tArg113\tArg114\tArg115\tArg116\tArg117\tArg118\tArg119\tArg120\tArg121\tArg122\tArg123\tArg124\tArg125\tArg126\tArg127\tArg128\tArg129\tArg130\tComment\t");

                var instanceRows = _instanceRows.Where(x => x.IsBackup == false).ToList();

                foreach (var instanceRow in instanceRows)
                {
                    var argrument = "";
                    foreach (var arg in instanceRow.Args)
                        argrument += "\t" + arg;

                    sw.WriteLine(instanceRow.ColumnA + "\t" + instanceRow.TestName + "\t" + instanceRow.Type + "\t" + instanceRow.Name + "\t" +
                         instanceRow.CalledAs + "\t" + instanceRow.DcCategory + "\t" + instanceRow.DcSelector + "\t" +
                         instanceRow.AcCategory + "\t" + instanceRow.AcSelector + "\t" + instanceRow.TimeSets + "\t" +
                         instanceRow.EdgeSets + "\t" + instanceRow.PinLevels + "\t" + instanceRow.MixedSignalTiming + "\t" +
                         instanceRow.Overlay + "\t" + instanceRow.ArgList + argrument);
                }


                var backup = _instanceRows.Where(x => x.IsBackup).ToList();
                if (backup.Any())
                {
                    sw.WriteLine("");
                    foreach (var instanceRow in backup)
                    {
                        var argrument = "";
                        foreach (var arg in instanceRow.Args)
                            argrument += "\t" + arg;

                        sw.WriteLine(instanceRow.ColumnA + "\t" + instanceRow.TestName + "\t" + instanceRow.Type + "\t" +
                                     instanceRow.Name + "\t" +
                                     instanceRow.CalledAs + "\t" + instanceRow.DcCategory + "\t" +
                                     instanceRow.DcSelector + "\t" +
                                     instanceRow.AcCategory + "\t" + instanceRow.AcSelector + "\t" +
                                     instanceRow.TimeSets + "\t" +
                                     instanceRow.EdgeSets + "\t" + instanceRow.PinLevels + "\t" +
                                     instanceRow.MixedSignalTiming + "\t" +
                                     instanceRow.Overlay + "\t" + instanceRow.ArgList + argrument);
                    }
                }
            }
        }

        public override void Write(string fileName, string version = "2.4")
        {
            //if (version == "2.4")
            //{
            //    var validate = new Action<string>((a) => { });
            //    var instanceGen = new GenTestInstanceSheet(fileName, validate, true);

            //    var instanceRows = _instanceRows.Where(x => x.IsBackup == false).ToList();
            //    foreach (var instanceRow in instanceRows)
            //    {
            //        for (var i = 0; i < instanceRow.Args.Count; i++)
            //        {
            //            if (instanceRow.Args[i].Contains('\n'))
            //                instanceRow.Args[i] = instanceRow.Args[i].Replace("\n", "");
            //        }

            //        instanceGen.AddTestInstance(instanceRow.TestName, instanceRow.Type, instanceRow.Name,
            //            instanceRow.CalledAs,
            //            instanceRow.ArgList, instanceRow.Args.ToArray(), instanceRow.DcCategory, instanceRow.DcSelector,
            //            instanceRow.PinLevels, "", instanceRow.TimeSets, instanceRow.AcCategory, instanceRow.AcSelector,
            //            instanceRow.EdgeSets, instanceRow.MixedSignalTiming, instanceRow.Overlay);
            //    }

            //    var backup = _instanceRows.Where(x => x.IsBackup).ToList();
            //    if (backup.Any())
            //    {
            //        var empty = new InstanceRow();
            //        instanceGen.AddTestInstance(empty.TestName, empty.Type, empty.Name,
            //            empty.CalledAs,
            //            empty.ArgList, empty.Args.ToArray(), empty.DcCategory, empty.DcSelector,
            //            empty.PinLevels, "", empty.TimeSets, empty.AcCategory, empty.AcSelector,
            //            empty.EdgeSets, empty.MixedSignalTiming, empty.Overlay);


            //        foreach (var instanceRow in backup)
            //        {
            //            for (var i = 0; i < instanceRow.Args.Count; i++)
            //            {
            //                if (instanceRow.Args[i].Contains('\n'))
            //                    instanceRow.Args[i] = instanceRow.Args[i].Replace("\n", "");
            //            }

            //            instanceGen.AddTestInstance(instanceRow.TestName, instanceRow.Type, instanceRow.Name,
            //                instanceRow.CalledAs,
            //                instanceRow.ArgList, instanceRow.Args.ToArray(), instanceRow.DcCategory,
            //                instanceRow.DcSelector,
            //                instanceRow.PinLevels, "", instanceRow.TimeSets, instanceRow.AcCategory,
            //                instanceRow.AcSelector,
            //                instanceRow.EdgeSets, instanceRow.MixedSignalTiming, instanceRow.Overlay);
            //        }
            //    }

            //    instanceGen.WriteSheet();
            //}
            //else
            //    throw new Exception(string.Format("The TestInstance sheet version:{0} is not supported!", version));

            //Support 2.4
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
                else if (dic.ContainsKey("2.4"))
                {
                    var igxlSheetsVersion = dic["2.4"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }


        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_instanceRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var testNameIndex = GetIndexFrom(igxlSheetsVersion, "Test Name");
                var typeIndex = GetIndexFrom(igxlSheetsVersion, "Test Procedure", "Type");
                var nameIndex = GetIndexFrom(igxlSheetsVersion, "Test Procedure", "Name");
                var calledAsIndex = GetIndexFrom(igxlSheetsVersion, "Test Procedure", "Called As");
                var dcSpecsCategoryIndex = GetIndexFrom(igxlSheetsVersion, "DC Specs", "Category");
                var dcSpecsSelectorIndex = GetIndexFrom(igxlSheetsVersion, "DC Specs", "Selector");
                var acSpecsCategoryIndex = GetIndexFrom(igxlSheetsVersion, "AC Specs", "Category");
                var acSpecsSelectorIndex = GetIndexFrom(igxlSheetsVersion, "AC Specs", "Selector");
                var timeSetsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Time Sets");
                var edgeSetsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Edge Sets");
                var pinLevelsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Pin Levels");
                var mixedSignalTimingIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Mixed Signal Timing");
                var overlayIndex = GetIndexFrom(igxlSheetsVersion, "Overlay");
                var arg0Index = GetIndexFrom(igxlSheetsVersion, "Other Parameters", "Arg0");
                var arg1Index = GetIndexFrom(igxlSheetsVersion, "Other Parameters", "Arg1");
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
                for (var index = 0; index < _instanceRows.Count; index++)
                {
                    var row = _instanceRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.TestName))
                    {
                        arr[0] = row.ColumnA;
                        arr[testNameIndex] = row.TestName;
                        arr[typeIndex] = row.Type;
                        arr[nameIndex] = row.Name;
                        arr[calledAsIndex] = row.CalledAs;
                        arr[dcSpecsCategoryIndex] = row.DcCategory;
                        arr[dcSpecsSelectorIndex] = row.DcSelector;
                        arr[acSpecsCategoryIndex] = row.AcCategory;
                        arr[acSpecsSelectorIndex] = row.AcSelector;
                        arr[timeSetsIndex] = row.TimeSets;
                        arr[edgeSetsIndex] = row.EdgeSets;
                        arr[pinLevelsIndex] = row.PinLevels;
                        arr[mixedSignalTimingIndex] = row.MixedSignalTiming;
                        arr[overlayIndex] = row.Overlay;
                        arr[arg0Index] = row.ArgList;
                        for (int i = 0; i < row.Args.Count; i++)
                        {
                            var arg = row.Args[i];
                            arr[arg1Index + i] = arg;
                        }
                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] { "\t" };
                    }
                    sw.WriteLine(string.Join("\t", arr));
                }
                #endregion
            }
        }

        public void Append(string oldFile, string newFile, List<InstanceRow> instanceRows, string version = "2.4")
        {
            File.Copy(oldFile, newFile, true);
            using (var sw = File.AppendText(newFile))
            {
                foreach (var instanceRow in instanceRows)
                {
                    var argrument = "";
                    foreach (var arg in instanceRow.Args)
                    {
                        argrument += "\t" + arg;
                    }
                    sw.WriteLine("TTR\t" + instanceRow.TestName + "\t" + instanceRow.Type + "\t" + instanceRow.Name + "\t" +
                         instanceRow.CalledAs + "\t" + instanceRow.DcCategory + "\t" + instanceRow.DcSelector + "\t" +
                         instanceRow.AcCategory + "\t" + instanceRow.AcSelector + "\t" + instanceRow.TimeSets + "\t" +
                         instanceRow.EdgeSets + "\t" + instanceRow.PinLevels + "\t" + instanceRow.MixedSignalTiming + "\t" +
                         instanceRow.Overlay + "\t" + instanceRow.ArgList + argrument);
                }
            }
        }

        public List<string> GetTestNameList(List<string> nopPatSets)
        {
            var testNames = new List<string>();
            foreach (var pattern in nopPatSets)
            {
                if (_instanceRows.Exists(x => x.Args.First().Equals(pattern, StringComparison.OrdinalIgnoreCase)))
                {
                    var rows = _instanceRows.Where(x => x.Args.First().Equals(pattern, StringComparison.OrdinalIgnoreCase)).ToList();
                    testNames.AddRange(rows.Select(x => x.TestName.ToUpper()));
                }
            }
            return testNames;
        }
        #endregion

        public void AddHeaderFooter(string sheetName)
        {
            var block = Regex.Replace(sheetName, "^Flow_", "", RegexOptions.IgnoreCase);
            var header = new InstanceRow();
            header.TestName = block + "_Header";
            header.Type = "VBT";
            header.ArgList = "PrintInfo";
            header.Name = "Print_Header";
            header.Args.Add(block);

            var footer = new InstanceRow();
            footer.TestName = block + "_Footer";
            footer.Type = "VBT";
            footer.ArgList = "PrintInfo";
            footer.Name = "Print_Footer";
            footer.Args.Add(block);
            AddRow(footer);
        }

        public InstanceSheet DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as InstanceSheet;
            }
        }
    }
}