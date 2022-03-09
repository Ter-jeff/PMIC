using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace IgxlData.Others
{
    public class PatternUsedResult
    {
        public List<FlowRow> SubFlowRows;
        public List<PatternInTestProgram> PatternInTestPrograms;
    }

    public class PatternInTestProgram
    {
        public string ProjectName { get; set; }
        public string SheetName { get; set; }
        public string PatternSet { get; set; }
        public string PatternSet1 { get; set; }
        public string PatternSet2 { get; set; }
        public string TimeSets { get; set; }
        public string FilePath { get; set; }
        public string GenericPatternName { get; set; }
        public bool IsExist { get; set; }
        public bool IsUsed { get; set; }
        public string VbtFuncName { get; set; }
        public string TestName { get; set; }
        public InstanceRow InstanceRow { get; set; }
    }

    public class NewSecondOrder
    {
        public string PatsetName { get; set; }
        public string PatternSet1 { get; set; }
        public string PatternSet2 { get; set; }
        public string RealFile { get; set; }
        public string GenericPatternName { get; set; }
    }

    public class PatternUsed
    {
        private string _testProgramPath;
        private readonly List<string> _enables;
        private readonly string _job;
        private readonly string _env;
        private readonly string _mainflow;
        private bool OnlyPrintUsed { get; set; }

        List<string> FlowParameterList { get; set; }
        Dictionary<string, PatSetSheet> PatSetSheets { get; set; }
        List<NewSecondOrder> NewSecondOrderLists { get; set; }
        private Dictionary<string, SheetType> _dicAllTxt;
        private Regex _pattenRegex = new Regex(@"(\.\\(?:[^\\\?\/\*\|<>:]+\\)+)([^\\\?\/\*\|<>:]+?)\.(Pat|pat|PAT)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private List<PatternInTestProgram> _patternInTestPrograms;

        public PatternUsed(string testProgramPath, List<string> enables, string job, string env,  string mainflow,bool printUsed = false)
        {
            FlowParameterList = new List<string>();
            PatSetSheets = new Dictionary<string, PatSetSheet>();
            NewSecondOrderLists = new List<NewSecondOrder>();
            _patternInTestPrograms = new List<PatternInTestProgram>();
            _dicAllTxt = new Dictionary<string, SheetType>();
            _testProgramPath = testProgramPath;
            _enables = enables;
            _job = job;
            _env = env;
            _mainflow = mainflow;
            OnlyPrintUsed = printUsed;
        }

        public PatternUsedResult WorkFlow(string tempPath)
        {
            PatternUsedResult patternUsedResult = new PatternUsedResult();
            var reader = new IgxlSheetReader();
            _dicAllTxt = reader.GetSheetTypeDic(tempPath);

            var patSetSheets = (from pair in _dicAllTxt where pair.Value.ToString().Equals(SheetType.DTPatternSetSheet.ToString()) select pair.Key).ToList();
            foreach (var sheet in patSetSheets)
            {
                var readPatSetSheet = new ReadPatSetSheet();
                var patSetSheet = readPatSetSheet.GetSheet(sheet);
                PatSetSheets.Add(patSetSheet.Name, patSetSheet);
            }

            CheckNewSecondOrder();

            if (OnlyPrintUsed)
            {
                var subFlowSheets = GetSubFlowSheets();
                var subFlowRows = subFlowSheets.SelectMany(x => x.FlowRows).Where(x => x.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase)).ToList();
                if (_enables != null)
                    subFlowRows = subFlowRows.Where(x => x.IsMatchEnable(_enables)).ToList();
                if (!string.IsNullOrEmpty(_job))
                    subFlowRows = subFlowRows.Where(x => x.IsMatchJob(_job)).ToList();
                if (!string.IsNullOrEmpty(_env))
                    subFlowRows = subFlowRows.Where(x => x.IsMatchEnv(_env)).ToList();
                if (!string.IsNullOrEmpty(_mainflow))
                {
                    subFlowRows = subFlowRows.Where(x => x.IsMatchEnv(_env)).ToList();
                    if (subFlowSheets.Exists(x=>x.Name.Equals(_mainflow,StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var mainFlowRows = subFlowSheets.Find(x=>x.Name.Equals(_mainflow,StringComparison.CurrentCultureIgnoreCase)).FlowRows
                            .Where(x => x.Opcode.Equals("Call", StringComparison.CurrentCultureIgnoreCase) || x.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase))
                            .ToList();
                        if (_enables != null)
                            mainFlowRows = mainFlowRows.Where(x => x.IsMatchEnable(_enables)).ToList();
                        if (!string.IsNullOrEmpty(_job))
                            mainFlowRows = mainFlowRows.Where(x => x.IsMatchJob(_job)).ToList();
                        if (!string.IsNullOrEmpty(_env))
                            mainFlowRows = mainFlowRows.Where(x => x.IsMatchEnv(_env)).ToList();
                         subFlowRows = GetFlowRowsByMainFlow(mainFlowRows, subFlowRows);
                    }
                }
                patternUsedResult.SubFlowRows = subFlowRows;
                FlowParameterList = subFlowRows.Select(x => x.Parameter.Split(' ').First()).Distinct().ToList();
                GetInstancePatForArgList();
            }
            else
            {
                PrintAllPatset();
            }
            patternUsedResult.PatternInTestPrograms = _patternInTestPrograms;
            return patternUsedResult;
        }

        private static List<FlowRow> GetFlowRowsByMainFlow(List<FlowRow> mainFlowRows, List<FlowRow> subFlowRows)
        {
            var newsubFlowRows = new List<FlowRow>();
            foreach (var mainFlowRow in mainFlowRows)
            {
                var rows =subFlowRows.Where(x => x.SheetName.Equals(mainFlowRow.Parameter, StringComparison.CurrentCultureIgnoreCase)).ToList();
                newsubFlowRows.AddRange(rows);
            }
            return newsubFlowRows;
        }

        private List<SubFlowSheet> GetSubFlowSheets()
        {
            List<SubFlowSheet> subFlowSheets = new List<SubFlowSheet>();
            var subFlow = (from pair in _dicAllTxt where pair.Value.ToString().Equals(SheetType.DTFlowtableSheet.ToString()) select pair.Key).ToList();
            foreach (var flow in subFlow)
            {
                var readFlowSheet = new ReadFlowSheet();
                subFlowSheets.Add(readFlowSheet.GetSheet(flow, true));
            }
            return subFlowSheets;
        }

        private void CheckNewSecondOrder()
        {
            var patternSetList = PatSetSheets.SelectMany(x => x.Value.PatSetRows).ToList();
            foreach (var item in patternSetList)
            {
                if (string.IsNullOrEmpty(item.PatSetName)) continue;

                foreach (var patSetRow in item.PatSetRows)
                {
                    NewSecondOrder secondOrder = new NewSecondOrder { PatsetName = item.PatSetName };
                    GetReadPatFileName(item.PatSetName, patSetRow, secondOrder, patternSetList, 0, NewSecondOrderLists);
                }
            }
        }

        private void GetReadPatFileName(string mainPatset, PatSetRow patSetRow, NewSecondOrder secondOrder, List<PatSet> patternSetList, int level, List<NewSecondOrder> newSecondOrderList)
        {
            try
            {
                if (!string.IsNullOrEmpty(secondOrder.RealFile)) return;
                var temp = patternSetList.FindAll(x => string.Equals(x.PatSetName, patSetRow.File, StringComparison.CurrentCultureIgnoreCase));
                if (temp.Any())
                {
                    foreach (var fileRow in temp[0].PatSetRows)
                    {
                        if (!string.IsNullOrEmpty(secondOrder.RealFile))
                        {
                            //newSecondOrderList.Add(secondOrder);
                            secondOrder = new NewSecondOrder { PatsetName = mainPatset };
                            level = 0;
                        }

                        if (!_pattenRegex.IsMatch(fileRow.File) && fileRow.File != string.Empty)
                        {
                            ++level;
                            if (level == 1)
                                secondOrder.PatternSet1 = patSetRow.File;
                            else if (level == 2)
                                secondOrder.PatternSet2 = patSetRow.File;
                            GetReadPatFileName(mainPatset, fileRow, secondOrder, patternSetList, level, newSecondOrderList);
                        }
                        else
                        {
                            if (level == 0)
                                secondOrder.PatternSet1 = patSetRow.File;
                            if (level == 1)
                                secondOrder.PatternSet2 = patSetRow.File;

                            secondOrder.RealFile = fileRow.File;
                            secondOrder.GenericPatternName = fileRow.PatternSet;

                            if (level == 2)
                            {
                                if (!_pattenRegex.IsMatch(secondOrder.RealFile) && !string.IsNullOrEmpty(secondOrder.RealFile))
                                    secondOrder.RealFile = "over 3 order";
                                newSecondOrderList.Add(secondOrder);
                            }

                            if (_pattenRegex.IsMatch(secondOrder.RealFile) && !string.IsNullOrEmpty(secondOrder.RealFile))
                            {
                                newSecondOrderList.Add(secondOrder);
                            }
                        }
                    }
                }
                else
                {
                    if (_pattenRegex.IsMatch(patSetRow.File) && !string.IsNullOrEmpty(patSetRow.File))
                    {
                        secondOrder.RealFile = patSetRow.File;
                        secondOrder.GenericPatternName = patSetRow.PatternSet;
                        newSecondOrderList.Add(secondOrder);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(ex.ToString()));
                throw;
            }
        }

        private void GetInstancePatForArgList()
        {
            var instanceFlow = (from pair in _dicAllTxt where pair.Value.ToString().Equals(SheetType.DTTestInstancesSheet.ToString()) select pair.Key).ToList();
            var patternSetList = PatSetSheets.SelectMany(x => x.Value.PatSetRows).ToList();
            foreach (var sheet in instanceFlow)
            {
                var readInstanceSheet = new ReadInstanceSheet();
                var instanceSheet = readInstanceSheet.GetSheet(sheet);
                foreach (var instRow in instanceSheet.InstanceRows)
                {
                    if (instRow.TestName == "LAPLL_T21S_MULTIPLE_PP_BTCA0_S_PL00_AN_LAPL_DLL_JTG_VIX_ALLFRV_SI_LAPLL_T21S_NV")
                    { 
                    }
                    //be used Instance Add to instForVbtCheck
                    if (FlowParameterList.Exists(p => p.Equals(instRow.TestName, StringComparison.OrdinalIgnoreCase)))
                    {
                        // if Instance FlowSheet arglist exist in PatternCopyList , add to  instForArgList,  XXX.pat also add to  PatternCopyList
                        for (int y = 0; y < instRow.Args.Count; y++)
                        {
                            if (string.IsNullOrEmpty(instRow.Args[y]))
                                continue;
                            var arr = Regex.Split(instRow.Args[y],";");// @"[^\w]");
                            foreach (var item in arr)
                            {
                                if (patternSetList.Exists(x => x.PatSetName.Equals(item, StringComparison.CurrentCultureIgnoreCase)) || _pattenRegex.IsMatch(item))
                                    RefreshPatternSecOrderFileName(instanceSheet.Name, instRow, item);
                            }
                        }
                    }
                }
            }

        }

        private void PrintAllPatset()
        {
            foreach (var item in NewSecondOrderLists)
            {
                var pat = item.RealFile;
                if (item.RealFile.Contains(":"))
                    pat = item.RealFile.Substring(0, item.RealFile.LastIndexOf(":", StringComparison.Ordinal));
                _patternInTestPrograms.Add(new PatternInTestProgram
                {
                    PatternSet = item.PatsetName,
                    PatternSet1 = item.PatternSet1,
                    PatternSet2 = item.PatternSet2,
                    FilePath = pat,
                    GenericPatternName = item.GenericPatternName
                });
            }
        }

        private void RefreshPatternSecOrderFileName(string sheet, InstanceRow instanceRow, string patName)
        {
            var testName = instanceRow.TestName;
            var timeSet = instanceRow.TimeSets;
            if (_pattenRegex.IsMatch(patName))
            {
                var pat = patName;
                if (patName.Contains(":"))
                    pat = patName.Substring(0, patName.LastIndexOf(":", StringComparison.Ordinal));
                _patternInTestPrograms.Add(new PatternInTestProgram
                {
                    SheetName = sheet,
                    TestName = testName,
                    PatternSet = pat,
                    TimeSets = timeSet,
                    FilePath = pat,
                    GenericPatternName = Path.GetFileNameWithoutExtension(pat),
                    InstanceRow = instanceRow
                });
            }
            else
            {
                var target = NewSecondOrderLists.FindAll(x => x.PatsetName.Equals(patName, StringComparison.CurrentCultureIgnoreCase));
                foreach (var item in target)
                {
                    var pat = item.RealFile;
                    if (item.RealFile.Contains(":"))
                        pat = item.RealFile.Substring(0, item.RealFile.LastIndexOf(":", StringComparison.Ordinal));
                    _patternInTestPrograms.Add(new PatternInTestProgram
                    {
                        SheetName = sheet,
                        TestName = testName,
                        PatternSet = patName,
                        PatternSet1 = item.PatternSet1,
                        PatternSet2 = item.PatternSet2,
                        TimeSets = timeSet,
                        FilePath = pat,
                        GenericPatternName = item.GenericPatternName,
                        InstanceRow = instanceRow
                    });
                }
            }
        }
    }
}
