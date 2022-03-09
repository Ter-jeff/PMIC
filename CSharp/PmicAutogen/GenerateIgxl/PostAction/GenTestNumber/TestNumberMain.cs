using System;
using System.Collections.Generic;
using System.IO;
using PmicAutogen.GenerateIgxl.PostAction.GenTestNumber.Business;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.PostAction.GenTestNumber
{
    public class TestNumberMain
    {
        public List<string> WorkFlow()
        {
            var sheet = InputFiles.SettingWorkbook.Worksheets[PmicConst.BlockTestNumber];
            var nonTestNumberSheets = new List<string>();
            if (sheet != null)
            {
                var testNumberSheetReader = new TestNumberSheetReader(sheet);
                var subFlowSheets = TestProgram.IgxlWorkBk.SubFlowSheets;
                foreach (var sheetKey in subFlowSheets.Keys)
                {
                    var subFlowName = subFlowSheets[sheetKey].SheetName.ToUpper();
                    if (testNumberSheetReader.TestNumList.ContainsKey(subFlowName))
                    {
                        var startNum = testNumberSheetReader.TestNumList[subFlowName].StartNum;
                        var interval = testNumberSheetReader.TestNumList[subFlowName].Interval;
                        foreach (var row in subFlowSheets[sheetKey].FlowRows)
                        {
                            if (row.OpCode == null) continue;
                            if (row.OpCode.Equals("Test", StringComparison.OrdinalIgnoreCase) ||
                                row.OpCode.Equals("test-defer-limits", StringComparison.OrdinalIgnoreCase) ||
                                row.OpCode.Equals("call", StringComparison.OrdinalIgnoreCase))
                            {
                                row.Num = startNum.ToString();
                                if (startNum + interval <= testNumberSheetReader.TestNumList[subFlowName].MaxNum)
                                    startNum += interval;
                            }
                        }
                    }
                    else
                    {
                        var path = Path.GetDirectoryName(sheetKey);
                        if (path != FolderStructure.DirVbt)
                            nonTestNumberSheets.Add(subFlowName);
                    }
                }
            }

            return nonTestNumberSheets;
        }
    }
}