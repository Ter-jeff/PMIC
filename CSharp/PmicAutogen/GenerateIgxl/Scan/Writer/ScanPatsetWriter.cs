using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanPatSetWriter
    {
        protected string SheetName;

        public ScanPatSetWriter()
        {
            SheetName = "PatSet_Scan";
        }

        #region Member function

        public PatSetSheet WritePatSet(List<ProdCharRowScan> prodCharRowScans)
        {
            var patSetSheet = new PatSetSheet(SheetName);

            #region PatSet

            foreach (var testInstance in prodCharRowScans)
            {
                var exist = true;

                //multiple init pattern
                var pSet = new PatSet();
                pSet.PatSetName = testInstance.PatSetName;
                foreach (var init in testInstance.InitList.Values)
                {
                    if (init.PatternName.Equals(""))
                        continue;

                    var patSetRow = WriteItem(testInstance, init.PatternName);
                    string status;
                    if (!InputFiles.PatternListMap.GetStatusInPatternList(init.PatternName, out status))
                    {
                        exist = false;
                        patSetRow.AddComment(status);
                    }

                    pSet.AddRow(patSetRow);
                }

                //multiple payload pattern
                foreach (var payload in testInstance.PayloadList)
                {
                    if (payload.PatternName.Equals(""))
                        continue;
                    var patSetRow = WriteItem(testInstance, payload.PatternName);
                    string status;
                    if (!InputFiles.PatternListMap.GetStatusInPatternList(payload.PatternName, out status))
                    {
                        exist = false;
                        patSetRow.AddComment(status);
                    }

                    pSet.AddRow(patSetRow);
                }

                if (!exist || testInstance.Nop)
                    pSet.IsBackup = true;
                patSetSheet.AddPatSet(pSet);
            }

            #endregion

            var row = new PatSet();
            var setRow = new PatSetRow();
            row.AddRow(setRow);
            patSetSheet.AddPatSet(row);

            return patSetSheet;
        }

        private PatSetRow WriteItem(ProdCharRowScan prodCharRowScan, string patName)
        {
            var rowTemp = new PatSetRow();
            rowTemp.File = patName;
            rowTemp.Burst = "No";
            rowTemp.Comment = "Row : " + prodCharRowScan.RowNum;
            return rowTemp;
        }

        #endregion
    }
}