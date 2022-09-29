using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanPatSetWriter
    {
        protected string SheetName;

        public ScanPatSetWriter()
        {
            SheetName = "PatSet_Scan";
        }

        public PatSetSheet WritePatSet(IEnumerable<ProdCharRow> prodCharRows)
        {
            var patSetSheet = new PatSetSheet(SheetName);

            #region PatSet

            foreach (var prodCharRow in prodCharRows)
            {
                var exist = true;

                //multiple init pattern
                var pSet = new PatSet();
                pSet.PatSetName = prodCharRow.PatSetName;
                foreach (var init in prodCharRow.InitList.Values)
                {
                    if (init.PatternName.Equals(""))
                        continue;

                    var patSetRow = WriteItem(prodCharRow, init.PatternName);
                    string status;
                    if (!InputFiles.PatternListMap.GetStatusInPatternList(init.PatternName, out status))
                    {
                        exist = false;
                        patSetRow.AddComment(status);
                    }

                    pSet.AddRow(patSetRow);
                }

                //multiple payload pattern
                foreach (var payload in prodCharRow.PayloadList)
                {
                    if (payload.PatternName.Equals(""))
                        continue;
                    var patSetRow = WriteItem(prodCharRow, payload.PatternName);
                    string status;
                    if (!InputFiles.PatternListMap.GetStatusInPatternList(payload.PatternName, out status))
                    {
                        exist = false;
                        patSetRow.AddComment(status);
                    }

                    pSet.AddRow(patSetRow);
                }

                if (!exist || prodCharRow.Nop)
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

        private PatSetRow WriteItem(ProdCharRow prodCharRow, string patName)
        {
            var rowTemp = new PatSetRow();
            rowTemp.File = patName;
            rowTemp.Burst = "No";
            rowTemp.Comment = "Row : " + prodCharRow.RowNum;
            return rowTemp;
        }
    }
}