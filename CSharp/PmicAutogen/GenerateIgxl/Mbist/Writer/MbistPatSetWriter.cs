using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistPatSetWriter : ScanPatSetWriter
    {
        public MbistPatSetWriter()
        {
            SheetName = "PatSet_Mbist";
        }

        #region Member function

        public PatSetSheet WritePatSet(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var patSetSheet = new PatSetSheet(SheetName);

            #region PatSet

            foreach (var testInstance in prodCharRowMbists)
            {
                if (testInstance.Nop)
                    continue;

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

        private PatSetRow WriteItem(ProdCharRowMbist instanceName, string patName)
        {
            var rowTemp = new PatSetRow();
            rowTemp.File = patName;
            rowTemp.Burst = "No";
            rowTemp.Comment = "Row : " + instanceName.RowNum;
            return rowTemp;
        }

        #endregion
    }
}