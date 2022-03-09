using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.PostAction.GenTestNumber.Base;

namespace PmicAutogen.GenerateIgxl.PostAction.GenTestNumber.Business
{
    public class TestNumberSheetReader
    {
        public Dictionary<string, TestNumberBase> TestNumList = new Dictionary<string, TestNumberBase>();

        public TestNumberSheetReader(ExcelWorksheet sheet)
        {
            #region Information

            var flowKeyWdSearch = CellStr("Sub Flow Name");
            var numStartWdSearch = CellStr("Testnumber (start)");
            var numEndWdSearch = CellStr("Testnumber (end)");
            var numItvWdSearch = CellStr("Step Num");
            var stRow = -1;
            var spRow = -1;
            var flowCol = -1;
            var numStCol = -1;
            var numSpCol = -1;
            var stepCol = -1;
            var searchColMax = 99;
            if (sheet.Dimension.End.Column < searchColMax) searchColMax = sheet.Dimension.End.Column;

            #endregion

            #region To obtain the column locations

            for (var iRow = 1; iRow <= sheet.Dimension.End.Row; iRow++)
            for (var iCol = 1; iCol <= searchColMax; iCol++)
            {
                if (Regex.IsMatch(CellStr(Convert.ToString(sheet.Cells[iRow, iCol].Value)), flowKeyWdSearch,
                    RegexOptions.IgnoreCase))
                {
                    stRow = iRow;
                    flowCol = iCol;
                }

                if (Regex.IsMatch(CellStr(Convert.ToString(sheet.Cells[iRow, iCol].Value)), numStartWdSearch,
                    RegexOptions.IgnoreCase)) numStCol = iCol;
                if (Regex.IsMatch(CellStr(Convert.ToString(sheet.Cells[iRow, iCol].Value)), numEndWdSearch,
                    RegexOptions.IgnoreCase)) numSpCol = iCol;
                if (Regex.IsMatch(CellStr(Convert.ToString(sheet.Cells[iRow, iCol].Value)), numItvWdSearch,
                    RegexOptions.IgnoreCase)) stepCol = iCol;
                if (flowCol > 0 && numStCol > 0 && stepCol > 0) break;
            }

            #endregion

            #region To obtain the Maximum Row

            for (var iRow = stRow; iRow <= sheet.Dimension.End.Row; iRow++)
            {
                spRow = iRow;
                if (CellStr(Convert.ToString(sheet.Cells[iRow, flowCol].Value)) == "")
                    break;
            }

            #endregion

            #region To obtain the values

            var lastSubFlow = "";
            for (var iRow = stRow + 1; iRow <= spRow; iRow++)
            {
                var newItem = new TestNumberBase();
                long tmp;
                var subFlow = CellStr(Convert.ToString(sheet.Cells[iRow, flowCol].Value)).ToUpper();
                if (subFlow == "") continue;
                if (CellStr(Convert.ToString(sheet.Cells[iRow, numStCol].Value)) != "")
                {
                    long.TryParse(CellStr(Convert.ToString(sheet.Cells[iRow, numStCol].Value)), out tmp);
                    newItem.StartNum = tmp;
                    if (lastSubFlow.Length > 0)
                        if (TestNumList[lastSubFlow].MaxNum == 999999999)
                            TestNumList[lastSubFlow].MaxNum = newItem.StartNum - 1;
                }

                if (CellStr(Convert.ToString(sheet.Cells[iRow, numSpCol].Value)) != "")
                {
                    long.TryParse(CellStr(Convert.ToString(sheet.Cells[iRow, numSpCol].Value)), out tmp);
                    newItem.MaxNum = tmp;
                }

                if (CellStr(Convert.ToString(sheet.Cells[iRow, stepCol].Value)) != "")
                {
                    long.TryParse(CellStr(Convert.ToString(sheet.Cells[iRow, stepCol].Value)), out tmp);
                    newItem.Interval = tmp;
                }

                TestNumList.Add(subFlow, newItem);
                lastSubFlow = subFlow;
            }

            #endregion

            #region Check the max test number against next subFlow

            var testNumSpErr = "";
            var checkCnt = 0;
            long maxLast = 0;
            foreach (var subFlow in TestNumList.Keys)
            {
                if (checkCnt == 0)
                {
                    maxLast = TestNumList[subFlow].MaxNum;
                }
                else
                {
                    if (TestNumList[subFlow].StartNum <= maxLast)
                        testNumSpErr += "Subflow=" + subFlow + ", the Start Test# less than last subFlow End! Start#=" +
                                        TestNumList[subFlow].StartNum + ";  last End#=" + maxLast + "\n";
                    maxLast = TestNumList[subFlow].MaxNum;
                }

                checkCnt++;
            }

            #endregion

            #region Error exporting

            if (testNumSpErr.Length > 0)
            {
                const string message = "TestNumber Assignment Contained Error";
                testNumSpErr = sheet.Name + "\n" + testNumSpErr;
                MessageBox.Show(testNumSpErr, message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            #endregion
        }

        private string CellStr(string inStr)
        {
            inStr = Convert.ToString(inStr).Replace(" ", "");
            inStr = Convert.ToString(inStr).Replace("\n", "");
            inStr = Convert.ToString(inStr).Replace("(", "");
            inStr = Convert.ToString(inStr).Replace(")", "");
            return inStr;
        }
    }
}