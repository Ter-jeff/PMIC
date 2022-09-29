using OfficeOpenXml;
using PmicAutogen.InputPackages.Base;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PmicAutogen.InputPackages.Inputs
{
    public class InputTestPlan : ExcelInput
    {
        public InputTestPlan(FileInfo fileInfo) : base(fileInfo, InputFileType.TestPlan)
        {
            var regEFuseSheets = @"^EFUSE_";
            var efuseSheetList = SheetList.FindAll(p => Regex.IsMatch(p, regEFuseSheets, RegexOptions.IgnoreCase));
            EfuseFormat = GetEFuseVersion(Workbook, efuseSheetList);
        }

        public int EfuseFormat { get; set; }

        private int GetEFuseVersion(ExcelWorkbook workbook, List<string> efuseSheetNames)
        {
            bool[] flag = { false, false, false, false, false, false, false, false, false };
            int revision;
            var revisionCheck = 0;
            var revisionCount = 0;
            var accessBitCount = 0;
            foreach (var sheet in efuseSheetNames)
            {
                if (Regex.IsMatch(sheet, @"EFUSE_BitDef_Table", RegexOptions.IgnoreCase))
                    flag[8] = true;

                if (Regex.IsMatch(sheet, @"UDR", RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(sheet, @"Revision", RegexOptions.IgnoreCase))
                {
                    if (Regex.IsMatch(sheet, @"USO|USI", RegexOptions.IgnoreCase))
                    {
                        flag[0] = true;
                    }
                    else
                    {
                        var wsUdr = workbook.Worksheets[sheet];

                        if (wsUdr.Dimension == null) continue;

                        var fmtCheck = 0;
                        var maxCol = wsUdr.Dimension.Columns < 40 ? wsUdr.Dimension.Columns : 40;
                        for (var iCol = 1; iCol < maxCol; iCol++)
                            if (Regex.IsMatch(wsUdr.Cells[2, iCol].Text, @"^Field|^Width", RegexOptions.IgnoreCase) ||
                                Regex.IsMatch(wsUdr.Cells[2, iCol].Text, @"USO|USI|Silicon", RegexOptions.IgnoreCase))
                                fmtCheck += 1;
                        if (fmtCheck >= 4)
                            flag[1] = true;
                    }
                }

                if (Regex.IsMatch(sheet, @"Sensor", RegexOptions.IgnoreCase))
                {
                    var wsSensor = workbook.Worksheets[sheet];

                    if (wsSensor.Dimension == null) continue;


                    //Direct-Access Mode (1-bit data)
                    if (Regex.IsMatch(wsSensor.Cells[1, 1].Text, @"^Direct", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(wsSensor.Cells[1, 1].Text, @"Access|Mode|bit|data", RegexOptions.IgnoreCase))
                        flag[3] = true;
                    else flag[2] = true;
                }
                else if (Regex.IsMatch(sheet, @"Mon", RegexOptions.IgnoreCase))
                {
                    var wsMon = workbook.Worksheets[sheet];

                    if (wsMon.Dimension == null) continue;


                    //Direct-Access Mode (1-bit data)
                    if (Regex.IsMatch(wsMon.Cells[1, 1].Text, @"^Direct", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(wsMon.Cells[1, 1].Text, @"Access|Mode|bit|data", RegexOptions.IgnoreCase))
                        flag[3] = true;
                }

                if (Regex.IsMatch(sheet, @"Revision", RegexOptions.IgnoreCase))
                {
                    var wsRevision = workbook.Worksheets[sheet];

                    if (wsRevision.Dimension == null) continue;


                    if (Regex.IsMatch(wsRevision.Cells[2, 1].Text, @"Fuse|Revision", RegexOptions.IgnoreCase) &&
                        !Regex.IsMatch(wsRevision.Cells[2, 2].Text, @"^b", RegexOptions.IgnoreCase)) revisionCheck += 1;
                    revisionCount++;
                }

                if (Regex.IsMatch(sheet, @"IDS", RegexOptions.IgnoreCase))
                {
                    var wsIds = workbook.Worksheets[sheet];

                    if (wsIds.Dimension == null) continue;


                    if (Regex.IsMatch(wsIds.Cells[1, 2].Text, @"CFG", RegexOptions.IgnoreCase)) flag[6] = true;
                    else flag[7] = true;
                }

                var ws = workbook.Worksheets[sheet];

                if (ws.Dimension == null) continue;
                if (Regex.IsMatch(ws.Cells[1, 1].Text, @"^Direct", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(ws.Cells[1, 1].Text, @"Access|Mode|bit|data", RegexOptions.IgnoreCase))
                    accessBitCount += 1;
                // if (Regex.IsMatch(ws.Cells[1, iCol].Text,@".*Force\s*Condition.*",RegexOptions.IgnoreCase)) fmtCheck += 1 << 1;
            }

            if (revisionCount > 1 && revisionCheck > 0)
            {
                if (revisionCheck > revisionCount - 2)
                    flag[4] = true;
                else
                    flag[5] = true;
            }
            else if (revisionCount > 0 && revisionCheck > 0)
            {
                if (revisionCheck > revisionCount - 1)
                    flag[4] = true;
                else
                    flag[5] = true;
            }

            if (flag[8])
                revision = 2;
            else if (flag[1] || flag[3] || flag[5] || flag[7])
                revision = 1;
            else if (flag[0] || flag[2] || flag[4] || flag[6])
                revision = 0;
            else
                revision = accessBitCount >= 1 ? 1 : 0;
            return revision;
        }

        protected override bool IsValidSheet(ExcelWorksheet sheet)
        {
            var name = sheet.Name;
            if (Regex.IsMatch(name,
                    @"^EVS|^IO|^POWER|^HARDIP_|^AC_|^DC_|^DCTEST_|TESTSETTING|VOLTAGETABLE|^EFUSE_|^BINCUT_|^Wireless_|^LCD_|^otp_|^ahb_|^PLLDEBUG_",
                    RegexOptions.IgnoreCase))
                return true;
            return false;
        }
    }
}