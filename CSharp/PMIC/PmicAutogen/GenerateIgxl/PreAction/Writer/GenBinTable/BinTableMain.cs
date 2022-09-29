using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.Setting.BinNumber;

namespace PmicAutogen.GenerateIgxl.PreAction.Writer.GenBinTable
{
    public class BinTableMain
    {
        #region Member Function

        public BinTableSheet WorkFlow()
        {
            var binTableSheet = new BinTableSheet(IgxlWorkBook.MainBinTableName);
            binTableSheet.AddRow(GetSystemError());
            return binTableSheet;
        }

        private BinTableRow GetSystemError()
        {
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.SystemError, "system error");
            BinNumberRuleRow bin;
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out bin);
            var binRow = new BinTableRow();
            binRow.Name = "Bin_System_Error";
            binRow.Op = "AND";
            binRow.Sort = bin.CurrentSoftBin.ToString();
            binRow.Bin = bin.HardBin;
            binRow.Result = "Fail";
            binRow.Items.Add("T");
            return binRow;
        }

        #endregion
    }
}