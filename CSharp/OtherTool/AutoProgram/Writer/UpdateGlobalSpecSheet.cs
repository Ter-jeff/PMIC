using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using System.Collections.Generic;

namespace AutoProgram.Writer
{
    public class UpdateGlobalSpecSheet
    {
        public GlobalSpecSheet Work(GlobalSpecSheet globalSpecSheet, List<string> pins)
        {
            if (globalSpecSheet == null)
                globalSpecSheet = new GlobalSpecSheet("Global Specs");

            foreach (var pin in pins)
            {
                var row = new GlobalSpec(pin);
                row.Value = "0";
                if (!globalSpecSheet.IsExist(pin))
                    globalSpecSheet.GlobalSpecsRows.Add(row);
            }

            return globalSpecSheet;
        }
    }
}