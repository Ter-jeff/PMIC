namespace CommonLib.Controls
{

    public class FileFilter
    {
        public string GetFilter(EnumFileFilter filter)
        {
            if (filter == EnumFileFilter.IdsDistribution)
            {
                return @"IDS Distribution" + @"(IDS_Distribution.txt)|IDS_Distribution.txt*";
            }

            if (filter == EnumFileFilter.BinCut)
            {
                return @"*Bin*Cut*" + @"(*.txt,*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx";
            }

            if (filter == EnumFileFilter.TestProgram)
            {
                return "(*.igxl*,*.xls*)|*.xls*;*.igxl*";
            }

            if (filter == EnumFileFilter.Igxl)
            {
                return "(*.igxl*)|*.igxl*";
            }

            if (filter == EnumFileFilter.PatternCsv)
            {
                return "(*pattern*.csv)|*pattern*.csv";
            }

            if (filter == EnumFileFilter.Excel)
            {
                return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";
            }

            if (filter == EnumFileFilter.TemplateFile)
            {
                return "(*.tmp)|*.tmp";
            }

            if (filter == EnumFileFilter.Txt)
            {
                return "(*.txt)|*.txt";
            }

            if (filter == EnumFileFilter.BasFile)
            {
                return "(*.bas,*.txt)|*.bas;*.txt";
            }

            if (filter == EnumFileFilter.PaFile)
            {
                return "(*.csv,*.xls*)|*.csv;*.xlsx;*.xlsm";
            }

            if (filter == EnumFileFilter.XmlFile)
            {
                return "(*.xml)|*.xml";
            }

            if (filter == EnumFileFilter.OTPFile)
            {
                return "(*.csv)|*.csv";
            }

            return "";
        }
    }

    public enum EnumFileFilter
    {
        Excel,
        TemplateFile,
        BasFile,
        PaFile,
        XmlFile,
        PatternCsv,
        TestProgram,
        Txt,
        Igxl,
        BinCut,
        IdsDistribution,
        OTPFile
    }
}