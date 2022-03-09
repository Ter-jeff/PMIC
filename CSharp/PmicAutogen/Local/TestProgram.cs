using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using PmicAutogen.Local.Version;

namespace PmicAutogen.Local
{
    public static class TestProgram
    {
        private static IgxlWorkBook _igxlWorkBk;
        private static NonIgxlSheets _nonIgxlSheetList;
        private static SrcInfoSheet _srcInfoSheet;
        private static VbtFunctionLib _vbtFunctionLib;

        public static IgxlWorkBook IgxlWorkBk
        {
            get { return _igxlWorkBk ?? (_igxlWorkBk = new IgxlWorkBook()); }
            set { _igxlWorkBk = value; }
        }

        public static NonIgxlSheets NonIgxlSheetsList => _nonIgxlSheetList ?? (_nonIgxlSheetList = new NonIgxlSheets());

        public static VbtFunctionLib VbtFunctionLib => _vbtFunctionLib ?? (_vbtFunctionLib = new VbtFunctionLib());

        public static SrcInfoSheet SourceInfoSheet => _srcInfoSheet ?? (_srcInfoSheet = new SrcInfoSheet());

        public static void Initialize()
        {
            _igxlWorkBk = new IgxlWorkBook();
            _nonIgxlSheetList = new NonIgxlSheets();
            _srcInfoSheet = new SrcInfoSheet();
            _vbtFunctionLib = new VbtFunctionLib();
        }
    }
}