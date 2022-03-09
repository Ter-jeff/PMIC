using PmicAutogen.Inputs.ScghFile;
using PmicAutogen.Inputs.ScghFile.Reader;

namespace PmicAutogen.Local
{
    public static class StaticScgh
    {
        public static ProdCharSheet ScghScanSheet;
        public static ProdCharSheet ScghMbistSheet;

        public static void AddSheets(ScghFileManager scghFileManager)
        {
            ScghScanSheet = scghFileManager.ScghScanSheet;
            ScghMbistSheet = scghFileManager.ScghMbistSheet;
        }
    }
}