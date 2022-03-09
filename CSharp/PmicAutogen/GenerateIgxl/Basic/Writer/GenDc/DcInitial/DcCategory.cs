namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.DcInitial
{
    public class DcCategory
    {
        public DcCategory(string categoryName, string block, string subCat, DcCategoryType type,
            string targetDcSpec = "")
        {
            CategoryName = categoryName;
            Block = block;
            SubCategory = subCat;
            Type = type;
            DcSpecSheet = targetDcSpec;
        }

        public DcCategoryType Type { set; get; }
        public string CategoryName { set; get; } //"Category" not always equal to "Block" + "SubCategory"
        public string Block { set; get; } //Cpu | Gfx | Soc | DDR
        public string SubCategory { set; get; } //MC601 | MG001 | MS001 | MD001 | Vmargin1
        public string DcSpecSheet { set; get; } //C | S | G | H | R
    }
}