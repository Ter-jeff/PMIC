using System.Collections.Generic;

namespace Library.Function.ErrorReport
{
    public class Error
    {
        public object ErrorType { set; get; }
        public List<string> Comments { get; set; }
        public string Link => InteropExcel.GetHyperlink(SheetName, RowNum, ColNum, "Link");
        public string SheetName { get; set; }
        public ErrorLevel ErrorLevel { get; set; }
        public int RowNum { get; set; }
        public int ColNum { get; set; }
        public string Message { get; set; }
    }
}