namespace CommonReaderLib
{
    public abstract class MyRow
    {
        public int RowNum;
        public string SheetName;

        public string SetColumnA()
        {
            return SheetName + ":Row" + RowNum;
        }
    }
}