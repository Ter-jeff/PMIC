namespace PmicAutogen.Local.Version
{
    public class SrcInfoRow
    {
        public SrcInfoRow(string inputFile, string comment)
        {
            InputFile = inputFile;
            Comment = comment;
        }

        public string InputFile { get; set; }
        public string Comment { get; set; }
    }
}