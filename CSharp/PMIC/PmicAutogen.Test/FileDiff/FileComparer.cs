using System.Collections.Generic;

namespace PmicAutogen.Test.FileDiff
{
    public class FileComparer
    {
        public List<string> AddItems = new List<string>();
        public List<FileDiff> DiffItems = new List<FileDiff>();
        public List<string> MissingItems = new List<string>();
        public string Output { get; internal set; }
        public string Expected { get; internal set; }
    }
}