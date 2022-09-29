using PmicAutogen.InputPackages.Inputs;
using System.IO;
using System.Linq;

namespace PmicAutogen.InputPackages.Base
{
    public abstract class Input
    {
        protected Input(FileInfo fileInfo, InputFileType inputFileType)
        {
            FullName = fileInfo.FullName;
            FileType = inputFileType;
            Selected = true;
        }

        public string FullName { get; set; }
        public InputFileType FileType { get; set; }
        public bool Selected { set; get; }

        public string GetProjectName()
        {
            var fileName = Path.GetFileName(FullName);
            if (fileName != null) return fileName.Split('_').First();
            return "";
        }
    }
}