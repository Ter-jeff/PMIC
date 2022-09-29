using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class WriterAhbRegisterMap
    {
        private readonly AhbRegisterMapSheet _ahbRegSheet;

        public WriterAhbRegisterMap(AhbRegisterMapSheet ahbRegSheet)
        {
            _ahbRegSheet = ahbRegSheet;
        }

        public void OutPutAhbRegisterMap(string path, string sheetName)
        {
            if (_ahbRegSheet == null || _ahbRegSheet.AhbRegRows == null)
                return;

            var fileNameTxt = Path.Combine(path, sheetName + ".txt");

            Directory.CreateDirectory(path);

            if (File.Exists(fileNameTxt)) File.Delete(fileNameTxt);

            using (var writer = new StreamWriter(fileNameTxt))
            {
                WriteTxtLine(writer, _ahbRegSheet.Headers);
                foreach (var regRow in _ahbRegSheet.AhbRegRows)
                    WriteTxtLine(writer, regRow.GetRegData());
            }

            TestProgram.NonIgxlSheetsList.Add(path, sheetName);
        }

        public void OutPutAhbEnum(string path)
        {
            if (_ahbRegSheet == null || _ahbRegSheet.AhbRegRows == null)
                return;

            var fileList = _ahbRegSheet.WriteAhbEnum(FolderStructure.DirOtp);
            var basMain = new BasMain(VersionControl.SrcInfoRows);
            foreach (var file in fileList)
                basMain.AddComment(file);
        }

        private void WriteTxtLine(StreamWriter streamWriter, List<string> lines)
        {
            streamWriter.WriteLine(string.Join("\t", lines.Select(x => x.Replace("\n", " "))));
        }
    }
}