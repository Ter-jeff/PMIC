using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PmicAutogen.Local.Version
{
    public class SrcInfoSheet
    {
        #region Constructor

        public SrcInfoSheet()
        {
            FileName = "InputFilesInfo.txt";
            _srcInfoRows = new List<SrcInfoRow>();
        }

        #endregion

        #region Property

        public string FileName { get; set; }

        #endregion

        #region Field

        private StreamWriter _sheetWriter;
        private readonly List<SrcInfoRow> _srcInfoRows;

        #endregion

        #region Member Function

        public void AddSrcInfo(SrcInfoRow srcInfoRow)
        {
            _srcInfoRows.Add(srcInfoRow);
        }

        public void Print(string tarDir)
        {
            if (!Directory.Exists(tarDir))
                Directory.CreateDirectory(tarDir);
            var fullFileName = GetFullFileName(tarDir);
            _sheetWriter = new StreamWriter(fullFileName);
            WriteHeader();
            WriteColumnHeader();
            WriteRows();
            CloseStreamWriter();
        }

        private void WriteRows()
        {
            foreach (var srcInfoRow in _srcInfoRows)
            {
                var srcRow = new StringBuilder();
                if (srcInfoRow == null) continue;
                srcRow.Append(srcInfoRow.InputFile);
                srcRow.Append("\t");
                srcRow.Append(srcInfoRow.Comment);
                srcRow.Append("\t");
                var str = Regex.Replace(srcRow.ToString(), @"[^\u0000-\u007F]", "?");
                _sheetWriter.WriteLine(str);
            }
        }

        private void WriteHeader()
        {
        }

        private void WriteColumnHeader()
        {
            const string header = "Input File\tComment\t";
            _sheetWriter.WriteLine(header);
        }

        public string GetFullFileName(string tarDir)
        {
            return Path.Combine(tarDir, FileName);
        }

        private void CloseStreamWriter()
        {
            if (_sheetWriter != null)
                _sheetWriter.Close();
        }

        #endregion
    }
}