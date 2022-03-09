using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutomationCommon.Utility
{
    public class Igxl
    {
        public List<string> SelectFileList(string sheetType, string exportDir, string fileName = "")
        {
            var files = Directory.GetFiles(exportDir);
            if (string.IsNullOrEmpty(fileName))
                return files.Where(x => ReadSheetHeader(x).StartsWith(sheetType, StringComparison.OrdinalIgnoreCase)).ToList();
            return files.Where(x => ReadSheetHeader(x).StartsWith(sheetType, StringComparison.OrdinalIgnoreCase))
                .Where(x => x.EndsWith(fileName, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private string ReadSheetHeader(string file)
        {
            var reader = new StreamReader(file);
            var header = reader.ReadLine();
            reader.Close();
            return header ?? "";
        }
    }
}
