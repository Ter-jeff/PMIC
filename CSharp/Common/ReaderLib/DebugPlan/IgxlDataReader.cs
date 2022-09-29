using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using Ionic.Zip;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CommonReaderLib.DebugPlan
{
    public class IgxlDataReader
    {
        public List<DcSpecSheet> DcSpecSheets = new List<DcSpecSheet>();
        public List<string> TimeSetBasicSheets = new List<string>();
        public List<PinMapSheet> PinMapSheets = new List<PinMapSheet>();
        public List<AcSpecSheet> AcSpecSheets = new List<AcSpecSheet>();

        public IgxlDataReader(string testProgram)
        {
            var igxlSheetReader = new IgxlSheetReader();
            using (var zip = new ZipFile(testProgram))
            {
                var zipArchiveEntries = zip.Entries.ToList();
                foreach (var zipArchiveEntry in zipArchiveEntries)
                {
                    var sheetName = Path.GetFileNameWithoutExtension(zipArchiveEntry.FileName);
                    var stream = zipArchiveEntry.OpenReader();
                    string firstLine;
                    using (var sr = new StreamReader(stream))
                        firstLine = sr.ReadLine();

                    var sheetType = igxlSheetReader.GetIgxlSheetType(firstLine);
                    if (sheetType == SheetTypes.DTTimesetBasicSheet)
                        TimeSetBasicSheets.Add(sheetName);
                    else if (sheetType == SheetTypes.DTDCSpecSheet)
                        DcSpecSheets.Add(new ReadDcSpecSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTACSpecSheet)
                        AcSpecSheets.Add(new ReadAcSpecSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTPinMap)
                        PinMapSheets.Add(new ReadPinMapSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                }
            }
        }
    }
}