using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.IgxlSheets
{
    public class WaveDefinitionSheet : IgxlSheet
    {
        public List<WaveDefRow> WaveDefRows;
        public WaveDefinitionSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            WaveDefRows = new List<WaveDefRow>();
            IgxlSheetName = IgxlSheetNameList.WaveDefinition;
        }

        public WaveDefinitionSheet(string sheetName)
            : base(sheetName)
        {
            WaveDefRows = new List<WaveDefRow>();
            IgxlSheetName = IgxlSheetNameList.WaveDefinition;
        }

        public void AddRow(WaveDefRow row)
        {
            WaveDefRows.Add(row);
        }

        public override void Write(string fileName, string version)
        {
            CreateFolder(Path.GetDirectoryName(fileName));
            var writer = new StreamWriter(fileName);

            const string firstline1 = "DTWaveDefinitionSheet,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            const string firstline2 = "Wave Definitions";
            writer.WriteLine(firstline1 + '\t' + firstline2);
            writer.WriteLine();
            writer.WriteLine("	WaveDefName	WaveDefType	WaveDef Component	Repeat Count	Relative Period	Relative Amplitude	Relative Offset	Primitive Parameters	Comment");
            foreach (var waveDefRow in WaveDefRows)
            {
                writer.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{8}",waveDefRow.ColumnA, waveDefRow.WaveDefName, waveDefRow.WaveDefType, waveDefRow.WaveDefComponent, waveDefRow.RepeatCount, waveDefRow.RelativePeriod, waveDefRow.RelativeAmplitude, waveDefRow.RelativeOffset, waveDefRow.PrimitiveParameters, waveDefRow.Comment);
            }
            writer.Close();
        }

        private void CreateFolder(string pFolder)
        {
            if (!Directory.Exists(pFolder))
            {
                Directory.CreateDirectory(pFolder);
            }
        }
    }
}
