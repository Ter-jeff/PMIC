using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.IgxlSheets
{
    public class MixedSignalSheet : IgxlSheet
    {
        #region Field
        public List<MixedSigRow> MixedSigRows;
        #endregion

        public MixedSignalSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            MixedSigRows = new List<MixedSigRow>();
            IgxlSheetName = IgxlSheetNameList.WaveDefinition;
        }

        public MixedSignalSheet(string sheetName)
            : base(sheetName)
        {
            MixedSigRows = new List<MixedSigRow>();
            IgxlSheetName = IgxlSheetNameList.WaveDefinition;
        }

        public void AddRow(MixedSigRow row)
        {
            MixedSigRows.Add(row);
        }

        public void AddRows(List<MixedSigRow> rows)
        {
            MixedSigRows.AddRange(rows);
        }

        public override void Write(string fileName, string version)
        {
            CreateFolder(Path.GetDirectoryName(fileName));
            var writer = new StreamWriter(fileName);

            const string firstline1 = "DTMixedSignalTimingSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            const string firstline2 = "Mixed Signal Timing";
            writer.WriteLine(firstline1 + '\t' + firstline2);
            writer.WriteLine();
            writer.WriteLine("			Resource		Clocking					Instrument	Waveform		MSW	Data for Pre v5.0 Upgrades			");
            writer.WriteLine("\tSet Name	Subset	Type	ID	Fs	N	Fr	M	USR	Data	Definition	Filter	Settings	WaveName	Amplitude	Offset	Old Instrument Data	Comment");
            foreach (var mixedSigRow in MixedSigRows)
            {
                writer.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}\t{13}\t{14}\t{15}\t{16}\t{17}\t{18}\t", mixedSigRow.ColumnA, mixedSigRow.Name,
                    mixedSigRow.Subset, mixedSigRow.Type, mixedSigRow.ID, mixedSigRow.Fs, mixedSigRow.N, mixedSigRow.Fr,
                    mixedSigRow.M, mixedSigRow.USR, mixedSigRow.Data, mixedSigRow.Definition, mixedSigRow.Filter, mixedSigRow.Settings,
                    mixedSigRow.WaveName, mixedSigRow.Amplitude, mixedSigRow.Offset, mixedSigRow.OldInstData, mixedSigRow.Comment);
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
