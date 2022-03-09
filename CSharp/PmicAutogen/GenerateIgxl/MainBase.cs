using System;
using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl
{
    public abstract class MainBase : IDisposable
    {
        protected Dictionary<IgxlSheet, string> IgxlSheets;


        public void Initialize()
        {
            IgxlSheets = new Dictionary<IgxlSheet, string>();
        }

        public void Dispose()
        {
            GC.Collect();
            for (var j = 0; j < GC.MaxGeneration; j++)
            {
                GC.Collect(j);
                GC.WaitForPendingFinalizers();
            }
            GC.SuppressFinalize(this);
        }

        public void AddIgxlSheets(Dictionary<IgxlSheet, string> igxlSheets)
        {
            foreach (var igxlSheet in igxlSheets)
            {
                if (igxlSheet.Key is ChannelMapSheet)
                    TestProgram.IgxlWorkBk.AddChannelMapSheet(igxlSheet.Value, (ChannelMapSheet)igxlSheet.Key);
                else if (igxlSheet.Key is PinMapSheet)
                    TestProgram.IgxlWorkBk.PinMapPair =
                        new KeyValuePair<string, PinMapSheet>(igxlSheet.Value, (PinMapSheet)igxlSheet.Key);
                else if (igxlSheet.Key is PortMapSheet)
                    TestProgram.IgxlWorkBk.AddPortMapSheet(igxlSheet.Value, (PortMapSheet)igxlSheet.Key);
                else if (igxlSheet.Key is SubFlowSheet)
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(igxlSheet.Value, (SubFlowSheet)igxlSheet.Key);
                else if (igxlSheet.Key is BinTableSheet)
                    TestProgram.IgxlWorkBk.AddBinTblSheet(igxlSheet.Value, (BinTableSheet)igxlSheet.Key);
                else if (igxlSheet.Key is InstanceSheet)
                    TestProgram.IgxlWorkBk.AddInsSheet(igxlSheet.Value, (InstanceSheet)igxlSheet.Key);
                else if (igxlSheet.Key is PatSetSheet)
                    TestProgram.IgxlWorkBk.AddPatSetSheet(igxlSheet.Value, (PatSetSheet)igxlSheet.Key);
                else if (igxlSheet.Key is GlobalSpecSheet)
                    TestProgram.IgxlWorkBk.GlbSpecSheetPair =
                        new KeyValuePair<string, GlobalSpecSheet>(FolderStructure.DirGlbSpec,
                            (GlobalSpecSheet)igxlSheet.Key);
                else if (igxlSheet.Key is DcSpecSheet)
                    TestProgram.IgxlWorkBk.AddDcSpecSheet(igxlSheet.Value, (DcSpecSheet)igxlSheet.Key);
                else if (igxlSheet.Key is AcSpecSheet)
                    TestProgram.IgxlWorkBk.AddAcSpecSheet(igxlSheet.Value, (AcSpecSheet)igxlSheet.Key);
                else if (igxlSheet.Key is LevelSheet)
                    TestProgram.IgxlWorkBk.AddLevelSheet(igxlSheet.Value, (LevelSheet)igxlSheet.Key);
                else if (igxlSheet.Key is TimeSetBasicSheet)
                    TestProgram.IgxlWorkBk.AddTimeSetSheet(igxlSheet.Value, (TimeSetBasicSheet)igxlSheet.Key);
                else if (igxlSheet.Key is PatSetSubSheet)
                    TestProgram.IgxlWorkBk.AddPatSetSubSheet(igxlSheet.Value, (PatSetSubSheet)igxlSheet.Key);
                else if (igxlSheet.Key is CharSheet)
                    TestProgram.IgxlWorkBk.AddCharSheet(igxlSheet.Value, (CharSheet)igxlSheet.Key);
            }
        }
    }
}