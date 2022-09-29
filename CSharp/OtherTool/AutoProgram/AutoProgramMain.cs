using AutoProgram.Writer;
using CommonReaderLib.DebugPlan;
using IgxlData.IgxlManager;
using IgxlData.IgxlSheets;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoProgram
{
    public class AutoProgramMain
    {
        public static Logger Logger = LogManager.GetCurrentClassLogger();
        public static string EnableWord = "";

        public string Main(string jobName, string testProgram, string patternFolder,
            string enableWords, DebugPlanMain debugTestPlan)
        {
            try
            {
                var tempFolder = Path.Combine(Path.GetDirectoryName(testProgram), "Temp");
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);
                var outputIgxl = Path.Combine(Path.GetDirectoryName(testProgram),
                    Path.GetFileNameWithoutExtension(debugTestPlan.InputFile) + ".igxl");
                if (File.Exists(outputIgxl))
                    File.Delete(outputIgxl);

                var igxlData = new IgxlDataReader(testProgram, jobName);

                var igxlSheets = new List<IgxlSheet>();
                Logger.Trace(@"Updating TimeSet ...");
                var updateTimeSet = new UpdateTimeSet().Work(debugTestPlan, igxlData.TimeSetBasicSheets,
                    igxlData.PortMapSheet, patternFolder);
                igxlSheets.AddRange(updateTimeSet);

                Logger.Trace(@"Updating AC spec ...");
                var updateAcSpecs = new UpdateAcSpecs().Work(igxlData.CurrentAcSpecSheet, updateTimeSet);
                igxlSheets.Add(updateAcSpecs);

                Logger.Trace(@"Generating PatSets_All_CZ ...");
                var patSetsAllCz = debugTestPlan.GenPatSetAllSheet(patternFolder);
                igxlSheets.Add(patSetsAllCz);

                Logger.Trace(@"Updating PatSets_All ...");
                var patSetsAll = igxlData.PatSetsAll.Remove(patSetsAllCz.PatSets.Select(x => x.PatSetName));
                igxlSheets.Add(patSetsAll);

                Logger.Trace(@"Updating Pattern_Subroutine ...");
                var patSetSubRows = debugTestPlan.GenPatSetSubRows(patternFolder);
                if (patSetSubRows.Any())
                {
                    igxlData.PatSetSubSheet.AddRows(patSetSubRows);
                    igxlSheets.Add(igxlData.PatSetSubSheet);
                }

                Logger.Trace(@"Generating PatSets_CZ ...");
                var patSetCz = debugTestPlan.GenPatSetSheet(patternFolder);
                igxlSheets.Add(patSetCz);

                Logger.Trace(@"Generating VBT_LIB_PV.bas ...");
                var basFiles = debugTestPlan.GenBas(enableWords, testProgram, tempFolder);
                igxlSheets.AddRange(basFiles);

                Logger.Trace(@"Generating Inst_CZ ...");
                var instanceSheet = debugTestPlan.GenInstSheet(igxlData.VbtFunctionLib, updateAcSpecs);
                igxlSheets.Add(instanceSheet);

                Logger.Trace(@"Generating Flow_CZ ...");
                var flowSheet = debugTestPlan.GenFlowSheet();
                igxlSheets.Add(flowSheet);

                Logger.Trace(@"Generating DFC_List ...");
                var DFC_List = debugTestPlan.GenDfcList();
                igxlSheets.Add(DFC_List);

                Logger.Trace(@"Generating Char_CZ ...");
                var charSheet = debugTestPlan.GenCharSheet(igxlData.CurrentPinMapSheet, igxlData.CurrentAcSpecSheet);
                igxlSheets.Add(charSheet);

                Logger.Trace(@"Updating JobList ...");
                var updateJobList = new UpdateJobList().Work(igxlData.JobListSheet, patSetsAllCz.SheetName,
                    patSetCz.SheetName,
                    instanceSheet.SheetName, charSheet.SheetName);
                igxlSheets.Add(updateJobList);

                Logger.Trace(@"Updating GlobalSpecSheet ...");
                var pins = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                    .SelectMany(x => x.Pins).Select(x => x.ShmooName).Distinct().ToList();
                var updateGlobalSpecSheet = new UpdateGlobalSpecSheet().Work(igxlData.GlobalSpecSheet, pins);
                igxlSheets.Add(updateGlobalSpecSheet);

                Logger.Trace(@"Inserting Char_CZ into Main_Flow ...");
                var mainFlow = igxlData.JobListSheet.Rows
                    .Find(x => x.JobName.Equals(jobName, StringComparison.CurrentCultureIgnoreCase)).FlowTable;
                var mainFlowSheet = igxlData.FlowSheets.Find(x =>
                    x.SheetName.Equals(mainFlow, StringComparison.CurrentCultureIgnoreCase));
                var updateMainFlow = new UpdateMainFlow().Work(mainFlowSheet, flowSheet.SheetName);
                igxlSheets.Add(updateMainFlow);

                Logger.Trace(@"Generating test program {0} ...", outputIgxl);
                File.Copy(testProgram, outputIgxl, true);
                var igxlManagerMain = new IgxlManagerMain();
                igxlManagerMain.AddIgxlSheets(outputIgxl, igxlSheets, tempFolder);

                return outputIgxl;
            }
            catch (Exception e)
            {
                Logger.Error(e.StackTrace);
            }
            return "";
        }
    }
}