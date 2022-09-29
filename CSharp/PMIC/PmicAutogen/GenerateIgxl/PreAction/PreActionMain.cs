using CommonLib.Enum;
using CommonLib.WriteMessage;
using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenBinTable;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenChannelMap;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap.PortMapModify;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using PmicAutogen.Local.Version;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.PreAction
{
    public class PreActionMain : MainBase
    {
        public void WorkFlow()
        {
            try
            {
                ReadFiles();

                if (StaticTestPlan.ChannelMapSheets != null && StaticTestPlan.ChannelMapSheets.Any())
                    foreach (var channelMapSheet in StaticTestPlan.ChannelMapSheets)
                        IgxlSheets.Add(channelMapSheet, FolderStructure.DirChannelMap);

                Response.Report("Generating PinMap ...", EnumMessageLevel.General, 40);
                var pinMapMain = new PinMapMain(StaticTestPlan.IoPinMapSheet, StaticTestPlan.IoPinGroupSheet,
                    StaticTestPlan.PortDefineSheet);
                var pinMapSheet = pinMapMain.GetPinMapSheet();
                IgxlSheets.Add(pinMapSheet ?? new PinMapSheet("PinMap"), FolderStructure.DirPinMap);

                var worksheet = InputFiles.SettingWorkbook.Worksheets[PmicConst.PortMap];
                if (worksheet != null)
                {
                    Response.Report("Generating PortMap ...", EnumMessageLevel.General, 50);
                    var igxlSheetReader = new IgxlSheetReader();
                    var portMapSheet = igxlSheetReader.GetPortMapSheet(worksheet);
                    var portSetList = new List<PortSet>();
                    new PortMapModifier().WorkFlow(portMapSheet, ref portSetList);
                    IgxlSheets.Add(portMapSheet, FolderStructure.DirPortMap);
                }

                Response.Report("Generating BinTable ...", EnumMessageLevel.General, 60);
                var binTableMain = new BinTableMain();
                IgxlSheets.Add(binTableMain.WorkFlow(), FolderStructure.DirBinTable);

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);

                Response.Report("Pre-Action Completed!", EnumMessageLevel.General, 100);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in pre-action of autogen. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        public void ReadFiles()
        {
            GenChannelFromExtraSheets();

            Response.Report("Initializing VBT functions ...", EnumMessageLevel.General, 40);
            var basMain = new BasMain(VersionControl.SrcInfoRows);
            basMain.WorkFlow(FolderStructure.DirOtherWaitForClassify);
        }

        private void GenChannelFromExtraSheets()
        {
            var extraFolder = LocalSpecs.ExtraPath;
            if (Directory.Exists(extraFolder))
            {
                var dir = new DirectoryInfo(extraFolder);
                var mExtraList = dir.GetFiles("*.txt");
                foreach (var extraTxt in mExtraList)
                    if (Regex.IsMatch(extraTxt.Name, "^Channel*", RegexOptions.IgnoreCase))
                    {
                        var channelMapMain = new ChannelMapMain();
                        var igxlSheet = channelMapMain.WorkFlow(extraTxt.FullName);
                        IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);
                    }
            }
        }
    }
}