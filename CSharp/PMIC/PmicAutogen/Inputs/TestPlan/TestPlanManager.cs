using CommonLib.Enum;
using CommonLib.ErrorReport;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PmicAutogen.Inputs.TestPlan
{
    public class TestPlanManager
    {
        public AhbRegisterMapSheet AhbRegisterMapSheet;
        public BscanCharSheet BscanCharSheet;
        public List<ChannelMapSheet> ChannelMapSheets = new List<ChannelMapSheet>();
        public DcTestContinuitySheet DcTestContinuitySheet;
        public List<TestPlanSheet> DcTestSheet = new List<TestPlanSheet>();
        public IfoldPowerTableSheet IfoldPowerTableSheet;
        public IoLevelsSheet IoLevelsSheet;
        public IoPinGroupSheet IoPinGroupSheet;
        public PinMapSheet IoPinMapSheet;
        public SubFlowSheet MainFlowSheet;
        public OtpSetupSheet OtpSetupSheet;
        public PmicIdsSheet PmicIdsSheet;
        public PmicLeakageSheet PmicLeakageSheet;
        public PortDefineSheet PortDefineSheet;
        public PowerOverWriteSheet PowerOverWriteSheet;
        public VddLevelsSheet VddLevelsSheet;

        public void CheckAll(ExcelWorkbook workbook, Application application)
        {
            try
            {
                #region Pre check
                if (new PreCheckIfoldPowerTable(workbook, PmicConst.IfoldPwrTable).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.IfoldPwrTable);
                    IfoldPowerTableSheet =
                        new IfoldPowerTableReader().ReadSheet(workbook.Worksheets[PmicConst.IfoldPwrTable]);
                }

                if (new PreCheckPowerOverWrite(workbook, PmicConst.PowerOverWrite).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.PowerOverWrite);
                    PowerOverWriteSheet =
                        new PowerOverWriteReader().ReadFlowMain(workbook.Worksheets[PmicConst.PowerOverWrite]);
                }

                if (new PreCheckPinMap(workbook, PmicConst.IoPinMap).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.IfoldPwrTable);
                    IoPinMapSheet = new IoPinMapReader().ReadSheet(workbook.Worksheets[PmicConst.IoPinMap]);
                }

                if (new PreCheckIoPinGroup(workbook, PmicConst.IoPinGroup).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.IoPinGroup);
                    IoPinGroupSheet = new IoPinGroupReader().ReadSheet(workbook.Worksheets[PmicConst.IoPinGroup]);
                }

                var channelMaps = workbook.Worksheets.Where(x =>
                    Regex.IsMatch(x.Name, "^" + PmicConst.ChannelMap, RegexOptions.IgnoreCase));
                foreach (var channelMap in channelMaps)
                {
                    application.StatusBar = string.Format("Checking {0} ...", channelMap.Name);
                    if (new PreCheckChannelMap(workbook, channelMap.Name).CheckMain())
                        ChannelMapSheets.Add(new ReadChanMapSheet().GetSheet(workbook.Worksheets[channelMap.Name]));
                }

                if (new PreCheckPortDefine(workbook, PmicConst.PortDefine).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.PortDefine);
                    PortDefineSheet = new PortDefineReader().ReadSheet(workbook.Worksheets[PmicConst.PortDefine]);
                }

                if (new PreCheckVddLevels(workbook, PmicConst.VddLevels).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.VddLevels);
                    VddLevelsSheet = new VddLevelsReader().ReadSheet(workbook.Worksheets[PmicConst.VddLevels]);
                }

                if (new PreCheckIoLevels(workbook, PmicConst.IoLevels).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.IoLevels);
                    IoLevelsSheet = new IoLevelsReader().ReadSheet(workbook.Worksheets[PmicConst.IoLevels]);
                }

                if (new PreCheckDcTestContinuity(workbook, PmicConst.DcTestContinuity).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.IfoldPwrTable);
                    DcTestContinuitySheet =
                        new DcTestContinuityReader().ReadSheet(workbook.Worksheets[PmicConst.DcTestContinuity]);
                }

                if (new PreCheckPmicIds(workbook, PmicConst.PmicIds).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.PmicIds);
                    PmicIdsSheet = new PmicIdsReader().ReadSheet(workbook.Worksheets[PmicConst.PmicIds]);
                }

                var leakage = workbook.Worksheets.First(x =>
                    x.Name.EndsWith("_" + PmicConst.Leakage, StringComparison.CurrentCultureIgnoreCase)).Name;
                if (string.IsNullOrEmpty(leakage))
                    leakage = PmicConst.PmicLeakage;
                if (new PreCheckPmicLeakage(workbook, leakage).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", leakage);
                    PmicLeakageSheet = new PmicLeakageReader().ReadSheet(workbook.Worksheets[leakage]);
                }

                if (new PreCheckBscanChar(workbook, PmicConst.BscanChar).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.BscanChar);
                    BscanCharSheet = new BscanCharReader().ReadSheet(workbook.Worksheets[PmicConst.BscanChar]);
                }

                var dcTests = workbook.Worksheets.Where(x =>
                    Regex.IsMatch(x.Name, "^" + PmicConst.DctTest, RegexOptions.IgnoreCase));
                foreach (var dcTest in dcTests)
                {
                    if (dcTest.Name.Equals(PmicConst.DcTestContinuity, StringComparison.CurrentCultureIgnoreCase))
                        continue;
                    application.StatusBar = string.Format("Checking {0} ...", dcTest.Name);

                    if (new PreCheckHardip(workbook, dcTest.Name).CheckMain())
                        DcTestSheet.Add(new TestPlanReader().ReadSheet(workbook.Worksheets[dcTest.Name]));
                }

                if (new PreCheckAhbRegisterMap(workbook, PmicConst.AhbRegisterMap).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.AhbRegisterMap);
                    AhbRegisterMapSheet = new AhbRegisterMapReader().ReadSheet(workbook.Worksheets[PmicConst.AhbRegisterMap]);
                }

                if (new PreCheckGenMainFlow(workbook, PmicConst.GenMainFlow).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.GenMainFlow);
                    MainFlowSheet = new ReadFlowSheet().GetSheet(workbook.Worksheets[PmicConst.GenMainFlow]);
                }

                if (new PreCheckOtpSetup(workbook, PmicConst.OtpSetup).CheckMain())
                {
                    application.StatusBar = string.Format("Checking {0} ...", PmicConst.OtpSetup);
                    OtpSetupSheet = new OtpSetupReader().ReadSheet(workbook.Worksheets[PmicConst.OtpSetup]);
                }

                #endregion

                #region Post check
                var pins = new List<string>();
                var pinGroupsByIoPinGroup = new List<string>();
                var pinGroupsByIoPinMap = new List<string>();
                if (IoPinMapSheet != null)
                    pins = IoPinMapSheet.PinList.Select(x => x.PinName).ToList();
                var totalPins = pins;
                if (IoPinMapSheet != null)
                    pinGroupsByIoPinMap = IoPinMapSheet.GroupList.Select(x => x.PinName).ToList();
                if (IoPinMapSheet != null)
                    pinGroupsByIoPinGroup = IoPinGroupSheet.Rows.Select(x => x.PinGroupName).ToList();
                totalPins.AddRange(pinGroupsByIoPinGroup);

                if (PortDefineSheet != null)
                {
                    var pinGroupsByPortDefine = PortDefineSheet.Rows.Select(x => x.ProtocolPortName).ToList();
                    totalPins.AddRange(pinGroupsByPortDefine);
                }

                totalPins = totalPins.Distinct().ToList();

                if (ChannelMapSheets != null)
                    foreach (var channelMapSheet in ChannelMapSheets)
                        foreach (var row in channelMapSheet.ChannelMapRows)
                        {
                            var pinName = row.DeviceUnderTestPinName;
                            if (!string.IsNullOrEmpty(pinName) &&
                                !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                    pinName);
                                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                    channelMapSheet.SheetName, row.RowNum, ChannelMapSheet.DeviceUnderTestPinName,
                                    errorMessage);
                            }
                        }

                if (IoPinMapSheet != null)
                    foreach (var pinGrp in IoPinMapSheet.GroupList)
                        foreach (var pin in pinGrp.PinList)
                        {
                            var grp = IoPinMapSheet.GroupList
                                .Where(o => o.PinName.Equals(pin.PinName, StringComparison.CurrentCultureIgnoreCase))
                                .Select(o => o).FirstOrDefault();
                            if (grp != null && grp.PinType != pinGrp.PinType)
                            {
                                var errorMessage = string.Format(
                                    "This pin group \"{0}\" include pin group \"{1}\", but the pin types are not matched!!!",
                                    pinGrp.PinName, grp.PinName);
                                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                    PmicConst.IoPinMap, 1, 1, errorMessage);
                            }
                        }

                if (IoPinGroupSheet != null)
                {
                    foreach (var row in IoPinGroupSheet.Rows)
                    {
                        var pinName = row.PinNameContainedInPinGroup;
                        if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                                x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                IoPinGroupSheet.SheetName, row.RowNum, IoPinGroupSheet.PinNameContainedInPinGroupIndex,
                                errorMessage);
                        }

                        var pinGroupName = row.PinGroupName;
                        if (!string.IsNullOrEmpty(pinName) && pinGroupsByIoPinMap.Exists(x =>
                                x.Equals(pinGroupName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format(
                                "This pin group \"{0}\"is already exist in pin map sheet !!!",
                                pinGroupName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                IoPinGroupSheet.SheetName, row.RowNum, IoPinGroupSheet.PinGroupNameIndex, errorMessage);
                        }
                    }

                    var pinGroups = IoPinGroupSheet.Rows.GroupBy(x => x.PinGroupName).ToList();
                    foreach (var pinGroup in pinGroups)
                    {
                        var pinGroupName = pinGroup.Key;
                        var type = IoPinMapSheet.GetPinType(pinGroup.First().PinNameContainedInPinGroup);
                        foreach (var row in pinGroup)
                        {
                            var pinType = IoPinMapSheet.GetPinType(row.PinNameContainedInPinGroup);
                            if (string.IsNullOrEmpty(pinType) && PortDefineSheet != null)
                                pinType = IoPinMapSheet.GetPinType(
                                    PortDefineSheet.GetFirstPin(row.PinNameContainedInPinGroup));
                            if (!type.Equals(pinType, StringComparison.CurrentCultureIgnoreCase))
                            {
                                var errorMessage = string.Format(
                                    "The Pin Type in the Group \"{0}\" not all content match. {1}",
                                    pinGroupName, row.PinNameContainedInPinGroup);
                                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                    IoPinGroupSheet.SheetName, row.RowNum, IoPinGroupSheet.PinGroupNameIndex,
                                    errorMessage);
                            }
                        }
                    }
                }

                if (PortDefineSheet != null)
                {
                    foreach (var row in PortDefineSheet.Rows)
                    {
                        var pinName = row.Pin;
                        if (!string.IsNullOrEmpty(pinName) &&
                            !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                PortDefineSheet.SheetName, row.RowNum, PortDefineSheet.PinIndex, errorMessage);
                        }
                    }

                    var emptyPinRows = PortDefineSheet.GetEmptyPinRows();
                    foreach (var emptyPinRow in emptyPinRows)
                    {
                        var errorMessage = "This pin name is empty !!!";
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            PortDefineSheet.SheetName, emptyPinRow.RowNum, PortDefineSheet.PinIndex, errorMessage);
                    }
                }

                if (VddLevelsSheet != null)
                    foreach (var row in VddLevelsSheet.Rows)
                    {
                        var pinName = row.WsBumpName;
                        if (!string.IsNullOrEmpty(pinName) &&
                            !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                VddLevelsSheet.SheetName, row.RowNum, VddLevelsSheet.WsBumpNameIndex, errorMessage);
                        }
                    }

                if (IoLevelsSheet != null)
                    foreach (var row in IoLevelsSheet.Rows)
                    {
                        var pinName = row.PinName;
                        if (!string.IsNullOrEmpty(pinName) &&
                            !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                IoLevelsSheet.SheetName, row.RowNum, IoLevelsSheet.PinNameIndex, errorMessage);
                        }
                    }

                if (DcTestContinuitySheet != null)
                    foreach (var row in DcTestContinuitySheet.Rows)
                    {
                        var pinName = row.PinGroup;
                        if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                                x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                DcTestContinuitySheet.SheetName, row.RowNum, DcTestContinuitySheet.PinGroupIndex,
                                errorMessage);
                        }
                    }

                if (PmicIdsSheet != null)
                    foreach (var row in PmicIdsSheet.Rows)
                    {
                        var pinName = row.MeasurePin;
                        if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                                x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!",
                                pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                PmicIdsSheet.SheetName, row.RowNum, PmicIdsSheet.MeasurePinIndex, errorMessage);
                        }
                    }

                if (PmicLeakageSheet != null)
                {
                    var inlegalInstanceNameRowList = PmicLeakageSheet.GetInlegalInstanceNameRows();
                    if (inlegalInstanceNameRowList.Any())
                    {
                        var caption = PmicLeakageSheet.SheetName + " inlegal instance name";
                        var inlegalInstanceNameStr = string.Join(",",
                            inlegalInstanceNameRowList.Select(o => o.InstanceName.Trim()));
                        var errorMsg = string.Format("These instance names are not legal: {0}", inlegalInstanceNameStr)
                                       + "\n" + "You can see more details in the ErrorReport.";
                        MessageBox.Show(errorMsg, caption, MessageBoxButton.OK, MessageBoxImage.Information);
                        foreach (var errorRow in inlegalInstanceNameRowList)
                        {
                            var errorMessage =
                                string.Format(
                                    "This instance name \"{0}\"is not legal, it must be end with \"low\" or \"high\"!!!",
                                    errorRow.InstanceName.Trim());
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                PmicLeakageSheet.SheetName, errorRow.RowNum, PmicLeakageSheet.InstanceNameIndex,
                                errorMessage);
                        }
                    }

                    //check whether measure pin is in pin map/group
                    var pinMapMain = new PinMapMain(IoPinMapSheet, IoPinGroupSheet, PortDefineSheet);
                    var pinMapSheet = pinMapMain.CreatePinMap(IoPinMapSheet, false);
                    foreach (var tpRow in PmicLeakageSheet.Rows)
                    {
                        var measPins = tpRow.MeasurePin.Split(',').ToList();
                        var forceVs = tpRow.ForceV.Split(',').ToList();
                        foreach (var measPinName in measPins)
                        {
                            var measPin = measPinName.Trim();
                            if (!string.IsNullOrEmpty(measPin))
                                if (!pinMapSheet.IsPinExist(measPin) && !pinMapSheet.IsGroupExist(measPin))
                                {
                                    var errorMessage =
                                        string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", measPin);
                                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                        PmicLeakageSheet.SheetName, tpRow.RowNum, PmicLeakageSheet.MeasurePinIndex,
                                        errorMessage);
                                }

                            if (pinMapSheet.IsPinExist(measPin))
                            {
                                if (measPin.EndsWith("_DM", StringComparison.CurrentCultureIgnoreCase) ||
                                    measPin.EndsWith("_DT", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    var errorMessage =
                                        string.Format("This pin \"{0}\" is end with '_DM' or '_DT', please confirm!!!",
                                            measPin);
                                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                        PmicLeakageSheet.SheetName, tpRow.RowNum, PmicLeakageSheet.MeasurePinIndex,
                                        errorMessage);
                                }
                            }
                            else if (pinMapSheet.IsGroupExist(measPin))
                            {
                                var pinsInGrp = pinMapSheet.GetPinsFromGroup(measPin);
                                var matchedPin = pinsInGrp.Find(o =>
                                    o.PinName.EndsWith("_DM", StringComparison.CurrentCultureIgnoreCase) ||
                                    o.PinName.EndsWith("_DT", StringComparison.CurrentCultureIgnoreCase));
                                if (matchedPin != null)
                                {
                                    var errorMessage =
                                        string.Format(
                                            "This pin group\"{0}\" contains pin which is end with '_DM' or '_DT', please confirm!!!",
                                            measPin);
                                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                        PmicLeakageSheet.SheetName, tpRow.RowNum, PmicLeakageSheet.MeasurePinIndex,
                                        errorMessage);
                                }
                            }
                        }

                        if (forceVs.Count > 1 && measPins.Count != forceVs.Count)
                        {
                            var errorMessage =
                                "This comma-separate ForveV format is not matched with comma-separate Measure Pin, they must have the same number of elements!!!";
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                PmicLeakageSheet.SheetName, tpRow.RowNum, PmicLeakageSheet.ForceVIndex,
                                errorMessage);
                        }
                    }

                    var legalInstanceNameAndTimeSetDic = PmicLeakageSheet.GetLegalInstanceNameAndTimeSet();
                    foreach (var legalInstanceNameAndTimeSetItem in legalInstanceNameAndTimeSetDic)
                        if (legalInstanceNameAndTimeSetItem.Value.Count > 1)
                            foreach (var leakageRowTuple in legalInstanceNameAndTimeSetItem.Value)
                            {
                                var errorMessage = "There are more than one TimeSet under the same Intance Name!!!";
                                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                    PmicLeakageSheet.SheetName, leakageRowTuple.Item2.RowNum,
                                    PmicLeakageSheet.TimeSetDefineIndex,
                                    errorMessage);
                            }
                }

                if (OtpSetupSheet != null && pins.Any())
                {
                    //List<string>  JTAGPortDefinePins = PortDefineSheet.Rows.Where(row => row.ProtocolPortName.Equals("NWIRE_JTAG", StringComparison.OrdinalIgnoreCase)).
                    //    Select(row => row.Pin).ToList();

                    var matchPattern = "^JTAG_([a-zA-Z]+)_Pin_Name$";
                    foreach (var tpRow in OtpSetupSheet.Rows)
                    {
                        var match = Regex.Match(tpRow.Variable, matchPattern);
                        if (match.Success)
                        {
                            var pinName = tpRow.Value.Trim();
                            if (!string.IsNullOrEmpty(pinName) && !pins.Exists(x =>
                                    x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                var errorMessage = string.Format("This pin \"{0}\"is not exist in IO_PinMap sheet !!!",
                                    pinName);
                                ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                    OtpSetupSheet.SheetName, tpRow.RowNum, OtpSetupSheet.ValueIndex,
                                    errorMessage);
                            }
                        }
                    }
                }

                if (VddLevelsSheet != null && BscanCharSheet != null)
                {
                    var domainNames = BscanCharSheet.GetDomainCurrentMapping().Keys.ToList();
                    foreach (var domainName in domainNames)
                    {
                        var row = VddLevelsSheet.Rows.Find(o =>
                            o.WsBumpName.Equals(domainName, StringComparison.CurrentCultureIgnoreCase));
                        var xRow = VddLevelsSheet.XRows.Find(o =>
                            o.WsBumpName.Equals(domainName, StringComparison.CurrentCultureIgnoreCase));
                        if (row == null && xRow == null)
                        {
                            var errorMessage =
                                string.Format(
                                    "The domain name \"{0}\" in BSCAN_CHAR sheet is not exist in VDD_Levels sheet's \"WS Bump Name\" column!!!",
                                    domainName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                VddLevelsSheet.SheetName, 0, VddLevelsSheet.WsBumpNameIndex,
                                errorMessage);
                            continue;
                        }

                        if (row != null && xRow == null)
                        {
                            var errorMessage =
                                string.Format(
                                    "The domain name \"{0}\" in BSCAN_CHAR sheet is matched with VDD_Levels sheet, but the Sequence column should be \"x\" or \"X\"!!!",
                                    domainName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                VddLevelsSheet.SheetName, row.RowNum, VddLevelsSheet.WsBumpNameIndex,
                                errorMessage);
                        }
                    }
                }

                //if (PmicLeakageSheet != null)
                //{
                //    //eg: MeasI pin = pin1,pin2               
                //    const string regMeasExpression =
                //        @"(?<MeasType>(Wi)*[(Meas)|(Src)]\S+)[\s]*(pin)?[\s]*=[\s]*(?<pin>(.*))";
                //    foreach (var row in PmicLeakageSheet.PatternRows)
                //    foreach (var patChildRows in row.PatChildRows)
                //        if (patChildRows is PatSubChildRow)
                //        {
                //            var childRow = (PatSubChildRow) patChildRows;
                //            foreach (var tpRow in childRow.TpRows)
                //            {
                //                var measStr = Regex.Match(tpRow.Meas, regMeasExpression, RegexOptions.IgnoreCase)
                //                    .Groups["pin"].ToString().Trim(',').Trim();
                //                foreach (var pinName in measStr.Split(','))
                //                    if (!string.IsNullOrEmpty(pinName.Trim()) && !totalPins.Exists(x =>
                //                        x.Equals(pinName.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                //                    {
                //                        var errorMessage =
                //                            string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName.Trim());
                //                        EpplusErrorManager.AddError(EnumErrorType.FormatError, ErrorLevel.Error,
                //                            PmicLeakageSheet.SheetName, tpRow.RowNum, PmicLeakageSheet.MeasIndex,
                //                            errorMessage);
                //                    }
                //            }
                //        }
                //}

                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void ReadAll(ExcelWorkbook workbook)
        {
            //HardIpDataMain.Initialize();

            #region Pre check

            if (new PreCheckIfoldPowerTable(workbook, PmicConst.IfoldPwrTable).CheckMain())
                IfoldPowerTableSheet =
                    new IfoldPowerTableReader().ReadSheet(workbook.Worksheets[PmicConst.IfoldPwrTable]);

            if (new PreCheckPowerOverWrite(workbook, PmicConst.PowerOverWrite).CheckMain())
                PowerOverWriteSheet =
                    new PowerOverWriteReader().ReadFlowMain(workbook.Worksheets[PmicConst.PowerOverWrite]);

            if (new PreCheckPinMap(workbook, PmicConst.IoPinMap).CheckMain())
                IoPinMapSheet = new IoPinMapReader().ReadSheet(workbook.Worksheets[PmicConst.IoPinMap]);

            if (new PreCheckIoPinGroup(workbook, PmicConst.IoPinGroup).CheckMain())
                IoPinGroupSheet = new IoPinGroupReader().ReadSheet(workbook.Worksheets[PmicConst.IoPinGroup]);

            var channelMaps = workbook.Worksheets.Where(x =>
                Regex.IsMatch(x.Name, "^" + PmicConst.ChannelMap, RegexOptions.IgnoreCase));
            foreach (var channelMap in channelMaps)
                if (new PreCheckChannelMap(workbook, channelMap.Name).CheckMain())
                    ChannelMapSheets.Add(new ReadChanMapSheet().GetSheet(workbook.Worksheets[channelMap.Name]));

            if (new PreCheckPortDefine(workbook, PmicConst.PortDefine).CheckMain())
                PortDefineSheet = new PortDefineReader().ReadSheet(workbook.Worksheets[PmicConst.PortDefine]);

            if (new PreCheckVddLevels(workbook, PmicConst.VddLevels).CheckMain())
                VddLevelsSheet = new VddLevelsReader().ReadSheet(workbook.Worksheets[PmicConst.VddLevels]);

            if (new PreCheckIoLevels(workbook, PmicConst.IoLevels).CheckMain())
                IoLevelsSheet = new IoLevelsReader().ReadSheet(workbook.Worksheets[PmicConst.IoLevels]);

            if (new PreCheckDcTestContinuity(workbook, PmicConst.DcTestContinuity).CheckMain())
                DcTestContinuitySheet =
                    new DcTestContinuityReader().ReadSheet(workbook.Worksheets[PmicConst.DcTestContinuity]);

            if (new PreCheckPmicIds(workbook, PmicConst.PmicIds).CheckMain())
                PmicIdsSheet = new PmicIdsReader().ReadSheet(workbook.Worksheets[PmicConst.PmicIds]);

            var leakage = workbook.Worksheets.First(x =>
                x.Name.EndsWith("_" + PmicConst.Leakage, StringComparison.CurrentCultureIgnoreCase)).Name;
            if (string.IsNullOrEmpty(leakage))
                leakage = PmicConst.PmicLeakage;
            if (new PreCheckPmicLeakage(workbook, leakage).CheckMain())
                PmicLeakageSheet = new PmicLeakageReader().ReadSheet(workbook.Worksheets[leakage]);

            if (new PreCheckBscanChar(workbook, PmicConst.BscanChar).CheckMain())
                BscanCharSheet = new BscanCharReader().ReadSheet(workbook.Worksheets[PmicConst.BscanChar]);

            var dcTests = workbook.Worksheets.Where(x =>
                Regex.IsMatch(x.Name, "^" + PmicConst.DctTest, RegexOptions.IgnoreCase));
            foreach (var dcTest in dcTests)
            {
                if (dcTest.Name.Equals(PmicConst.DcTestContinuity, StringComparison.CurrentCultureIgnoreCase))
                    continue;

                if (new PreCheckHardip(workbook, dcTest.Name).CheckMain())
                    DcTestSheet.Add(new TestPlanReader().ReadSheet(workbook.Worksheets[dcTest.Name]));
            }

            if (new PreCheckAhbRegisterMap(workbook, PmicConst.AhbRegisterMap).CheckMain())
                AhbRegisterMapSheet =
                    new AhbRegisterMapReader().ReadSheet(workbook.Worksheets[PmicConst.AhbRegisterMap]);

            if (new PreCheckGenMainFlow(workbook, PmicConst.GenMainFlow).CheckMain())
                MainFlowSheet = new ReadFlowSheet().GetSheet(workbook.Worksheets[PmicConst.GenMainFlow]);

            if (new PreCheckOtpSetup(workbook, PmicConst.OtpSetup).CheckMain())
                OtpSetupSheet = new OtpSetupReader().ReadSheet(workbook.Worksheets[PmicConst.OtpSetup]);

            #endregion

            #region Post check

            var pins = IoPinMapSheet.PinList.Select(x => x.PinName).ToList();
            var pinGroupsByIoPinMap = IoPinMapSheet.GroupList.Select(x => x.PinName).ToList();
            var pinGroupsByIoPinGroup = IoPinGroupSheet.Rows.Select(x => x.PinGroupName).ToList();
            var totalPins = pins;
            totalPins.AddRange(pinGroupsByIoPinMap);
            totalPins.AddRange(pinGroupsByIoPinGroup);
            if (PortDefineSheet != null)
            {
                var pinGroupsByPortDefine = PortDefineSheet.Rows.Select(x => x.ProtocolPortName).ToList();
                totalPins.AddRange(pinGroupsByPortDefine);
            }

            totalPins = totalPins.Distinct().ToList();

            if (ChannelMapSheets != null)
                foreach (var channelMapSheet in ChannelMapSheets)
                    foreach (var row in channelMapSheet.ChannelMapRows)
                    {
                        var pinName = row.DeviceUnderTestPinName;
                        if (!string.IsNullOrEmpty(pinName) &&
                            !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                            ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                                channelMapSheet.SheetName, row.RowNum, ChannelMapSheet.DeviceUnderTestPinName,
                                errorMessage);
                        }
                    }

            if (IoPinGroupSheet != null)
                foreach (var row in IoPinGroupSheet.Rows)
                {
                    var pinName = row.PinNameContainedInPinGroup;
                    if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                            x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            IoPinGroupSheet.SheetName, row.RowNum, IoPinGroupSheet.PinNameContainedInPinGroupIndex,
                            errorMessage);
                    }

                    var pinGroupName = row.PinGroupName;
                    if (!string.IsNullOrEmpty(pinName) && pinGroupsByIoPinMap.Exists(x =>
                            x.Equals(pinGroupName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin group \"{0}\"is already exist in pin map sheet !!!",
                            pinGroupName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            IoPinGroupSheet.SheetName, row.RowNum, IoPinGroupSheet.PinGroupNameIndex, errorMessage);
                    }
                }

            if (PortDefineSheet != null)
                foreach (var row in PortDefineSheet.Rows)
                {
                    var pinName = row.Pin;
                    if (!string.IsNullOrEmpty(pinName) &&
                        !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            PortDefineSheet.SheetName, row.RowNum, PortDefineSheet.PinIndex, errorMessage);
                    }
                }

            if (VddLevelsSheet != null)
                foreach (var row in VddLevelsSheet.Rows)
                {
                    var pinName = row.WsBumpName;
                    if (!string.IsNullOrEmpty(pinName) &&
                        !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            VddLevelsSheet.SheetName, row.RowNum, VddLevelsSheet.WsBumpNameIndex, errorMessage);
                    }
                }

            if (IoLevelsSheet != null)
                foreach (var row in IoLevelsSheet.Rows)
                {
                    var pinName = row.PinName;
                    if (!string.IsNullOrEmpty(pinName) &&
                        !pins.Exists(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            IoLevelsSheet.SheetName, row.RowNum, IoLevelsSheet.PinNameIndex, errorMessage);
                    }
                }

            if (DcTestContinuitySheet != null)
                foreach (var row in DcTestContinuitySheet.Rows)
                {
                    var pinName = row.PinGroup;
                    if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                            x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            DcTestContinuitySheet.SheetName, row.RowNum, DcTestContinuitySheet.PinGroupIndex,
                            errorMessage);
                    }
                }

            if (PmicIdsSheet != null)
                foreach (var row in PmicIdsSheet.Rows)
                {
                    var pinName = row.MeasurePin;
                    if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
                            x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var errorMessage = string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
                        ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                            PmicIdsSheet.SheetName, row.RowNum, PmicIdsSheet.MeasurePinIndex, errorMessage);
                    }
                }

            //if (PmicLeakageSheet != null)
            //{
            //    //eg: MeasI pin = pin1,pin2               
            //    const string regMeasExpression =
            //        @"(?<MeasType>(Wi)*[(Meas)|(Src)]\S+)[\s]*(pin)?[\s]*=[\s]*(?<pin>(.*))";
            //    foreach (var row in PmicLeakageSheet.PatternRows)
            //    foreach (var patChildRows in row.PatChildRows)
            //        if (patChildRows is PatSubChildRow)
            //        {
            //            var childRow = (PatSubChildRow) patChildRows;
            //            foreach (var tpRow in childRow.TpRows)
            //            {
            //                var measStr = Regex.Match(tpRow.Meas, regMeasExpression, RegexOptions.IgnoreCase)
            //                    .Groups["pin"].ToString().Trim(',').Trim();
            //                foreach (var pinName in measStr.Split(','))
            //                    if (!string.IsNullOrEmpty(pinName) && !totalPins.Exists(x =>
            //                        x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
            //                    {
            //                        var errorMessage =
            //                            string.Format("This pin \"{0}\"is not exist in pin map sheet !!!", pinName);
            //                        EpplusErrorManager.AddError(EnumErrorType.FormatError, ErrorLevel.Error,
            //                            PmicLeakageSheet.SheetName, row.RowNum, PmicLeakageSheet.ForceIndex,
            //                            errorMessage);
            //                    }
            //            }
            //        }
            //}

            #endregion
        }
    }
}