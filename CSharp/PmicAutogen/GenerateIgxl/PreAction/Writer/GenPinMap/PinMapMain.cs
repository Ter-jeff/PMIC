using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap.PortMapModify;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap
{
    public class PinMapMain
    {
        private readonly IoPinGroupSheet _ioPinGroupSheet;
        private readonly PinMapSheet _ioPinMapSheet;
        private readonly PortDefineSheet _portDefineSheet;

        public PinMapMain(PinMapSheet ioPinMapSheet, IoPinGroupSheet ioPinGroupSheet, PortDefineSheet portDefineSheet)
        {
            _ioPinMapSheet = ioPinMapSheet ?? new PinMapSheet(PmicConst.PinMap);
            _ioPinGroupSheet = ioPinGroupSheet;
            _portDefineSheet = portDefineSheet;
        }

        public PinMapSheet GetPinMapSheet()
        {
            var pinMap = CreatePinMap(_ioPinMapSheet);
            foreach (var pin in pinMap.PinList)
                if (StaticTestPlan.ChannelMapSheets != null && StaticTestPlan.ChannelMapSheets.Any())
                    if (StaticTestPlan.ChannelMapSheets.SelectMany(x => x.ChannelMapRows).ToList().Exists(y =>
                        y.DeviceUnderTestPinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var channelMapRow = StaticTestPlan.ChannelMapSheets.SelectMany(x => x.ChannelMapRows).ToList()
                            .Find(y => y.DeviceUnderTestPinName.Equals(pin.PinName,
                                StringComparison.OrdinalIgnoreCase));
                        pin.ChannelType = channelMapRow.Type;
                        pin.InstrumentType = channelMapRow.InstrumentType;
                    }

            return pinMap;
        }

        public PinMapSheet CreatePinMap(PinMapSheet pinmapSheet, bool isMsg = true)
        {
            if (pinmapSheet == null) return null;

            if (_ioPinGroupSheet != null)
            {
                var groupFromPinGroupSheet = ReadGroupSheet(pinmapSheet, _ioPinGroupSheet, _portDefineSheet);
                MergePinMapSheet(pinmapSheet, groupFromPinGroupSheet, isMsg);
            }
            return pinmapSheet;
        }

        private List<PinGroup> ReadGroupSheet(PinMapSheet pinMap, IoPinGroupSheet pinGroupSheet, PortDefineSheet portDefineSheet)
        {
            if (portDefineSheet != null)
            {
                var groups = ParsePinGroupSheet(pinMap, portDefineSheet);
                foreach (var group in groups)
                {
                    if (!pinMap.GroupList.Exists(x => x.PinName.Equals(group.PinName, StringComparison.OrdinalIgnoreCase)))
                    {
                        pinMap.GroupList.Add(group);
                    }
                }
            }
            var pinGroups = ParsePinGroupSheet(pinMap, pinGroupSheet);
            pinGroups = ModifyPinGroups(pinMap, pinGroups);
            return pinGroups;
        }

        private List<PinGroup> ParsePinGroupSheet(PinMapSheet pinMap, PortDefineSheet portDefineSheet)
        {
            var pinGroups = new List<PinGroup>();
            var groups = portDefineSheet.Rows.GroupBy(x => x.ProtocolPortName);
            foreach (var group in groups)
            {
                var pinGroup = new PinGroup(group.Key);
                foreach (var pin in group)
                {
                    var onePin = pinMap.GetPin(pin.Pin);
                    onePin.Comment = "PortMap";
                    pinGroup.AddPin(onePin);
                }
                pinGroups.Add(pinGroup);
            }
            return pinGroups;
        }

        private List<PinGroup> ParsePinGroupSheet(PinMapSheet pinMap, IoPinGroupSheet pinGroupSheet)
        {
            var pinGroups = new List<PinGroup>();
            var startCol = -1;

            for (var i = 0; i < pinGroupSheet.Rows.Count;)
            {
                var row = pinGroupSheet.Rows[i];
                var groupName = row.PinGroupName;
                if (string.IsNullOrEmpty(groupName))
                    continue;

                var pins = new List<Pin>();
                while (i < pinGroupSheet.Rows.Count && (pinGroupSheet.Rows[i].PinGroupName.Equals("") ||
                                                        pinGroupSheet.Rows[i].PinGroupName.Equals(groupName,
                                                            StringComparison.OrdinalIgnoreCase)))
                {
                    var pinName = pinGroupSheet.Rows[i].PinNameContainedInPinGroup;
                    if (pinName.Contains("+") || pinName.Contains("-"))
                    {
                        pins.AddRange(PinGroupOperation(pinMap, pinName));
                    }
                    else if (pinName.Contains("*"))
                    {
                        pins.AddRange(SearchPins(pinMap, pinName));
                    }
                    else
                    {
                        pinName = pinName.Replace("[", "").Replace("]", "");
                        if (!pinName.Equals(""))
                        {
                            if (pinMap.IsGroupExist(pinName))
                            {
                                pins.Add(new Pin(pinName, pinMap.GetGroup(pinName).PinType));
                            }
                            else if (pinGroups.Exists(p => p.PinName.ToLower().Equals(pinName.ToLower())))
                            {
                                PinGroup grp = pinGroups.Find(p => p.PinName.ToLower().Equals(pinName.ToLower()));
                                if (grp != null)
                                {
                                    pins.Add(new Pin(pinName, grp.PinType));
                                }
                            }
                            else
                            {
                                if (pinMap.IsPinExist(pinName))
                                {
                                    pins.Add(pinMap.GetPin(pinName));
                                }
                                else
                                {
                                    var newPin = new Pin(pinName, "");
                                    pins.Add(newPin);
                                    var errorMessage = string.Format("The pin {0} in IO_PinGroup can not be found !!!",
                                        pinName);
                                    EpplusErrorManager.AddError(BasicErrorType.Business, ErrorLevel.Error,
                                        pinGroupSheet.SheetName, i, startCol + 1, errorMessage);
                                }
                            }
                        }
                    }

                    i++;
                }

                var pinTypes = pins.Select(x => x.PinType).Distinct().ToList();
                var pinType = pinTypes.Count == 1 ? pinTypes.First() : "";
                if (pinTypes.Count > 1)
                {
                    var errorMessage = string.Format("The pin group {0} has more than two pin types !!!", groupName);
                    EpplusErrorManager.AddError(BasicErrorType.Business, ErrorLevel.Error, pinGroupSheet.SheetName, i,
                        startCol, errorMessage);
                }

                var group = new PinGroup(groupName);
                if (pins.Any(x => x.PinName.Equals(groupName)))
                {
                    var errorMessage = string.Format("The pin group name {0} and pin name can not be the same !!!",
                        groupName);
                    EpplusErrorManager.AddError(BasicErrorType.Business, ErrorLevel.Error, pinGroupSheet.SheetName, i,
                        startCol, errorMessage);
                }

                group.AddPins(pins, pinGroupSheet.SheetName);
                pinGroups.Add(group);
            }

            return pinGroups;
        }

        private List<Pin> PinGroupOperation(PinMapSheet pinMap, string operationCommand)
        {
            var lStrPattern = "[+|-]";
            var keyWordList = Regex.Split(operationCommand, lStrPattern).ToList();
            var opList = Regex.Matches(operationCommand, lStrPattern).Cast<Match>().Select(a => a.Value).ToList();

            var pins = SearchPins(pinMap, keyWordList.First());
            for (var i = 0; i < keyWordList.Count - 1; i++)
            {
                var pinListTemp = SearchPins(pinMap, keyWordList[i + 1]);
                if (opList[i].Equals("+"))
                {
                    //Union
                    foreach (var pin in pinListTemp)
                        if (!pins.Exists(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                            pins.Add(pin);
                }
                else if (opList[i].Equals("-"))
                {
                    //Intersection
                    foreach (var pin in pinListTemp)
                        if (pins.Exists(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                            pins.RemoveAll(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    throw new Exception(string.Format("Unknown operation {0}", opList[i]));
                }
            }

            return pins;
        }

        private List<Pin> SearchPins(PinMapSheet pinMap, string keyWord)
        {
            var pins = new List<Pin>();
            if (keyWord.Contains("*"))
            {
                //use greed match method
                var lStrPattern = keyWord.Replace("*", ".*");
                pins.AddRange(
                    pinMap.PinList.FindAll(p => Regex.IsMatch(p.PinName, lStrPattern, RegexOptions.IgnoreCase)));
            }
            else if (pinMap.IsGroupExist(keyWord))
            {
                //if it is a group
                var group = pinMap.GetGroup(keyWord);
                foreach (var pin in group.PinList) pins.AddRange(SearchPins(pinMap, pin.PinName));
            }
            else if (pinMap.IsPinExist(keyWord))
            {
                //if it is a pin
                pins.Add(pinMap.GetPin(keyWord));
            }
            else
            {
                throw new Exception(string.Format("Unknown command {0}", keyWord));
            }

            return pins;
        }

        private void MergePinMapSheet(PinMapSheet pinMapSheet, List<PinGroup> pinGroupList, bool isMsg = true)
        {
            foreach (var pinGroup in pinGroupList)
            {
                //Ignore empty pinGroup
                if (!pinGroup.PinList.Any())
                    continue;

                if (pinMapSheet.IsGroupExist(pinGroup.PinName))
                {
                    var mapGroup = pinMapSheet.GroupList.FirstOrDefault(a =>
                        a.PinName.Equals(pinGroup.PinName, StringComparison.OrdinalIgnoreCase));
                    if (mapGroup != null && (!mapGroup.PinList.All(a =>
                                                 pinGroup.PinList.Select(p => p.PinName).ToList().Contains(a.PinName)) ||
                                             !pinGroup.PinList.All(a =>
                                                 mapGroup.PinList.Select(p => p.PinName).ToList().Contains(a.PinName))))
                    {
                        if (isMsg)
                        {
                            var outString = "PinGroup : " + pinGroup.PinName + " not match between " +
                                            PmicConst.IoPinGroup +
                                            " sheet and generated from Default or existed in PinMap!"; //IO_PinGroup
                            EpplusErrorManager.AddError(PreActionErrorType.PinGroupNotMatch.ToString(), ErrorLevel.Error,
                                "", 1, outString);
                        }
                        //merge the pin group if there is some pins not grouped in the pin map
                        if (!pinGroup.PinList.All(a => mapGroup.PinList.Contains(a)))
                        {
                            var pins = pinGroup.PinList.Select(p => p).Where(p =>
                                    !mapGroup.PinList.Select(x => x.PinName.ToUpper()).Contains(p.PinName.ToUpper()))
                                .ToList();
                            mapGroup.PinList.AddRange(pins);
                        }
                    }

                    continue;
                }
                pinMapSheet.AddRow(pinGroup);
            }
        }

        private List<PinGroup> ModifyPinGroups(PinMapSheet pinMap, List<PinGroup> pinGroups)
        {
            var pinGroupNew = new List<PinGroup>();
            foreach (var group in pinGroups)
            {
                if (!group.PinList.Any()) continue;
                var firstPinName = group.PinList.First();
                var pin = pinMap.PinList.FirstOrDefault(a =>
                    a.PinName.Equals(firstPinName.PinName, StringComparison.OrdinalIgnoreCase));
                if (pin != null && pin.PinType != "")
                    group.PinType = pin.PinType;
                if (Regex.IsMatch(group.PinName, @"_DIFF\d?$|_DIFF_", RegexOptions.IgnoreCase) &&
                    !group.PinName.EndsWith("_Port", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var item in group.PinList)
                        item.PinType = "Differential";
                    group.PinType = "Differential";
                }

                pinGroupNew.Add(group);
            }


            return pinGroupNew;
        }
    }
}