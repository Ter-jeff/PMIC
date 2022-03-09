//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Bug#            Notes
//
// 2019 March 28 Oliver Ou                    Initial creation
//
//------------------------------------------------------------------------------ 

using FWFrame;
using FWFrame.InputLoader;
using nWireDefinition.Enums;
using nWireDefinition.InputModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace nWireDefinition.InputLoader
{
    public class AsciiPatternFileInputLoader : IInputLoader
    {
        public void Load(Context ctx)
        {
            GUIInfo guiInfo = ctx.Get<GUIInfo>("guiInfo");

            // Show processing status
            Action<int, string> _reportStatus = guiInfo.GetParameter<Action<int, string>>("reportStatus");
            _reportStatus((int)ProcessPhaseEnum.ANALYSE_COMPILED_ATP_FILES, "Analysing pattern files with atp format");

            List<Protocol> protocals = ctx.Get<List<Protocol>>("protocals");

            string timeSetName = guiInfo.GetParameter<string>("timeSetName");
            Dictionary<string, Tuple<string, string>> pinMappingInfo = guiInfo.GetParameter<Dictionary<string, Tuple<string, string>>>("pinMappingInfo");
            List<string> frameNames = guiInfo.GetParameter<List<string>>("frameNames");
            Dictionary<string, List<List<string>>> fieldInfoForAllFrames = guiInfo.GetParameter<Dictionary<string, List<List<string>>>>("fieldInfoForAllFrames");

            List<string> atpFiles = ctx.Get<List<string>>("atpFiles");
            Dictionary<string, List<Cycle>> dataInfoDic = new Dictionary<string, List<Cycle>>();
            Dictionary<string, List<Field>> fieldInfoDic = new Dictionary<string, List<Field>>();

            for (int frameIndex = 0; frameIndex < atpFiles.Count; frameIndex++)
            {
                string frameName = frameNames[frameIndex];

                // Get Pin Names
                List<string> allLines = new List<string>(File.ReadAllLines(atpFiles[frameIndex]));

                int contentStartIndex = allLines.FindIndex(x => x.Trim().Equals("{"));
                int contentEndIndex = allLines.FindIndex(contentStartIndex, x => x.Trim().Equals("}"));

                int maxNameHeight = 0;
                for (int i = contentStartIndex + 1; i < contentEndIndex; i++)
                {
                    if (!allLines[i].StartsWith("//"))
                    {
                        break;
                    }
                    maxNameHeight++;
                }

                List<string> listPinNameRow = new List<string>();
                for (int i = contentStartIndex + 1; i < contentStartIndex + 1 + maxNameHeight; i++)
                {
                    string tmp = allLines[i].TrimStart('/');
                    if (!string.IsNullOrEmpty(tmp))
                    {
                        listPinNameRow.Add(tmp);
                    }
                }

                List<string> listPinName = new List<string>();
                int pinCount = listPinNameRow[0].Length;
                for (int i = 0; i < pinCount; i++)
                {
                    StringBuilder pinName = new StringBuilder();
                    for (int j = 0; j < listPinNameRow.Count; j++)
                    {
                        pinName.Append(listPinNameRow[j][i]);
                    }
                    listPinName.Add(pinName.ToString().Trim());
                }
                listPinName.RemoveAll(x => string.IsNullOrWhiteSpace(x));

                // Getting corresponding indice for user input Pins
                List<Tuple<int, string>> listPinIndex = new List<Tuple<int, string>>();
                foreach (var info in pinMappingInfo)
                {
                    int index = listPinName.FindIndex(x => x.Equals(info.Value.Item1, StringComparison.CurrentCultureIgnoreCase));
                    if (index == -1)
                    {
                        throw new FWFrameException("Can not find Pin Name [" + info.Value + "] for " + info.Key);
                    }
                    listPinIndex.Add(new Tuple<int, string>(index, info.Value.Item2));
                }

                // Get corresponding data for user input Pins
                string validVectorPattern = string.Format(@"^(repeat\s+\d+)?(.*)\s+>\s+tsetJTAG\s+(.+);", timeSetName);
                List<Cycle> listCycle = new List<Cycle>();
                int vectorIndex = 0;
                int cycleIndex = 0;
                for (int i = contentStartIndex + 1 + maxNameHeight; i < contentEndIndex; i++)
                {
                    // Skip comment line
                    if (allLines[i].Trim().StartsWith("//"))
                    {
                        continue;
                    }

                    Match match = Regex.Match(allLines[i], validVectorPattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        GroupCollection groups = match.Groups;
                        string repeatPart = groups[1].Value;
                        string dataPart = Regex.Replace(groups[3].Value, @"\s+", "");

                        Cycle cycle = new Cycle();
                        cycle.VectorIndex = vectorIndex;
                        cycle.CycleIndex = cycleIndex;
                        if (!string.IsNullOrWhiteSpace(repeatPart))
                        {
                            cycle.Repeat = int.Parse(Regex.Replace(repeatPart, @"repeat\s+", "", RegexOptions.IgnoreCase)) - 1;
                            cycleIndex += cycle.Repeat + 1;
                        }
                        else
                        {
                            cycleIndex += 1;
                        }
                        foreach (var index in listPinIndex)
                        {
                            // For all pins, if data is "V", then output "E"
                            // For clock pin, if data is "1", then output "-"
                            string patternData = dataPart[index.Item1].ToString();
                            if (patternData.Equals("V"))
                            {
                                patternData = "E";
                            }

                            if (index.Item2.Equals("Clock", StringComparison.CurrentCultureIgnoreCase))
                            {
                                if (patternData.Equals("1"))
                                {
                                    patternData = "-";
                                }
                            }

                            cycle.Data.Add(patternData);
                        }
                        listCycle.Add(cycle);

                        vectorIndex++;
                    }
                }
                dataInfoDic.Add(frameName, listCycle);

                // Add Field Info
                List<Field> fields = new List<Field>();
                List<List<string>> fieldInfoForSingleFrame = fieldInfoForAllFrames[frameName];
                foreach (List<string> fieldInfo in fieldInfoForSingleFrame)
                {
                    Field field = new Field();
                    field.FieldName = fieldInfo[0];

                    int startVector = int.Parse(fieldInfo[3]);
                    int stopVector = int.Parse(fieldInfo[4]);

                    int maxVector = listCycle.Count - 1;
                    if (stopVector > maxVector && startVector > maxVector)
                    {
                        throw new FWFrameException(string.Format("Start Vector or Stop Vector for Frame [{0}] is out of range, max vector is {1}", frameName, maxVector));
                    }

                    Cycle startCycle = listCycle.Find(x => x.VectorIndex == startVector);
                    Cycle stopCycle = listCycle.Find(x => x.VectorIndex == stopVector);

                    string portName = string.Empty;
                    foreach (string key in pinMappingInfo.Keys)
                    {
                        if (pinMappingInfo[key].Item1.Equals(fieldInfo[1], StringComparison.CurrentCultureIgnoreCase))
                        {
                            portName = key;
                        }
                    }

                    for (int i = startCycle.CycleIndex; i <= stopCycle.CycleIndex + stopCycle.Repeat; i++)
                    {
                        field.PortNames.Add(portName);
                        field.CycleIndice.Add(i.ToString());
                    }

                    fields.Add(field);
                }
                fieldInfoDic.Add(frameName, fields);
            }

            ctx.Add("dataInfoDic", dataInfoDic);
            ctx.Add("fieldInfoDic", fieldInfoDic);
        }
    }
}
