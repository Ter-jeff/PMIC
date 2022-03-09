using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AutomationCommon.DataStructure;
using IgxlData.Others.PatternListCsvFile;
using PmicAutogen.InputPackages;

namespace PmicAutogen.Inputs.PatternList
{
    public class PatternListReader
    {
        public List<PatternListCsvRow> ReadPatList(string fileName)
        {
            var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
            var fileReader = new StreamReader(fs);
            var headerOrder = new Dictionary<string, int>();
            var patternListCsvRows = new List<PatternListCsvRow>();
            var line = fileReader.ReadLine();
            var rowIndex = 0;
            try
            {
                var index = 0;
                if (line != null)
                {
                    foreach (var str in line.Split(','))
                    {
                        headerOrder.Add(str.Replace("\"", ""), index);
                        index++;
                    }

                    while ((line = fileReader.ReadLine()) != null)
                    {
                        rowIndex++;
                        line = line.Replace("\"", "");
                        if (line.Trim() == "")
                        {
                            Response.Report(
                                "Blank Row " + rowIndex + " In Pattern List csv " + fileName + " is skipped.",
                                MessageLevel.Warning, 100);
                            continue;
                        }

                        var patternListCsvRow = new PatternListCsvRow();
                        var lineData = line.Split(',').ToList();
                        var lineCount = lineData.Count;
                        if (lineCount <= index)
                            for (var i = 0; i < index - lineCount; i++)
                                lineData.Add("");

                        var isBlankRow = true;
                        for (var i = 0; i < index; i++)
                            if (lineData[i].Trim() != "")
                            {
                                isBlankRow = false;
                                break;
                            }

                        if (isBlankRow)
                        {
                            Response.Report(
                                "Blank Row " + rowIndex + " In Pattern List csv " + fileName + " is skipped.",
                                MessageLevel.Warning, 100);
                            continue;
                        }

                        var strArray = lineData.ToArray();
                        if (headerOrder.ContainsKey("Pattern"))
                            patternListCsvRow.PatternName = strArray[headerOrder["Pattern"]].ToLower();
                        if (patternListCsvRow.PatternName.Trim() == "")
                        {
                            Response.Report(
                                "Because Pattern is Blank, Row " + rowIndex + " Content " + line +
                                " In Pattern List csv " +
                                fileName + " is skipped.", MessageLevel.Warning, 100);
                            continue;
                        }

                        if (headerOrder.ContainsKey("Latest Version"))
                            patternListCsvRow.LatestVersion = strArray[headerOrder["Latest Version"]].ToLower();

                        if (headerOrder.ContainsKey("USE/No Use"))
                            patternListCsvRow.Use = strArray[headerOrder["USE/No Use"]].ToLower();

                        if (headerOrder.ContainsKey("Org"))
                            patternListCsvRow.Org = strArray[headerOrder["Org"]].ToLower();

                        if (headerOrder.ContainsKey("Type Spec"))
                            patternListCsvRow.TypeSpec = strArray[headerOrder["Type Spec"]].ToLower();
                        {
                            if (headerOrder.ContainsKey("Timeset Latest"))
                                patternListCsvRow.TimeSetVersion =
                                    Path.GetFileNameWithoutExtension(strArray[headerOrder["Timeset Latest"]]);
                            if (headerOrder.ContainsKey("Timeset Version"))
                                patternListCsvRow.TimeSetVersion =
                                    Path.GetFileNameWithoutExtension(strArray[headerOrder["Timeset Version"]]);
                        }

                        if (headerOrder.ContainsKey("File Versions"))
                            patternListCsvRow.FileVersion = strArray[headerOrder["File Versions"]].ToLower();

                        if (headerOrder.ContainsKey("OpCode"))
                            patternListCsvRow.OpCode = strArray[headerOrder["OpCode"]].ToLower();

                        if (headerOrder.ContainsKey("ScanMode"))
                            patternListCsvRow.ScanMode = strArray[headerOrder["ScanMode"]].ToLower();

                        if (headerOrder.ContainsKey("Halt"))
                            patternListCsvRow.Halt = strArray[headerOrder["Halt"]].ToLower();

                        if (headerOrder.ContainsKey("Original Timing Mode"))
                            patternListCsvRow.OriginalTimingMode =
                                strArray[headerOrder["Original Timing Mode"]].ToLower();

                        if (headerOrder.ContainsKey("Check"))
                            patternListCsvRow.Check = strArray[headerOrder["Check"]].ToLower();

                        if (headerOrder.ContainsKey("CheckComment"))
                            patternListCsvRow.CheckComment = strArray[headerOrder["CheckComment"]].ToLower();

                        if (headerOrder.ContainsKey("T/P Category"))
                            patternListCsvRow.TpCategory = strArray[headerOrder["T/P Category"]].ToLower();
                        patternListCsvRows.Add(patternListCsvRow);
                    }
                }

                return patternListCsvRows;
            }
            catch (Exception e)
            {
                throw new Exception("Reading pattern list failed, may be caused by wrong format of pattern list. " +
                                    e.Message);
            }
            finally
            {
                fileReader.Close();
                fs.Close();
            }
        }
    }
}