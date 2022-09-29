using CommonLib.Enum;
using CommonLib.WriteMessage;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Business
{
    public abstract class ProdCharPatSetConstructorBase
    {
        private const char ConstSplitChar = (char)1;
        protected string Block;
        protected string Domain;
        protected List<IProdCharSheetRow> InitList;
        protected List<string> PatternListUsage = new List<string>();
        protected List<IProdCharSheetRow> PayloadList;
        protected List<string> PayloadNamingFields = new List<string> { "full" };
        protected List<string> PerformanceModeList;

        #region Constructor

        protected ProdCharPatSetConstructorBase(IEnumerable<IProdCharSheetRow> inputRows)
        {
            var prodCharSheetRows = inputRows.ToList();
            InitList = GetAllInitRows(prodCharSheetRows.ToList());
            PayloadList = GetAllPayloadRows(prodCharSheetRows.ToList());
        }

        #endregion

        protected List<ProdCharRow> GetPatSetFromProdChar(List<IProdCharSheetRow> initList,
            List<IProdCharSheetRow> payloadList)
        {
            var missingInitNameList = new List<string>();
            var prodCharRows = new List<ProdCharRow>();
            var all = new List<IProdCharSheetRow>();
            all.AddRange(initList);
            all.AddRange(payloadList);
            var hasInit = initList.Count != 0;
            foreach (var dataRow in payloadList)
            {
                if (IsPatternIgnored(dataRow.PayloadValue))
                    continue;
                if (hasInit)
                {
                    var patSets = CartesianGenerate(dataRow, all, missingInitNameList);
                    foreach (var patSet in patSets)
                        patSet.RowNum = dataRow.RowNum;
                    prodCharRows.AddRange(patSets);
                }
                else
                {
                    var patSet = DirectGenerate(dataRow, all);
                    patSet.RowNum = dataRow.RowNum;
                    prodCharRows.Add(patSet);
                }
            }

            missingInitNameList = missingInitNameList.Distinct().ToList();
            foreach (var missingInit in missingInitNameList)
                Response.Report(string.Format("Can not Find Init :{0} ", missingInit), EnumMessageLevel.Error, 0);

            //Remove the same pattern
            var newPatSetList = DeleteSamePattern(prodCharRows);

            return newPatSetList;
        }

        private List<ProdCharRow> DeleteSamePattern(List<ProdCharRow> prodCharRows)
        {
            foreach (var patSet in prodCharRows)
            {
                var newInitList = new Dictionary<int, PatternWithMode>();
                var i = -1;
                foreach (var init in patSet.InitList.Values)
                {
                    i++;
                    if (string.IsNullOrEmpty(init.PatternName)) continue;
                    if (newInitList.Count == 0)
                    {
                        newInitList.Add(i, init);
                    }
                    else
                    {
                        if (init.PatternName != newInitList.Values.Last().PatternName)
                            newInitList.Add(i, init);
                    }
                }

                patSet.InitList = newInitList;

                var newPayloadList = new List<PatternWithMode>();
                foreach (var payLoad in patSet.PayloadList)
                {
                    if (string.IsNullOrEmpty(payLoad.PatternName)) continue;
                    if (newPayloadList.Count == 0)
                    {
                        newPayloadList.Add(payLoad);
                    }
                    else
                    {
                        if (payLoad != newPayloadList.Last())
                            newPayloadList.Add(payLoad);
                    }
                }

                patSet.PayloadList = newPayloadList;
            }

            return prodCharRows;
        }

        protected ProdCharRow DirectGenerate(IProdCharSheetRow row, List<IProdCharSheetRow> rows)
        {
            var instanceTemp = new ProdCharRow(row);
            var i = -1;
            foreach (var initName in row.GetInitList())
            {
                i++;
                //init name is null
                if (initName != "" && initName != "NA" && initName != "N/A")
                    instanceTemp.InitList.Add(i,
                        new PatternWithMode { PatternName = initName, Mode = GetMode(initName) });
                else
                    instanceTemp.InitList.Add(i, new PatternWithMode { PatternName = "", Mode = "" });
            }

            foreach (var payload in GetPayloadList(row, rows))
                instanceTemp.PayloadList.Add(new PatternWithMode { PatternName = payload, Mode = GetMode(payload) });
            return instanceTemp;
        }

        private List<string> GetPayloadList(IProdCharSheetRow row, List<IProdCharSheetRow> rows)
        {
            var currentPayloadList = new List<string>();
            foreach (var currentPayload in row.GetPayloadList())
                if (currentPayload != "" && currentPayload != "NA" && currentPayload != "N/A")
                {
                    var payloads = FindAllInMode(currentPayload, rows);
                    if (payloads.Count != 0) row.PayLoads += string.Join(",", payloads) + ";";

                    if (payloads.Count == 0)
                    {
                        payloads = FindAllInItem(currentPayload, rows);
                        if (payloads.Count != 0) row.PayLoads += string.Join(",", payloads) + ";";
                    }

                    if (payloads.Count == 0)
                    {
                        payloads.Add(currentPayload);
                        row.PayLoads += currentPayload + ",";
                    }

                    currentPayloadList.AddRange(payloads);
                }
                else
                {
                    currentPayloadList.Add("");
                }

            return currentPayloadList;
        }

        protected List<ProdCharRow> CartesianGenerate(IProdCharSheetRow row, List<IProdCharSheetRow> rowList,
            List<string> missingInitList)
        {
            var patSetList = new List<ProdCharRow>();
            var payloadListTemp = new List<string>();
            payloadListTemp.Add(row.PayloadValue);

            //Merge all the init to a list connect by char (1), and store them in payloadListTemp
            foreach (var currentInit in row.GetInitList())
            {
                var currentPayloadList = new List<string>();
                if (currentInit != "" && currentInit != "NA" && currentInit != "N/A")
                {
                    var inits = FindInitInMode(currentInit, rowList);
                    if (inits.Count != 0) row.Inits += string.Join(",", inits) + ";";

                    if (inits.Count == 0)
                    {
                        inits = FindInitInItem(currentInit, rowList);
                        if (inits.Count != 0) row.Inits += string.Join(",", inits) + ";";
                    }

                    if (inits.Count == 0)
                    {
                        inits.Add(currentInit);
                        row.Inits += currentInit + ",";
                    }

                    currentPayloadList.AddRange(inits);
                }
                else
                {
                    currentPayloadList.Add("");
                }

                BinaryCartesianMerge(payloadListTemp, currentPayloadList);
            }

            foreach (var listPayload in payloadListTemp)
            {
                var prodCharRowTemp = new ProdCharRow(row);
                //Split all the init and payload in the payloadListTemp
                var splitPayload = listPayload.Split(ConstSplitChar);
                if (splitPayload.Any())
                {
                    for (var i = 1; i < splitPayload.Length; i++)
                        prodCharRowTemp.InitList.Add(i - 1,
                            new PatternWithMode { PatternName = splitPayload[i], Mode = GetMode(splitPayload[i]) });
                    prodCharRowTemp.InitAliasList.AddRange(row.GetInitList());

                    foreach (var payload in GetPayloadList(row, rowList))
                        prodCharRowTemp.PayloadList.Add(new PatternWithMode
                        { PatternName = payload, Mode = GetMode(payload) });
                    prodCharRowTemp.PayloadAliasList.AddRange(row.GetPayloadList());
                }

                patSetList.Add(prodCharRowTemp);
            }

            return patSetList;
        }

        protected List<string> FindInitInMode(string init, List<IProdCharSheetRow> rows)
        {
            return rows.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase))
                .Select(p => p.PayloadValue).ToList();
        }

        protected List<string> FindInitInItem(string init, List<IProdCharSheetRow> rows)
        {
            return rows.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase))
                .Select(p => p.PayloadValue).ToList();
        }

        protected List<string> FindAllInMode(string init, List<IProdCharSheetRow> rows)
        {
            var allList = rows.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase))
                .SelectMany(p => p.GetInitList()).ToList();
            allList.AddRange(rows.FindAll(p => p.Mode.Equals(init, StringComparison.OrdinalIgnoreCase))
                .Select(p => p.PayloadValue).ToList());
            return allList;
        }

        protected List<string> FindAllInItem(string init, List<IProdCharSheetRow> rows)
        {
            var allList = rows.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase))
                .SelectMany(p => p.GetInitList()).ToList();
            allList.AddRange(rows.FindAll(p => p.Item.Equals(init, StringComparison.OrdinalIgnoreCase))
                .Select(p => p.PayloadValue).ToList());
            return allList;
        }

        protected List<IProdCharSheetRow> GetAllInitRows(List<IProdCharSheetRow> prodCharSheetRows)
        {
            return prodCharSheetRows.Where(p => IsInitPattern(p.PayloadValue)).ToList();
        }

        protected List<IProdCharSheetRow> GetAllPayloadRows(List<IProdCharSheetRow> prodCharSheetRows)
        {
            return prodCharSheetRows.Where(p => !IsInitPattern(p.PayloadValue)).ToList();
        }

        protected virtual bool IsPatternIgnored(string pattern)
        {
            return false;
        }

        protected List<IProdCharSheetRow> FilterProChar(List<IProdCharSheetRow> inputList)
        {
            return inputList.Where(a => a.Usage != "0").ToList();
        }

        protected bool IsInitPattern(string pattern)
        {
            var numbersStrings = pattern.Split('_');
            if (numbersStrings.Length < 4)
                return false;


            var lStrMatchPattern = "IN.*";
            if (Regex.IsMatch(numbersStrings[3], lStrMatchPattern, RegexOptions.IgnoreCase)) return true;

            if (numbersStrings.Length < 7)
                return false;

            // for gfx SPC0 1
            if (Regex.IsMatch(numbersStrings[2], "L", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(numbersStrings[6], "SPC", RegexOptions.IgnoreCase))
                return true;

            return false;
        }

        protected void BinaryCartesianMerge(List<string> payloadList1, List<string> payloadList2)
        {
            var resultList = new List<string>();
            const char splitChar = ConstSplitChar; //Connect all the string by (char)1
            foreach (var name1 in payloadList1)
                foreach (var name2 in payloadList2)
                    resultList.Add(name1 + splitChar + name2);
            payloadList1.Clear();
            payloadList1.AddRange(resultList);
        }

        protected string GetSubName(string name, string rule)
        {
            var resultList = new List<string>();
            var words = name.Split('_');
            var numbersStrings = rule.Split(',');
            foreach (var numbers in numbersStrings)
            {
                if (rule.ToLower() == "full") return name;

                if (rule == "")
                {
                    //no operation
                }
                else
                {
                    var getNumber = int.Parse(numbers);
                    if (words.Length > getNumber && getNumber >= 0) resultList.Add(words[getNumber]);
                }
            }

            var resultName = string.Join("_", resultList);

            return resultName;
        }

        protected string GetPerformanceMode(ProdCharRow prodCharRow, List<string> modeList)
        {
            var performanceMode = "";
            foreach (var init in prodCharRow.InitList.Values)
            {
                var mode = GetMode(init.PatternName);
                performanceMode = string.IsNullOrEmpty(mode) ? performanceMode : mode;
            }

            return performanceMode;
        }

        private string GetMode(string pattern)
        {
            var performanceMode = "";
            var subStrings = pattern.Split('_');
            {
                if (subStrings.Length > 9 && Regex.IsMatch(subStrings[9], @"^M[a-zA-Z]+\d+$"))
                    if (!Regex.IsMatch(subStrings[9], "999|010"))
                        performanceMode = subStrings[9];
            }
            return performanceMode;
        }

        protected virtual string GetPayLoadName(ProdCharRow prodCharRow)
        {
            if (PayloadNamingFields.Any())
                return GetSubName(prodCharRow.PayLoadName, PayloadNamingFields.First()).Replace(" ", "_");
            return GetSubName(prodCharRow.PayLoadName, "full");
        }

        protected virtual string GetPrefix(string payloadType = "")
        {
            return Domain + Block;
        }

        protected List<ProdCharRow> FilterPatSetWithUsedPatterns(List<ProdCharRow> prodCharRows)
        {
            var result = new List<ProdCharRow>();
            foreach (var oriPatSet in prodCharRows)
            {
                var isUsed = false;
                foreach (var init in oriPatSet.InitList.Values)
                    if (PatternListUsage.Exists(p => p.Equals(init.PatternName, StringComparison.OrdinalIgnoreCase)))
                    {
                        result.Add(oriPatSet);
                        isUsed = true;
                        break;
                    }

                if (isUsed) continue;
                foreach (var payload in oriPatSet.PayloadList)
                    if (PatternListUsage.Exists(p => p.Equals(payload.PatternName, StringComparison.OrdinalIgnoreCase)))
                    {
                        result.Add(oriPatSet);
                        break;
                    }
            }

            result.ForEach(p => p.SkipCheckRule = true);
            return result;
        }
    }
}