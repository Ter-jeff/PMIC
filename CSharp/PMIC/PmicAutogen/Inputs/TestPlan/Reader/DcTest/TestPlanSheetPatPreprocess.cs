using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest
{
    public class TestPlanSheetPatPreprocess
    {
        private readonly TestPlanSheet _planSheet;

        public TestPlanSheetPatPreprocess(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
        }

        public void UpdateSheetPattern()
        {
            //MergeSplitPatterns(_planSheet);
            var instInProductList = _planSheet.PatternRows
                .Where(x => x.ForceCondition.IsShmooInProdInst | !x.ForceCondition.IsShmooInForce).ToList();
            GetPatternDupIndex(instInProductList);
            var instInCharList = _planSheet.PatternRows.Where(x => x.ForceCondition.IsShmooInCharInst).ToList();
            GetPatternDupIndex(instInCharList);
        }

        private void GetPatternDupIndex(List<PatternRow> planList)
        {
            var patItems = new Dictionary<string, int>();
            var duplicatePatterns = new List<string>();
            string mistSubstring;
            var substringItems = new Dictionary<string, int>();
            var duplicateSubstring = new List<string>();
            var rowList = new List<PatternRow>();
            foreach (var row in planList)
            {
                mistSubstring = row.MiscInfo.GetInstNameSubStr().ToLower();
                var blockName = row.MiscInfo.GetBlockName(row.SheetName);
                var subBlock = row.MiscInfo.GetSubBlockName(row.Pattern.GetLastPayload(), blockName).ToLower();
                var subBlock2 = row.MiscInfo.GetSubBlock2Name().ToLower();
                var patternName = subBlock + "_" + subBlock2 + row.Pattern.GetLastPayload().ToLower();
                //check whether pattern is repeated
                if (patItems.ContainsKey(patternName.ToLower()))
                {
                    if (mistSubstring != "")
                    {
                        if (substringItems.ContainsKey(mistSubstring))
                        {
                            substringItems[mistSubstring]++;
                            row.DupIndex = substringItems[mistSubstring];
                            duplicateSubstring.Add(mistSubstring);
                        }
                        else
                        {
                            row.DupIndex = 1;
                            substringItems.Add(mistSubstring, 1);
                        }
                    }
                    else
                    {
                        patItems[patternName.ToLower()]++;
                        row.DupIndex = patItems[patternName];

                        if (row.ForceCondition.IsShmooInCharInst)
                        {
                            if (!IsSameShmooRowAndSetDupIndex(row, rowList, patternName))
                                duplicatePatterns.Add(patternName);
                        }
                        else
                        {
                            if (row.ForceCondition.IsShmooInForce)
                            {
                                if (!IsSameShmooRowAndSetDupIndex(row, rowList, patternName))
                                    duplicatePatterns.Add(patternName);
                            }
                            else
                            {
                                duplicatePatterns.Add(patternName);
                            }
                        }
                    }
                }
                else
                {
                    row.DupIndex = 1;
                    if (mistSubstring != "")
                    {
                        if (substringItems.ContainsKey(mistSubstring))
                        {
                            substringItems[mistSubstring]++;
                            row.DupIndex = substringItems[mistSubstring];
                        }
                        else
                        {
                            substringItems.Add(mistSubstring, 1);
                        }
                    }
                    else
                    {
                        patItems.Add(patternName, 1);
                    }
                }

                rowList.Add(row);
            }

            // To set DupIndex from 1 to 0
            var precessRowList = new List<PatternRow>();
            foreach (var row in planList)
            {
                mistSubstring = row.MiscInfo.GetInstNameSubStr().ToLower();
                var blockName = row.MiscInfo.GetBlockName(row.SheetName);
                var subBlock = row.MiscInfo.GetSubBlockName(row.Pattern.GetLastPayload(), blockName).ToLower();
                var subBlock2 = row.MiscInfo.GetSubBlock2Name().ToLower();
                var patternName = subBlock + "_" + subBlock2 + row.Pattern.GetLastPayload().ToLower();
                if (row.ForceCondition.IsShmooInForce & (row.DupIndex != 1))
                    IsSameShmooRowAndSetDupIndex(row, precessRowList, patternName);
                if (duplicatePatterns.Contains(patternName.ToLower()))
                {
                    if (!duplicateSubstring.Contains(mistSubstring.ToLower()) && mistSubstring != "") row.DupIndex = 0;
                }
                else
                {
                    row.DupIndex = 0;
                }

                precessRowList.Add(row);
            }
        }

        private bool IsSameShmooRowAndSetDupIndex(PatternRow rootRow, List<PatternRow> rowList, string patternName)
        {
            var flag = false;
            foreach (var row in rowList)
            {
                var blockName = row.MiscInfo.GetBlockName(row.SheetName);
                var subBlock = row.MiscInfo.GetSubBlockName(row.Pattern.GetLastPayload(), blockName).ToLower();
                var ipName = row.MiscInfo.GetIpName().ToLower();
                var patternNameString = subBlock + "_" + ipName + row.Pattern.GetLastPayload().ToLower();
                //if (patternName == patternNameString)
                //if ((rootRow.IsShmooInProdInst & row.IsShmooInProdInst) | (rootRow.IsShmooInCharInst & row.IsShmooInCharInst))
                //if (HardipCharSetup.IsSameForceShmoo(rootRow.ForceCondition.ForceCondition,
                //    row.ForceCondition.ForceCondition))
                //{
                //    flag = true;
                //    rootRow.DupIndex = row.DupIndex;
                //    break;
                //}
            }

            return flag;
        }
    }
}