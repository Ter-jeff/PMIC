using System.Collections.Generic;
using System.Text.RegularExpressions;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    public class TestPlanSheet
    {
        public TestPlanSheet()
        {
            PatternRows = new List<PatternRow>();
            PatternItems = new List<HardIpPattern>();
            MultipleInit = false;
            PlanHeaderIdx = new Dictionary<string, int>();
        }

        public string SheetName { get; set; }
        public List<PatternRow> PatternRows { get; set; }
        public List<HardIpPattern> PatternItems { get; set; }

        public string ForceStr { get; set; }
        public int ForceIndex { get; set; }
        public int MeasIndex { get; set; }
        public bool MultipleInit { get; set; }
        public Dictionary<string, int> PlanHeaderIdx { get; set; }

        public void DividePatternRow()
        {
            for (var i = 0; i < PatternRows.Count; i++)
            {
                var patRow = PatternRows[i];

                //merge DDR pattern
                if (Regex.IsMatch(patRow.MiscInfo, HardIpConstData.DdrLabel, RegexOptions.IgnoreCase))
                    if (i != PatternRows.Count - 1)
                    {
                        patRow.DdrExtraPat = PatternRows[i + 1];
                        PatternRows.RemoveAt(i + 1);
                    }

                //Divide pattern row by MultiplePayload
                var arr = patRow.Pattern.RealPatternName.Split('|');
                for (var a = 1; a < arr.Length; a++)
                {
                    var secondPatRow = patRow.DeepClone();
                    secondPatRow.Pattern = new PatternClass(arr[a]);
                    i++;
                    PatternRows.Insert(i, secondPatRow);
                }

                patRow.Pattern = new PatternClass(arr[0]);

                // Set for CZ2only
                var cz2Only = CommonGenerator.IsCz2Only(patRow.MiscInfo);
                if (cz2Only)
                {
                    patRow.ForceCondition.IsShmooInProdInst = false;
                    patRow.ForceCondition.IsShmooInProdFlow = false;
                    patRow.ForceCondition.IsShmooInCharInst = true;
                    patRow.ForceCondition.IsShmooInCharFlow = true;
                }

                // Divide pattern row by shmoo
                patRow.ForceCondition.IsShmooInForce = IsShmooInForce(patRow);
                if (patRow.ForceCondition.IsShmooInForce)
                {
                    var sameFlag = true;
                    if (patRow.ForceConditionChar != null)
                        sameFlag = HardipCharSetup.IsSameForceShmoo(patRow.ForceCondition.ForceCondition,
                            patRow.ForceConditionChar);

                    if (patRow.ForceConditionChar == null)
                    {
                        if (cz2Only)
                        {
                            patRow.ForceCondition.IsShmooInProdInst = false;
                            patRow.ForceCondition.IsShmooInProdFlow = false;
                            patRow.ForceCondition.IsShmooInCharInst = true;
                            patRow.ForceCondition.IsShmooInCharFlow = true;
                        }
                        else
                        {
                            patRow.ForceCondition.IsShmooInProdInst = true;
                            patRow.ForceCondition.IsShmooInProdFlow = true;
                            patRow.ForceCondition.IsShmooInCharInst = false;
                            patRow.ForceCondition.IsShmooInCharFlow = true;
                        }
                    }
                    else
                    {
                        if (patRow.ForceConditionChar == "")
                            patRow.ForceConditionChar = patRow.ForceCondition.ForceCondition;

                        if (cz2Only)
                        {
                            patRow.ForceCondition.IsShmooInProdInst = false;
                            patRow.ForceCondition.IsShmooInProdFlow = false;
                            patRow.ForceCondition.IsShmooInCharInst = true;
                            patRow.ForceCondition.IsShmooInCharFlow = true;
                        }
                        else
                        {
                            if (sameFlag)
                            {
                                patRow.ForceCondition.IsShmooInProdInst = true;
                                patRow.ForceCondition.IsShmooInProdFlow = true;
                                patRow.ForceCondition.IsShmooInCharInst = false;
                                patRow.ForceCondition.IsShmooInCharFlow = true;
                                patRow.ForceCondition.ForceCondition = patRow.ForceConditionChar;
                            }
                            else
                            {
                                var shmooRow = patRow.DeepClone();
                                shmooRow.ForceCondition.IsShmooInProdInst = false;
                                shmooRow.ForceCondition.IsShmooInProdFlow = false;
                                shmooRow.ForceCondition.IsShmooInCharInst = true;
                                shmooRow.ForceCondition.IsShmooInCharFlow = true;
                                shmooRow.ForceCondition.IsCz2InstName = true;
                                shmooRow.ForceCondition.ForceCondition = patRow.ForceConditionChar;
                                i++;
                                PatternRows.Insert(i, shmooRow);

                                patRow.ForceCondition.IsShmooInProdInst = true;
                                patRow.ForceCondition.IsShmooInProdFlow = true;
                                patRow.ForceCondition.IsShmooInCharInst = false;
                                patRow.ForceCondition.IsShmooInCharFlow = false;
                            }
                        }
                    }
                }
            }
        }

        private bool IsShmooInForce(PatternRow patRow)
        {
            if (Regex.IsMatch(patRow.ForceCondition.ForceCondition, HardIpConstData.RegShmoo,
                RegexOptions.IgnoreCase)) return true;
            if (!string.IsNullOrEmpty(patRow.ForceConditionChar))
                if (Regex.IsMatch(patRow.ForceConditionChar, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase))
                    return true;
            return false;
        }
    }
}