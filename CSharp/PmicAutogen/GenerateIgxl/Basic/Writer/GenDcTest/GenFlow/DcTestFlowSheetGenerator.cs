using AutomationCommon.DataStructure;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.DividerManager;
using PmicAutogen.GenerateIgxl.HardIp.DividerManager.FlowDividerManager;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow
{
    public class DcTestFlowSheetGenerator : FlowSheetGeneratorBase
    {
        public DcTestFlowSheetGenerator(string sheetName, List<HardIpPattern> patternList = null) : base(sheetName,
            patternList)
        {
            FlowRowGenerator = new DcTestFlowRowGenerator(sheetName);
        }

        protected override List<HardIpPattern> DividePatterns()
        {
            var dividedPatList = new List<HardIpPattern>();
            foreach (var pattern in PatternList)
                try
                {
                    var isHardIpUniversal = SearchInfo.GetVbtNameByPattern(pattern) == "";
                    if (Regex.IsMatch(pattern.Pattern.GetLastPayload(), HardIpConstData.RegInsInPattern,
                        RegexOptions.IgnoreCase) && pattern.MeasPins.Count == 0)
                    {
                        dividedPatList.Add(pattern);
                        continue;
                    }

                    var tempList1 = DividerMain.DivideMeasPins(pattern, isHardIpUniversal, true);
                    var tempList2 = FlowLimitDivider.DivideUseLimit(tempList1);
                    dividedPatList.AddRange(tempList2);
                }
                catch (Exception e)
                {
                    Response.Report(e.ToString(), MessageLevel.Error, 0);
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }

            return dividedPatList;
        }

        protected override List<FlowRow> GenFlowBodyRows(bool shmooFlag = false)
        {
            var flowBodyRows = new List<FlowRow>();

            var voltages = new List<string>();
            if (SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
            {
                foreach (var pattern in ExtendedPatList)
                {
                    flowBodyRows.AddRange(GenFlowTestRowsByVoltageForPattern(pattern, "", shmooFlag));
                    flowBodyRows.AddRange(WriteBinTableForPattern(pattern));
                }

            }
            else
            {
                voltages.Add("NV");
                voltages.Add("LV");
                voltages.Add("HV");

                if (LocalSpecs.HasUltraVoltageULv)
                {
                    voltages.Add("ULV");
                }

                if (LocalSpecs.HasUltraVoltageUHv)
                {
                    voltages.Add("UHV");
                }

                // according to the voltage order
                //foreach (var voltage in voltages)
                //    flowBodyRows.AddRange(GenFlowTestRowsByVoltage(voltage, shmooFlag));

                //flowBodyRows.AddRange(WriteBinTable());

                //change to according to the pattern order
                foreach (var pattern in ExtendedPatList)
                {
                    flowBodyRows.AddRange(GenFlowTestRowsByPattern(pattern, voltages, shmooFlag));
                }
            }

            //if (!shmooFlag)
            //{
            //    if (ExtendedPatList.Where(p => p.UseDeferLimit).ToList().Count > 0) // if exist test defer limit=> generate limits all
            //        flowBodyRows.Add(new FlowRow { OpCode = "limits-all" });
            //    flowBodyRows.AddRange(GenFlowBinTableRows());
            //    flowBodyRows.AddRange(FlowRowGenerator.GenTtrFlagClearRow(flowBodyRows));
            //}

            return flowBodyRows;
        }

        protected List<FlowRow> GenFlowTestRowsByVoltage(string labelVoltage, bool shmooCharFlag)
        {
            var flowRows = new List<FlowRow>();
            FlowRowGenerator.LabelVoltage = labelVoltage;
            foreach (var pattern in ExtendedPatList)
            {
                try
                {
                    flowRows.AddRange(GenFlowTestRowsByVoltageForPattern(pattern, labelVoltage, shmooCharFlag));
                }
                catch (Exception e)
                {
                    Response.Report(e.ToString(), MessageLevel.Error, 0);
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }

            //flowRows.AddRange(GenResetRelayRows(labelVoltage));
            return flowRows;
        }

        protected List<FlowRow> GenFlowTestRowsByVoltageForPattern(HardIpPattern pattern, string labelVoltage, bool shmooCharFlag)
        {
            var flowRows = new List<FlowRow>();
            if (!IsNeedGenerate(pattern)) return flowRows;

            FlowRowGenerator.Pat = pattern;
            if (FlowRowGenerator.NoNeedToGen)
                return flowRows;
            var sweepShmooRows = FlowRowGenerator.GenShmooRows(labelVoltage);
            if (shmooCharFlag && !pattern.ForceCondition.IsShmooInCharFlow)
                return flowRows;

            var sweepCodeForRow = FlowRowGenerator.GenSweepCodeForRow();
            var sweepVoltageRows = FlowRowGenerator.GenSweepVoltageForRow();
            var retestIfRow = FlowRowGenerator.GenRetestIfRow();
            var retestEndIfRow = FlowRowGenerator.GenRetestEndIfRow();
            var sweepCodeNextRow = FlowRowGenerator.GenSweepCodeOrVoltageNextRow();

            //flowRows.AddRange(FlowRowGenerator.GenExtraRowsByMisc());
            //flowRows.AddRange(FlowRowGenerator.GenRelayRows());
            //flowRows.AddRange(FlowRowGenerator.GenNwireChangeRows());
            //flowRows.AddRange(FlowRowGenerator.GenNwireDisOrEnableRows());

            var opCodeBeforeList = SearchInfo.GetOpCode(pattern, "B");
            CommonGenerator.ConvertPatNameInOpCode(opCodeBeforeList, flowRows, labelVoltage);
            if (!shmooCharFlag && pattern.ForceCondition.IsShmooInProdFlow)
                flowRows.AddRange(FlowRowGenerator.GenOpCodeRowsBefPat(opCodeBeforeList));

            if (sweepVoltageRows != null)
                flowRows.AddRange(sweepVoltageRows);
            if (sweepCodeForRow != null)
                flowRows.AddRange(sweepCodeForRow);
            flowRows.AddRange(FlowRowGenerator.GenPreRetestRows());
            if (retestIfRow != null)
                flowRows.Add(retestIfRow);

            if (shmooCharFlag)
            {
                if (pattern.ForceCondition.IsShmooInCharFlow)
                {
                    flowRows.AddRange(FlowRowGenerator.GenTestRows(true));
                    if (pattern.ForceCondition.IsShmooInForce)
                        flowRows.AddRange(sweepShmooRows);
                }
            }
            else
            {
                if (pattern.ForceCondition.IsShmooInProdFlow)
                {
                    if (pattern.ForceCondChar == null)
                        flowRows.AddRange(FlowRowGenerator.GenTestRows());
                    else
                        flowRows.AddRange(FlowRowGenerator.GenTestRows());
                }
            }

            if (retestEndIfRow != null)
                flowRows.Add(retestEndIfRow);
            if (sweepCodeNextRow != null)
                flowRows.AddRange(sweepCodeNextRow);

            if (!shmooCharFlag && pattern.ForceCondition.IsShmooInProdFlow)
                flowRows.AddRange(FlowRowGenerator.GenOpCodeRowsAftPat());

            if (Regex.IsMatch(pattern.AcUsed, @"XI0|RT_CLK", RegexOptions.IgnoreCase))
                flowRows.Add(FlowRowGenerator.GenNWireHardIpRow());
            return flowRows;
        }

        private List<FlowRow> WriteBinTable()
        {
            var flowRows = new List<FlowRow>();
            foreach (var pattern in ExtendedPatList)
            {
                flowRows.AddRange(WriteBinTableForPattern(pattern));
            }

            return flowRows;
        }

        private List<FlowRow> WriteBinTableForPattern(HardIpPattern pattern)
        {
            var flowRows = new List<FlowRow>();
            FlowRowGenerator.Pat = pattern;
            if (SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
                flowRows.Add(FlowRowGenerator.GenBinTableRow(""));
            else
            {
                flowRows.Add(FlowRowGenerator.GenBinTableRow("HNLV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("NLV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("HLV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("HNV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("HV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("NV"));
                flowRows.Add(FlowRowGenerator.GenBinTableRow("LV"));
                if (LocalSpecs.HasUltraVoltageUHv)
                {
                    flowRows.Add(FlowRowGenerator.GenBinTableRow("UHV"));
                }
                if (LocalSpecs.HasUltraVoltageULv)
                {
                    flowRows.Add(FlowRowGenerator.GenBinTableRow("ULV"));
                }
            }
            return flowRows;
        }

        private bool IsNeedGenerate(HardIpPattern pattern)
        {
            if (pattern.IsNonHardIpBlock)
            {
                if (pattern.MeasPins.Count == 0)
                    if (pattern.Pattern.InstancePayloadName.Count == 0)
                        return false;
                return true;
            }

            return true;
        }

        private List<FlowRow> GenFlowTestRowsByPattern(HardIpPattern pattern, List<string> voltagelist, bool shmooFlag)
        {
            var flowRows = new List<FlowRow>();
            foreach (var labelVoltage in voltagelist)
            {
                FlowRowGenerator.LabelVoltage = labelVoltage;
                try
                {
                    flowRows.AddRange(GenFlowTestRowsByVoltageForPattern(pattern, labelVoltage, shmooFlag));
                }
                catch (Exception e)
                {
                    Response.Report(e.ToString(), MessageLevel.Error, 0);
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }
            }

            flowRows.AddRange(WriteBinTableForPattern(pattern));

            return flowRows;
        }
    }
}