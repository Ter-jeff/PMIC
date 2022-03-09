using System.Collections.Generic;
using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;
using System;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenInstance
{
    public class DcTestBlockInsGenerator : BlockInstanceGenerator
    {
        public DcTestBlockInsGenerator(string sheetName, List<HardIpPattern> pattenList) : base(sheetName, pattenList)
        {
            InstanceRowGenerator = new DcTestInsRowGenerator(sheetName);
        }

        public override List<InstanceSheet> GenBlockInsRows()
        {
            var instanceSheetList = new List<InstanceSheet>();

            foreach (var hardIpPattern in HardIpPatterns)
            {
                if (SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
                {
                    instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, ""));
                }
                else
                {
                    instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, HardIpConstData.LabelNv));
                    instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, HardIpConstData.LabelLv));
                    instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, HardIpConstData.LabelHv));
                    if (LocalSpecs.HasUltraVoltageULv)
                    {
                        instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, HardIpConstData.LabelULv));
                    }
                    if (LocalSpecs.HasUltraVoltageUHv)
                    {
                        instanceSheetList.AddRange(GenBlockInsRowsByVoltage(hardIpPattern, HardIpConstData.LabelUHv));
                    }

                }
            }
            return instanceSheetList;
        }

        private List<InstanceSheet> GenBlockInsRowsByVoltage(HardIpPattern hardIpPattern, string labelVoltage)
        {
            var instanceSheetList = new List<InstanceSheet>();
            var voltage = new InstanceSheet(HardIpConstData.PrefixInsSheetByVoltage + labelVoltage);

            hardIpPattern.FunctionName = VbtFunctionLib.FunctionalTUpdated;
            InstanceRowGenerator.LabelVoltage = labelVoltage;
            InstanceRowGenerator.Pat = hardIpPattern;
            var insRowList = InstanceRowGenerator.GenInsRows();
            foreach (var insRow in insRowList)
                voltage.AddRow(insRow);

            if (voltage.InstanceRows.Count != 0) instanceSheetList.Add(voltage);
            return instanceSheetList;
        }
    }
}