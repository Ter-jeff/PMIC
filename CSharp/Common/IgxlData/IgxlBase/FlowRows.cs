using System;
using System.Collections.Generic;
using System.Linq;

namespace IgxlData.IgxlBase
{
    public class FlowRows : List<FlowRow>
    {
        public void Add_A_Enable_MP_SBIN(string parameter)
        {
            var flowRow = new FlowRow
            {
                OpCode = FlowRow.OpCodeBinTable,
                Parameter = "Bin_" + parameter,
                Enable = "A_Enable_MP_SBIN"
            };
            Add(flowRow);
        }

        public void AddDebugPrint()
        {
            Add(new FlowRow
            {
                OpCode = FlowRow.OpCodeTest,
                Parameter = "Debug_Print_Instrument_Status_Check_End",
                Enable = "B_DebugPrint_Instrument_Status"
            });
        }

        public void AddReturnRow()
        {
            Add(new FlowRow
            {
                OpCode = "Return"
            });
        }

        public void AddHeaderRow(string sheetName, string enable)
        {
            Add(new FlowRow
            {
                OpCode = FlowRow.OpCodeTest,
                Parameter = sheetName + "_Header",
                Enable = enable
            });
        }

        public void AddFooterRow(string sheetName, string enable)
        {
            Add(new FlowRow
            {
                OpCode = FlowRow.OpCodeTest,
                Parameter = sheetName + "_Footer",
                Enable = enable
            });
        }

        public void AddPrintStartRow(string sheetName)
        {
            Add(new FlowRow
            {
                OpCode = "Print",
                Parameter = "\"" + sheetName + " Start\""
            });
        }

        public void AddPrintEndRow(string sheetName)
        {
            Add(new FlowRow
            {
                OpCode = "Print",
                Parameter = "\"" + sheetName + " End\""
            });
        }

        public void AddStartRows(string sheetName, string enable = "")
        {
            var arr = sheetName.Split('_').ToList();
            arr.RemoveAt(0);
            var name = string.Join("_", arr);

            //Set Error Bin
            if (Count > 0 &&
                this[0].OpCode.Equals("set-error-bin", StringComparison.CurrentCultureIgnoreCase))
            {
                //pass
            }
            else
            {
                AddSetErrorBinRow();
            }

            //Print
            AddPrintStartRow(name);
            //Header
            AddHeaderRow(name, enable);
        }

        public void AddEndRows(string sheetName, string enable = "", bool isPrintDebug = true)
        {
            var arr = sheetName.Split('_').ToList();
            arr.RemoveAt(0);
            var name = string.Join("_", arr);

            //Footer
            AddFooterRow(name, enable);
            //Print
            AddPrintEndRow(name);
            //Debug print
            if (isPrintDebug)
                AddDebugPrint();
            //Return
            AddReturnRow();
        }

        public void AddClearFlags()
        {
            var flagClears = this.Where(x => !string.IsNullOrEmpty(x.FailAction))
                .Select(x => x.FailAction.Split(',')).SelectMany(x => x)
                .Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            var flowRows = new FlowRows();
            foreach (var flagClear in flagClears)
            {
                var flowRow = new FlowRow();
                flowRow.OpCode = "flag-clear";
                flowRow.Parameter = flagClear;
                flowRows.Add(flowRow);
            }

            InsertRange(0, flowRows);
        }

        public void AddFlowRow(string opCode, string parameter, string enableWord = "")
        {
            var row = new FlowRow();
            row.OpCode = opCode;
            row.Parameter = parameter;
            row.Enable = enableWord;
            Add(row);
        }

        public void AddSetErrorBinRow(string binFail = "999", string sortFail = "999")
        {
            Add(new FlowRow
            {
                OpCode = "set-error-bin",
                BinFail = binFail,
                SortFail = sortFail
            });
        }
    }
}