using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.Others
{
    public class DivideFlowMain
    {
        public Dictionary<SubFlowSheet, string> WorkFlow(Dictionary<string, SubFlowSheet> subFlowSheets)
        {
            foreach (var flow in subFlowSheets)
            {
                _cnt = 0;
                if (flow.Value.FlowRows.Count > MaxDivideRow)
                {
                    var index = GetDivideIndex(flow);
                    if (index > 0)
                    {
                        var newFlow = DivideFlow(flow, index);
                        var index1 = GetDivideIndex(newFlow);
                        while (index1 > MaxDivideRow)
                        {
                            newFlow = DivideFlow(newFlow, index1);
                            index1 = GetDivideIndex(newFlow);
                        }
                    }
                }
            }

            return _tempSubFlowSheets;
        }

        private KeyValuePair<string, SubFlowSheet> DivideFlow(KeyValuePair<string, SubFlowSheet> flow, int index)
        {
            var newFlow = new KeyValuePair<string, SubFlowSheet>();
            if (index > 0)
            {
                _cnt = _cnt + 1;
                var location = flow.Value.SheetName.IndexOf("Part", StringComparison.Ordinal);
                var newSheetName = location == -1
                    ? flow.Value.SheetName + "_Part" + _cnt
                    : flow.Value.SheetName.Substring(0, location) + "Part" + _cnt;
                var newSubSheet = new SubFlowSheet(newSheetName);
                newSubSheet.FlowRows.AddRange(flow.Value.FlowRows.GetRange(index, flow.Value.FlowRows.Count - index));
                var key = flow.Key.Replace(flow.Value.SheetName, newSheetName);
                newFlow = new KeyValuePair<string, SubFlowSheet>(key, newSubSheet);
                _tempSubFlowSheets.Add(newSubSheet, key);
                flow.Value.FlowRows.RemoveRange(index, flow.Value.FlowRows.Count - index);
                flow.Value.FlowRows.Add(GetSubFlowCall(newSheetName));
                flow.Value.FlowRows.Add(GetReturnCall());
            }

            return newFlow;
        }

        private int GetDivideIndex(KeyValuePair<string, SubFlowSheet> flow)
        {
            var loopList = flow.Value.FlowRows.Where(x => x.OpCode != null)
                .Select((v, i) => new { opcode = v.OpCode, index = i })
                .Where(x => Regex.IsMatch(x.opcode, "if|else|endif|for|next|test", RegexOptions.IgnoreCase)).ToList();
            var ifValue = 0;
            var forValue = 0;
            var index = 0;
            foreach (var item in loopList)
            {
                if (Regex.IsMatch(item.opcode, "test|if|for", RegexOptions.IgnoreCase) & (item.index > MaxDivideRow) &
                    (ifValue == 0) & (forValue == 0))
                {
                    index = item.index;
                    break;
                }

                if (item.opcode.Equals("if", StringComparison.OrdinalIgnoreCase))
                    ifValue++;
                else if (item.opcode.Equals("endif", StringComparison.OrdinalIgnoreCase))
                    ifValue--;
                else if (item.opcode.Equals("for", StringComparison.OrdinalIgnoreCase))
                    forValue++;
                else if (item.opcode.Equals("next", StringComparison.OrdinalIgnoreCase))
                    forValue--;
            }

            return index;
        }

        private FlowRow GetSubFlowCall(string subSheetName)
        {
            return new FlowRow { OpCode = "Call", Parameter = subSheetName };
        }

        private FlowRow GetReturnCall()
        {
            return new FlowRow { OpCode = "return" };
        }

        #region Field

        private int _cnt;
        private const int MaxDivideRow = 100000 / 2;
        private readonly Dictionary<SubFlowSheet, string> _tempSubFlowSheets = new Dictionary<SubFlowSheet, string>();

        #endregion

        #region Constructor

        #endregion
    }
}