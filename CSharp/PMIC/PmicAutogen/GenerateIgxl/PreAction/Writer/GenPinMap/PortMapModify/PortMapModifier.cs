using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PreAction.Writer.GenPinMap.PortMapModify
{
    public class PortMapModifier
    {
        public void WorkFlow(PortMapSheet portMapSheet, ref List<PortSet> validPortSets)
        {
            var portDefineSheet = StaticTestPlan.PortDefineSheet;
            if (portDefineSheet != null)
            {
                var portDefineDic = portDefineSheet.GroupByPortName();
                var validPortDefineRows = portDefineDic.Where(pair =>
                {
                    if (pair.Value.Any(o => string.IsNullOrEmpty(o.Pin))) return false;
                    return true;
                }).Select(o => o).ToList();

                //update pin name
                if (validPortDefineRows.Count() > 0)
                    foreach (var portSet in portMapSheet.PortSets)
                    {
                        var validPortDefineRow = validPortDefineRows.FindAll(o =>
                            o.Key.Equals(portSet.PortName, StringComparison.CurrentCultureIgnoreCase));
                        if (validPortDefineRow.Count() == 0)
                            continue;
                        foreach (var row in validPortDefineRow.First().Value)
                            if (portSet.PortRows.Exists(x =>
                                    x.PortName.Equals(row.ProtocolPortName,
                                        StringComparison.CurrentCultureIgnoreCase) &&
                                    x.FunctionName.Equals(row.Type, StringComparison.CurrentCultureIgnoreCase)))
                            {
                                var portRows = portSet.PortRows.Where(x =>
                                    x.PortName.Equals(row.ProtocolPortName,
                                        StringComparison.CurrentCultureIgnoreCase) &&
                                    x.FunctionName.Equals(row.Type, StringComparison.CurrentCultureIgnoreCase));
                                foreach (var portRow in portRows)
                                    portRow.FunctionPin = row.Pin;
                            }
                    }

                //adjust the order of prot group in PortMapsheet
                var validPortNameList = validPortDefineRows.Select(o => o.Key).ToList();

                var newOrderPortSetList = new List<PortSet>();
                var invalidPortSetList = new List<PortSet>();

                foreach (var portSet in portMapSheet.PortSets)
                {
                    if (string.IsNullOrEmpty(portSet.PortName.Trim()) && newOrderPortSetList.Count > 0)
                    {
                        newOrderPortSetList.Add(portSet);
                        continue;
                    }

                    if (validPortNameList.Any(o =>
                            o.Equals(portSet.PortName.Trim(), StringComparison.CurrentCultureIgnoreCase)))
                    {
                        newOrderPortSetList.Add(portSet);
                        validPortSets.Add(portSet);
                    }
                    else
                    {
                        invalidPortSetList.Add(portSet);
                    }
                }

                //add an empty row
                var emptyPortSet = new PortSet("");
                emptyPortSet.AddPortRow(new PortRow());
                newOrderPortSetList.Add(emptyPortSet);
                newOrderPortSetList.AddRange(invalidPortSetList);

                portMapSheet.PortSets = newOrderPortSetList;
            }
        }
    }
}