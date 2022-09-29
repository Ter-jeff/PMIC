using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class SetJatgPinName
    {
        private static readonly Dictionary<string, string> JtagPinGroupMap = new Dictionary<string, string>
            {{"TRST", "CTRL0"}};

        private readonly PinMapSheet _pinMapSheet;

        public SetJatgPinName()
        {
            _pinMapSheet = TestProgram.IgxlWorkBk.PinMapPair.Value;
        }

        //public void SetPinMapByOTPSetup()
        //{
        //    if(_pinMapSheet == null)
        //        return;

        //    List<PinGroup> pinGrpList = new List<PinGroup>();

        //    PortMapSheet portMapSheet = TestProgram.IgxlWorkBk.PortMapSheets.Count > 0 ? TestProgram.IgxlWorkBk.PortMapSheets.Values.ToList()[0] : null;
        //    if (portMapSheet != null)
        //    {
        //        List<PortSet> nwireJtagPortSetlist = portMapSheet.PortSets.FindAll(s => s.PortName.Equals("NWIRE_JTAG", StringComparison.OrdinalIgnoreCase));
        //        if (nwireJtagPortSetlist != null)
        //        {
        //            foreach (PortSet nwireJtagPortSet in nwireJtagPortSetlist)
        //            {
        //                foreach (var row in nwireJtagPortSet.PortRows)
        //                {
        //                    PinGroup pinGrp;
        //                    if(_JATPPinNameMap.Keys.Contains(row.FunctionName))
        //                    {
        //                        pinGrp = new PinGroup("POP_"+ _JATPPinNameMap[row.FunctionName], "I/O");
        //                    }
        //                    else
        //                    {
        //                        pinGrp = new PinGroup("POP_" + row.FunctionName, "I/O");
        //                    }
        //                    pinGrp.AddPin(new Pin(row.FunctionPin,"I/O","POPGen"));
        //                    pinGrpList.Add(pinGrp);
        //                }
        //            }
        //        }
        //    }
        //    _pinMapSheet.AddPinGroups(pinGrpList);
        //}

        public void SetPinMapByOtpSetup()
        {
            if (_pinMapSheet == null)
                return;

            var pinGrpList = new List<PinGroup>();

            var otpSetupSheet = StaticTestPlan.OtpSetupSheet;
            if (otpSetupSheet != null)
            {
                var jtagPinNameMap = otpSetupSheet.GetJtagPinNameMap();
                foreach (var jtagPinNameItem in jtagPinNameMap)
                {
                    PinGroup pinGrp;
                    if (JtagPinGroupMap.Keys.Contains(jtagPinNameItem.Key))
                        pinGrp = new PinGroup("POP_" + JtagPinGroupMap[jtagPinNameItem.Key]);
                    else
                        pinGrp = new PinGroup("POP_" + jtagPinNameItem.Key);
                    pinGrp.AddPin(new Pin(jtagPinNameItem.Value, "I/O", "POPGen"));
                    pinGrpList.Add(pinGrp);
                }
            }

            _pinMapSheet.AddPinGroups(pinGrpList);
        }
    }
}