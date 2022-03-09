using System;
using System.Collections.Generic;
using System.Linq;
using PmicAutogen.Local;
using IgxlData.IgxlSheets;
using IgxlData.IgxlBase;
using PmicAutogen.Local.Const;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.Reader;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class SetJATGPinName
    {
        private PinMapSheet _pinMapSheet;
        private static Dictionary<string, string> _JTAGPinGroupMap = new Dictionary<string, string>
        { { "TRST","CTRL0" } };

        public SetJATGPinName()
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

        public void SetPinMapByOTPSetup()
        {
            if (_pinMapSheet == null)
                return;

            List<PinGroup> pinGrpList = new List<PinGroup>();

            OTPSetupSheet otpSetupSheet = StaticTestPlan.OtpSetupSheet;
            if (otpSetupSheet != null)
            {
                Dictionary<string, string> JTAGPinNameMap = otpSetupSheet.GetJTAGPinNameMap();
                foreach (var JTAGPinNameItem in JTAGPinNameMap)
                {
                    PinGroup pinGrp;
                    if (_JTAGPinGroupMap.Keys.Contains(JTAGPinNameItem.Key))
                    {
                        pinGrp = new PinGroup("POP_" + _JTAGPinGroupMap[JTAGPinNameItem.Key]);
                    }
                    else
                    {
                        pinGrp = new PinGroup("POP_" + JTAGPinNameItem.Key);
                    }
                    pinGrp.AddPin(new Pin(JTAGPinNameItem.Value, "I/O", "POPGen"));
                    pinGrpList.Add(pinGrp);
                }
            }

            _pinMapSheet.AddPinGroups(pinGrpList);
        }
    }
}
