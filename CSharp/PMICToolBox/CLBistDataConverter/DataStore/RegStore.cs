using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CLBistDataConverter
{
    static class RegStore
    {
        public static Regex RegLine = new Regex(@"\S*", RegexOptions.IgnoreCase);
        //^(?<Number>\d+)\s+(?<Site>\d+)\s+(?<TestName>\S+)\s+((?<Pin>\w+)\s+)?((?<Channel>(N/A)|(-1)|(\d+.\w+))\s+)?(?<Low>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Measured>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?(\s+[(][a-z]+[)])?))\s+(?<High>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Force>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Loc>\d+)$
        public static Regex RegClBistDatalogRow = new Regex(@"^(?<Number>\d+)\s+(?<Site>\d+)\s+(?<TestName>\S+)\s+((?<Pin>\w+)\s+)?(?<Low>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Measured>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?(\s+[(][a-z]+[)])?))\s+(?<High>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Force>(N/A)|(([-]?\d+([.]\d+)?)(\s[a-z]+)?))\s+(?<Loc>\d+)$", RegexOptions.IgnoreCase);
        public static Regex RegOTP_UpdateECID_Instance = new Regex(@"^<OTP_UpdateECID>$", RegexOptions.IgnoreCase);
        public static Regex RegReadLotId = new Regex(@"^Site[\(](?<site>[\d]+)[\)](\s)*OTP(\s)*Read(\s)*LotID(\s)*=(\s)*(?<lotid>.*)$", RegexOptions.IgnoreCase);
        public static Regex RegReadWaferId = new Regex(@"^Site[\(](?<site>[\d]+)[\)](\s)*OTP(\s)*Read(\s)*WaferID(\s)*=(\s)*(?<waferid>[\w]+)$", RegexOptions.IgnoreCase);
        public static Regex RegReadXCoord = new Regex(@"^Site[\(](?<site>[\d]+)[\)](\s)*OTP(\s)*Read(\s)*X_Coord(\s)*=(\s)*(?<xcoord>[\w]+)$", RegexOptions.IgnoreCase);
        public static Regex RegReadYCoord = new Regex(@"^Site[\(](?<site>[\d]+)[\)](\s)*OTP(\s)*Read(\s)*Y_Coord(\s)*=(\s)*(?<ycoord>[\w]+)$", RegexOptions.IgnoreCase);
        public static Regex RegReadActStrM = new Regex(@"^Site[\(](?<site>[\d]+)[\)](\s)*READ(\s)*DUT(\s)*ECID(\s)*READBACK(\s)*ActStrM(\s)*=(\s)*(?<actstrm>[\w]+)", RegexOptions.IgnoreCase);
        public static Regex RegClk_Cfg5 = new Regex(@"^bDAC(?<bdacnumber>[\d]+)_clk_cfg5(\s)*=(\s)*(?<clk_cfg5>[\w]+)$", RegexOptions.IgnoreCase);
        public static Regex RegInstanceNameRow = new Regex(@"^[\<](?<instancename>[\w]+)[\>]$", RegexOptions.IgnoreCase);
        public static Regex RegDacNumber = new Regex(@"_bDAC(?<dacNumber>[\d]+)_", RegexOptions.IgnoreCase);
        public static Regex RegFreq = new Regex(@"(?<freq>([\d]|[\.])+)MHz", RegexOptions.IgnoreCase);
        public static Regex RegBIRef = new Regex(@"bIRef(?<biref>([\d])+)", RegexOptions.IgnoreCase);
        public static Regex RegPhase = new Regex(@"^[A-Z]+(?<phase>[\d]+)$", RegexOptions.IgnoreCase);
        public static Regex RegVddh = new Regex(@"VDDH(?<vddh>([\d]|[\.])+)", RegexOptions.IgnoreCase);
    }
}
