using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class CharSetupConst
    {
        public const string TestMethodReBurst = "Reburst";
        public const string TestMethodRetest = "Retest";
        public const string TestMethodReBurstSerial = "Reburst Serial";
        public const string TestMethodRunFunction = "Run Function";
        public const string TestMethodRunPattern = "Run Pattern";

        public static readonly Dictionary<string, string> TestMethod =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"Reburst", "Reburst"}, {"Retest", "Retest"}, {"ReburstSerial", "Reburst Serial"},
                {"RunFunction", "Run Function"}, {"RunPattern", "Run Pattern"}
            };
    }
}