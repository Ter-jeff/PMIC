using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CLBistDataConverter.Libs
{
    public static class ReportLib
    {
        public static Action<string, MessageLevel> HandllerWriteMsg = null;
        public static Action<int, int> HandllerReportPrgress = null;
        public static Action<string, int, int> HandllerReportPrgressAndMsg = null;
        public static Action<string> HandllerReportState = null;

        public static void WriteMsg(string msg, MessageLevel level = MessageLevel.info)
        {
            HandllerWriteMsg?.Invoke(msg, level);
        }

        public static void ReportProgress(int val, int max = 100)
        {
            HandllerReportPrgress?.Invoke(val, max);
        }

        public static void ReportPrgressAndMsg(string msg, int val, int max = 100)
        {
            HandllerReportPrgressAndMsg?.Invoke(msg, val, max);
        }

        public static void ReportState(string state)
        {
            HandllerReportState?.Invoke(state);
        }
    }

    public enum MessageLevel
    {
        info,
        warn,
        err
    }
}
