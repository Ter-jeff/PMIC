using CommonLib.Enum;
using System;
using System.Threading;

namespace CommonLib.WriteMessage
{
    public class Response
    {
        private static IProgress<ProgressStatus> progress { get; set; }

        public static IProgress<ProgressStatus> Progress
        {
            set
            {
                progress = value;
            }
        }

        public static void Report(string message, EnumMessageLevel messageLevel, int percentage)
        {
            var progressStatus = new ProgressStatus();
            progressStatus.Message = message;
            progressStatus.Level = messageLevel;
            progressStatus.Percentage = percentage;
            Thread.Sleep(10);
            if (progress != null)
                progress.Report(progressStatus);
        }
    }
}