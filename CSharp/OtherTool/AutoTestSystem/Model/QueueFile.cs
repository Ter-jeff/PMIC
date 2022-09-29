using AutoIgxl;
using System;
using System.IO;
using System.Text;

namespace AutoTestSystem.Model
{
    public class QueueFile
    {
        public string InputFile { get; set; }
        public string IniFile { get; set; }

        public string TimeStamp
        {
            get { return Time.ToString("yyyyMMdd_HHmmss"); }
        }

        public bool ValidationPass { get; set; }

        public DateTime Time { get; set; }

        public string Output
        {
            get { return Path.GetDirectoryName(InputFile).Replace("Input", "Output"); }
        }

        public RunCondition RunCondition { get; set; }
        public string MailTo { get; set; }
        public string OutputProcessLog { get; set; }
        public string OutputIniFile { get; set; }

        public string Print()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<h2>Input Condition</h2>");
            sb.AppendLine("<ul>");
            sb.AppendLine("<li>Input File : " + InputFile + "</li>");
            sb.AppendLine("<li>Start Time : " + TimeStamp + "</li>");
            if (ValidationPass)
                sb.AppendLine(@"<li style='color: blue' >Validation Fail : " + "Pass" + "</li>");
            else
                sb.AppendLine(@"<li style='color: red' > Validation Fail : " + "Fail" + "</li>");
            sb.AppendLine("</ul>");

            if (RunCondition != null)
            {
                sb.AppendLine("<h2>Run Condition</h2>");
                sb.AppendLine("<ul>");
                sb.AppendLine("<li>Tester : " + RunCondition.Tester + "</li>");
                sb.AppendLine("<li>Job : " + RunCondition.Job + "</li>");
                sb.AppendLine("<li>Lot Id : " + RunCondition.LotId + "</li>");
                sb.AppendLine("<li>Wafer Id : " + RunCondition.WaferId + "</li>");
                sb.AppendLine("<li>Position : " + RunCondition.SetXy + "</li>");
                sb.AppendLine("<li>EnableWords : " + string.Join(",", RunCondition.ExecEnableWords) + "</li>");
                sb.AppendLine("<li>DoAll : " + RunCondition.DoAll + "</li>");
                sb.AppendLine("<li>OverrideFailStop : " + RunCondition.OverrideFailStop + "</li>");
                sb.AppendLine("</ul>");
            }


            sb.AppendLine("<h2>Output Files</h2>");
            sb.AppendLine("<ul>");
            if (string.IsNullOrEmpty(OutputProcessLog))
                sb.AppendLine("<li style='color: red'>Process Log : " + OutputProcessLog + "</li>");
            else
                sb.AppendLine("<li>Process Log : " + OutputProcessLog + "</li>");
            if (string.IsNullOrEmpty(OutputIniFile))
                sb.AppendLine("<li style='color: red'>Setting.ini : " + OutputIniFile + "</li>");
            else
                sb.AppendLine("<li>Setting.ini : " + OutputIniFile + "</li>");
            if (RunCondition != null)
            {
                if (string.IsNullOrEmpty(RunCondition.FinalOutputLog))
                    sb.AppendLine("<li style='color: red'>Data Log : " + RunCondition.FinalOutputLog + "</li>");
                else
                    sb.AppendLine("<li>Data Log : " + RunCondition.FinalOutputLog + "</li>");
                if (string.IsNullOrEmpty(RunCondition.OutputReport))
                    sb.AppendLine("<li style='color: red'>Shmoo Report : " + RunCondition.OutputReport + "</li>");
                else
                    sb.AppendLine("<li>Shmoo Report : " + RunCondition.OutputReport + "</li>");
            }
            sb.AppendLine("</ul>");
            return sb.ToString();
        }
    }
}