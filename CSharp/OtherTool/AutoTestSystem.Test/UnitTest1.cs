using AutoIgxl;
using AutoTestSystem.Model;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace AutoTestSystem.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var queueFile = new QueueFile();
            queueFile.InputFile = "InputFile";
            queueFile.IniFile = "IniFile";
            queueFile.RunCondition = new RunCondition();
            queueFile.RunCondition.Job = "Job";
            queueFile.RunCondition.LotId = "LotId";
            queueFile.RunCondition.WaferId = "WaferId";
            queueFile.RunCondition.SetXy = "SetXy";
            queueFile.RunCondition.ExecEnableWords = new List<string>() { "EnableWords" };
            queueFile.RunCondition.OutputLog = "OutputLog";
            queueFile.RunCondition.FinalOutputLog = "FinalOutputLog";
            queueFile.RunCondition.OutputReport = "OutputReport";
            queueFile.OutputIniFile = "OutputIniFile";
            queueFile.OutputProcessLog = "OutputProcessLog";

            new TeradyneMail().SendMail("jeff.li@teradyne.com", queueFile);
        }
    }
}
