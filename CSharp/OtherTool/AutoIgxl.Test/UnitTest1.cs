using Microsoft.VisualStudio.TestTools.UnitTesting;
using ShmooLog;
using System;
using System.IO;

namespace AutoIgxl.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void RunIgxl()
        {
            var runCondition = new RunCondition();
            runCondition.LotId = "FM1234";
            runCondition.WaferId = "13";
            var autoIgxlMain = new AutoIgxlMain(Path.Combine(Environment.CurrentDirectory, @"Sample\Sample.igxl"));
            autoIgxlMain.RunProgram(runCondition);
        }
    }
}