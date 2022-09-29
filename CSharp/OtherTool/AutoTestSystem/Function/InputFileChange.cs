using AutoIgxl;
using AutoProgram;
using AutoTestSystem.Model;
using AutoTestSystem.Setting;
using CommonLib.Enum;
using CommonLib.Utility;
using CommonReaderLib.DebugPlan;
using IgxlData.IgxlReader;
using NLog;
using ShmooLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace AutoTestSystem.Function
{
    internal class InputFileChange
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        public readonly QueueFile QueueFile;
        private SettingIni _settingIni;

        public InputFileChange(QueueFile queueFile)
        {
            QueueFile = queueFile;
        }

        public void Work(DebugPlanMain debugTestPlan)
        {
            var bat = _settingIni.PatternSync;
            if (!string.IsNullOrEmpty(bat))
            {
                _logger.Trace("Starting to sync patterns ...");
                RunCmd(bat, "");
            }

            var outputIgxl = new AutoProgramMain().Main(_settingIni.JobName, _settingIni.TestProgram, _settingIni.PatternFolder,
                 _settingIni.EnableWords, debugTestPlan);

            var output = QueueFile.Output;
            _logger.Trace("Output folder => " + output + " ...");
            if (!Directory.Exists(output))
                Directory.CreateDirectory(output);
            if (!string.IsNullOrEmpty(outputIgxl))
            {
                var runCondition = new RunCondition();
                var igxlSheetReader = new IgxlSheetReader();
                runCondition.Job = _settingIni.JobName;
                runCondition.LotId = _settingIni.LotId;
                runCondition.WaferId = _settingIni.WaferId;
                runCondition.SetXy = _settingIni.SetXy;
                runCondition.ExecSites = _settingIni.Sites.Split(',').ToList();
                var totalSites = igxlSheetReader.GetSites(_settingIni.TestProgram);
                runCondition.TotalSites = totalSites;
                #region enable words
                var enableWords = _settingIni.EnableWords.Split(',').ToList();
                var enableWordsByAutogen = AutoProgramMain.EnableWord.Split(',').ToList();
                enableWords.AddRange(enableWordsByAutogen);
                runCondition.ExecEnableWords = enableWords.Distinct().Where(x => !string.IsNullOrEmpty(x)).ToList();

                var enables = igxlSheetReader.GetEnables(_settingIni.TestProgram);
                enables.AddRange(enableWordsByAutogen);
                runCondition.TotalEnableWords = enables.Distinct().ToList();
                #endregion

                runCondition.DoAll = _settingIni.DoAll.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
                runCondition.OverrideFailStop = _settingIni.OverrideFailStop.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
                runCondition.OutputLog = Path.Combine(Path.GetDirectoryName(outputIgxl),
                    Path.GetFileNameWithoutExtension(debugTestPlan.InputFile) + "_" + _settingIni.JobName + "_" +
                    QueueFile.TimeStamp + ".txt");
                new AutoIgxlMain(outputIgxl).RunProgram(runCondition);

                if (File.Exists(runCondition.OutputLog))
                {
                    QueueFile.ValidationPass = true;
                    var shmooLog = new ShmooLog.ShmooLog(runCondition.OutputLog);
                    shmooLog.ParseEachDevices();
                    var shmooLogs = new ShmooLogs { shmooLog };
                    var report = shmooLogs.ConvertExcel(output);
                    runCondition.OutputReport = report;
                    if (string.IsNullOrEmpty(report))
                        _logger.Error("Output report => No report !!!" + report);
                    else
                        _logger.Trace("Output report => " + report + " ...");
                    var outputLog = Path.Combine(output, Path.GetFileName(runCondition.OutputLog));
                    if (runCondition.OutputLog != outputLog)
                    {
                        _logger.Trace("Output log => " + outputLog + " ...");
                        runCondition.FinalOutputLog = outputLog;
                        File.Copy(runCondition.OutputLog, outputLog, true);
                        File.Delete(runCondition.OutputLog);
                    }
                    PostCheck(outputLog, runCondition);
                }

                if (File.Exists(QueueFile.IniFile))
                {
                    QueueFile.OutputIniFile = Path.Combine(output,
                        Path.GetFileNameWithoutExtension(QueueFile.InputFile) + "_" + Path.GetFileName(QueueFile.IniFile));
                    File.Copy(QueueFile.IniFile, QueueFile.OutputIniFile, true);
                }

                QueueFile.RunCondition = runCondition;
            }
        }

        private void PostCheck(string outputLog, RunCondition runCondition)
        {
            CheckEnableWord(outputLog, runCondition);
        }

        private void CheckEnableWord(string outputLog, RunCondition runCondition)
        {
            var enables = new List<string>();
            using (var sr = new StreamReader(File.OpenRead(outputLog)))
            {
                var startFlag = false;
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (line.Equals("*print: PrintEnableWords end*", StringComparison.CurrentCultureIgnoreCase))
                        break;
                    if (startFlag)
                        if (!line.StartsWith("****************************"))
                            enables.Add(line.Split(':').First());
                    if (line.Equals("*print: PrintEnableWords start*", StringComparison.CurrentCultureIgnoreCase))
                        startFlag = true;
                }
            }

            var less = runCondition.ExecEnableWords.Except(enables).ToList();
            var more = enables.Except(runCondition.ExecEnableWords).ToList();
            if (less.Any())
                _logger.Error("[" + EnumNLogMessage.Input + "] " + "Missing {0} in data log !!!",
                    string.Join(",", less));
            if (more.Any())
                _logger.Error("[" + EnumNLogMessage.Input + "] " + "{0} in data log should not be trun on !!!",
                    string.Join(",", more));
        }

        public void RunCmd(string fileName, string arguments)
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Normal,
                FileName = fileName,
                Arguments = arguments
            };

            nProcess.StartInfo = startInfo;
            nProcess.Start();
            var result = nProcess.StandardOutput.ReadToEnd();
            nProcess.WaitForExit();
            if (!result.Equals(""))
            {
            }
        }

        public bool DoTask(int period, int maxWaitTime)
        {
            _logger.Trace("Starting to process " + QueueFile.InputFile + " ...");

            #region check ini

            var failFlag = false;
            if (!File.Exists(QueueFile.IniFile))
            {
                _logger.Error("[" + EnumNLogMessage.Environment + "] " + "Can not find {0} for {1} !!! ",
                    QueueFile.IniFile, QueueFile.InputFile);
                return true;
            }

            var settingIni = new SettingIni();
            settingIni.Read(QueueFile.IniFile);
            _settingIni = settingIni;
            QueueFile.MailTo = _settingIni.MailTo;
            #endregion

            var debugTestPlan = new DebugPlanMain(QueueFile.InputFile);
            debugTestPlan.Read();
            var reCheck = debugTestPlan.CheckAll(_settingIni.PatternFolder, _settingIni.TimeFolder, _settingIni.TestProgram);

            #region check pattern in dashboard and in pattern folder
            var totalTime = (TimeProvider.Current.Now - QueueFile.Time).Minutes;
            while (totalTime < maxWaitTime)
            {
                if (debugTestPlan.Errors.Count > 0 && reCheck)
                {
                    foreach (var error in debugTestPlan.Errors)
                        _logger.Error("[" + EnumNLogMessage.Input + "] " + error.Message + " ...");
                    _logger.Error("[" + EnumNLogMessage.Input + "] " + "Waiting to receive patterns ...");
                    Thread.Sleep(period * 60 * 1000);
                    reCheck = debugTestPlan.CheckAll(_settingIni.PatternFolder, _settingIni.TimeFolder, _settingIni.TestProgram);
                    failFlag = true;
                }
                else
                {
                    failFlag = false;
                    break;
                }
                totalTime = (TimeProvider.Current.Now - QueueFile.Time).Minutes;
            }

            if (failFlag)
                _logger.Error("[" + EnumNLogMessage.Input + "] " + "Wait time of pattern exceeds than {0} minutes !!! ",
                    maxWaitTime);

            if (debugTestPlan.Errors.Count > 0)
                foreach (var error in debugTestPlan.Errors)
                    _logger.Error("[" + EnumNLogMessage.Input + "] " + error.Message + " ...");

            #endregion

            Work(debugTestPlan);
            _logger.Trace("Ending " + QueueFile.InputFile + " ...");
            return failFlag;
        }
    }
}