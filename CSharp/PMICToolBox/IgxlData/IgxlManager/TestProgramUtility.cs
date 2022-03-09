using System;
using System.Diagnostics;
using System.IO;

namespace IgxlData.IgxlManager
{
    public static class TestProgramUtility
    {
        public static void ExportWorkBookCmd(string testprogramname, string exportfolder)
        {
            if (!FindInstalledOasis())
                throw new Exception("No Oasis Installed!!!");
            CheckFolderExist(exportfolder);
            var option = "-w \"" + testprogramname + "\" -d \"" + exportfolder + "\"";
            // assume ExportWorkbook is in the PATH
            RunCmd("ExportWorkbook", option);
        }

        private static void RunCmd(string cmd, string argment = "")
        {
            var nProcess = new Process();
            var startInfo = new ProcessStartInfo
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                FileName = cmd,
                Arguments = argment
            };
            nProcess.StartInfo = startInfo;
            nProcess.Start();
            nProcess.WaitForExit();
        }

        private static bool FindInstalledOasis()
        {
            string oasisRoot = Environment.GetEnvironmentVariable("OASISROOT");
            if (string.IsNullOrEmpty(oasisRoot))
                return false;
            else
                return true;
        }
        private static void CheckFolderExist(string folder)
        {
            if (Directory.Exists(folder))
                Directory.Delete(folder, true);
            Directory.CreateDirectory(folder);
        }
    }
}