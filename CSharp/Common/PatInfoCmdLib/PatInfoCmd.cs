using System;
using System.Diagnostics;
using System.IO;

namespace PatInfoCmdLib
{
    public class PatInfoCmd
    {
        public bool ConvertByArgs(string sourcePat, ref string output, string strArg)
        {
            string args = string.Format(" {0} \"{1}\"", strArg, sourcePat);
            try
            {
                var p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardOutput = true;
                var getEnvironmentVariable = Environment.GetEnvironmentVariable("igxlroot");
                if (string.IsNullOrEmpty(getEnvironmentVariable))
                    p.StartInfo.FileName = Path.Combine(getEnvironmentVariable, "bin", "patinfo.exe");
                else
                    p.StartInfo.FileName = Path.Combine(Directory.GetCurrentDirectory(), "patinfo.exe");
                p.StartInfo.Arguments = args;
                p.Start();

                output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                if (p.ExitCode == 0)
                    return true;
                return false;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}