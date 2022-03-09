using System.IO;

namespace PinNameRemoval
{
    public class Compile : CompileResult
    {
        string switchTemplate = "/c apc \"[inputFileName]\" -output \"[outputFileName]\" -pinmap_workbook \"[pinmapWorkbook]\" -digital_inst HSDMQ -opcode_mode single -comments";
        string logFilePath;
        public bool isScanType = false;
        public string PinmapPath { get; set; }

        public Compile(string inputPath, string pinmapPath)
        {
            FilePath = inputPath;
            FileName = Path.GetFileName(FilePath);
            PinmapPath = pinmapPath;
        }

        public override void Run()
        {
            // if scan type, switchTemplate += "-scan_type x2";
            if (isScanType)
                switchTemplate += " -scan_type x2";
            string msg;
            // -logfile \"[logPath]\" 
            string switches = switchTemplate.Replace(inputFileName, FilePath).Replace(outputFileName, OutputPath).Replace(pinmapWorkbook, PinmapPath);
            string nameWoE = Path.GetFileNameWithoutExtension(FilePath);
            logFilePath = FilePath.Replace(".atp", ".log").Replace(nameWoE, "PinRemoval_" + nameWoE);
            switches += " -logfile \"" + logFilePath + "\"";
            try
            {
                File.Delete(OutputPath);
            }
            catch (IOException)
            {
                OutputMsg = "The file is in use.";
                return;
            }
            Result = Ctrl.CompilerCLI(FilePath, OutputPath, switches, out msg);
            //OutputMsg = msg;
            OutputMsg = logFilePath;
        }
    }
}
