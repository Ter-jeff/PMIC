using System.IO;

namespace PinNameRemoval
{
    public class ReverseCompile : CompileResult
    {
        string switchTemplate = "/c aprc \"[inputFileName]\" -output \"[outputFileName]\" -force";
        public ReverseCompile(string inputPath)
        {
            FilePath = inputPath;
            FileName = Path.GetFileName(FilePath);
        }

        public override void Run()
        {
            string switches = switchTemplate.Replace(inputFileName, FilePath).Replace(outputFileName, OutputPath);
            string msg;
            Result = Ctrl.CompilerCLI(FilePath, OutputPath, switches, out msg);
            OutputMsg = msg;
        }
    }
}
