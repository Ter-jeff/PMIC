using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace PinNameRemoval
{
    public class CompileResult : Compiler
    {
        public static string inputFileName = "[inputFileName]";
        public static string outputFileName = "[outputFileName]";
        public static string pinmapWorkbook = "[pinmapWorkbook]";
        public static string logPath = "[logPath]";

        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string OutputPath { get; set; }
        public bool Result = false;
        public string OutputMsg { get; set; }
        public CompileResult()
        {
            Result = false;
            OutputMsg = string.Empty;
        }

        public virtual void Run()
        {
            throw new NotImplementedException();
        }
    }

    public interface Compiler
    {
        void Run();
    }

    public class PatternInfo
    {
        public string InputPath { get; set; }
        public string PinmapPath { get; set; }
        public string TempFolder = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Temp");
        public string OutputPath { get; set; }
        public string FileNameWithoutExt;
        public ReverseCompile APRC { get; set; }
        public Compile APC { get; set; }

        public PatternInfo(string inputPath, string pinmapPath, string inputFolder, string outputFolder)
        {
            InputPath = inputPath;
            PinmapPath = pinmapPath;
            FileNameWithoutExt = Path.GetFileNameWithoutExtension(InputPath);
            Directory.CreateDirectory(TempFolder);

            APRC = new ReverseCompile(InputPath);
            //APRC.OutputPath = Path.Combine(TempFolder, FileNameWithoutExt + ".atp");
            APRC.OutputPath = inputPath.Replace(".PAT", ".atp").Replace(".pat", ".atp").Replace(inputFolder, TempFolder);
            Directory.CreateDirectory(Path.GetDirectoryName(APRC.OutputPath));

            APC = new Compile(APRC.OutputPath, pinmapPath);
            string folder = Path.GetDirectoryName(InputPath);
            APC.OutputPath = OutputPath = Path.Combine(folder, FileNameWithoutExt + ".pat").Replace(inputFolder, outputFolder);
            Directory.CreateDirectory(Path.GetDirectoryName(OutputPath));
        }

        public static PatternInfo ReadPatternInfo(string inputPath, string pinmapPath, string inputFolder, string outputFolder)
        {
            Regex compressed = new Regex(".pat.gz", RegexOptions.IgnoreCase);
            if (compressed.Match(inputPath).Success)
                return new CompressedPatternInfo(inputPath, pinmapPath, inputFolder, outputFolder);
            else
                return new PatternInfo(inputPath, pinmapPath, inputFolder, outputFolder);
        }
    }

    public class CompressedPatternInfo : PatternInfo
    {
        public CompressedPatternInfo(string inputPath, string pinmapPath, string inputFolder, string outputFolder)
            : base(inputPath, pinmapPath, inputFolder, outputFolder)
        {
            InputPath = inputPath;
            PinmapPath = pinmapPath;
            FileNameWithoutExt = Path.GetFileNameWithoutExtension(InputPath);
            FileNameWithoutExt = Path.GetFileNameWithoutExtension(FileNameWithoutExt);
            Directory.CreateDirectory(TempFolder);

            APRC = new ReverseCompile(InputPath);
            //APRC.OutputPath = Path.Combine(TempFolder, FileNameWithoutExt + ".atp");
            APRC.OutputPath = inputPath.Replace(".PAT.gz", "_gz.atp").Replace(".pat.gz", "_gz.atp").Replace(inputFolder, TempFolder);
            Directory.CreateDirectory(Path.GetDirectoryName(APRC.OutputPath));

            APC = new Compile(APRC.OutputPath, pinmapPath);
            string folder = Path.GetDirectoryName(InputPath);
            APC.OutputPath = OutputPath = Path.Combine(folder, FileNameWithoutExt + ".pat.gz").Replace(inputFolder, outputFolder).Replace("_gz", "");
            Directory.CreateDirectory(Path.GetDirectoryName(OutputPath));
        }
    }
}
