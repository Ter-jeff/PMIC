using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PmicAutomation.Utility.ErrorHandler;
using Library.Common;

namespace ErrorHandlerCommandLine
{
    class ErrorHandlerCommandLine
    { 
        public static void Help()
        {
            string[] help = new string[]
            {
                "ErrorHandler Check Tool",
                "Command-line arguments:",
                "-tpPath    test program file path(required)     -- IGXL test program file path(.igxl or .xlsm file)",
                "-addErrorHandler add missing errorHandler vbt code, 1 or 0, default is 0(optional)   -- Wether add missing errorHandler vbt code to vbt module",
                "-output    output file path(required)         -- The output file path(.xlsx)",
                "-logPath   log folder path(optional)              -- Log folder path",
                "Example: -tpPath testProgramPath -addErrorHandler 1 -output OutputPath -logPath logPath",
                "Surround file/folder paths with double quotes if they contain spaces"
            };

            foreach (string s in help)
            {
                Console.WriteLine("{0}", s);
            }
        }
        public static void Main(string[] args)
        {
            try
            {
                string tpPath = null;
                string outputFilePath = null;
                string logPath = null;
                string logFile = null;
                bool addErrorHandler = false;
                if (!ParseArgList(args, ref tpPath, ref outputFilePath, ref logPath, ref addErrorHandler))
                    return;
                if (tpPath == null)
                    throw new ArgumentException("Missing igxl test program file path, input '-?' to get help.");
                if (outputFilePath == null)
                    throw new ArgumentException("Missing output folder path, input '-?' to get help.");

                if (logPath != null)
                    logFile = Path.Combine(logPath, "ErrorHandler.log");

                Logger.PrintLog(logFile, "Start to do ErrorHandler check......");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": Start to do ErrorHandler check......");

                try
                {
                    new ErrorHandlerMain(tpPath, outputFilePath, addErrorHandler).WorkFlow();
                }
                catch (Exception e)
                {
                    Logger.PrintLog(logFile, "ErrorHandler check error: " + e.ToString());
                    throw new Exception(e.ToString());
                }
                Logger.PrintLog(logFile, "ErrorHandler check success!");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": ErrorHandler check success!");
            }

            catch (Exception ex)
            {
                Console.Error.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": ErrorHandler check error: {0}", ex.Message);
            }
        }

        private static bool ParseArgList(string[] args, ref string tpPath, ref string outputPath, ref string logPath, ref bool addErrorHandler)
        {
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-?":
                        Help();
                        if (args.Length == 1)
                            return false;
                        break;
                    case "-tppath":
                        i += 1;
                        tpPath = CheckTestProgramPath(i, args);
                        break;
                    case "-adderrorhandler":
                        i += 1;
                        addErrorHandler = CheckAddErrorHandlerArg(i, args);
                        break;
                    case "-output":
                        i += 1;
                        outputPath = CheckOutputFilePath(i, args);
                        break;
                    case "-logpath":
                        i += 1;
                        logPath = CheckLogPath(i, args);
                        break;
                    default:
                        break;
                }
            }
            return true;
        }

        private static bool CheckAddErrorHandlerArg(int index, string[] args)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing AddErrorHandler value at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }
            if (!args[index].Equals("0", StringComparison.OrdinalIgnoreCase) &&
                !args[index].Equals("1", StringComparison.OrdinalIgnoreCase))
            {
                message = string.Format("AddErrorHandler value at argument {1} is not valid, it must be 1 or 0, input '-?' to get help.", args[index], index);
                throw new ArgumentException(message);
            }
            bool result = args[index].Equals("1", StringComparison.OrdinalIgnoreCase) ? true : false;
            return result;
        }

        private static string CheckTestProgramPath(int index, string[] args)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing igxl test program file path at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }
            if (!args[index].EndsWith(".igxl", StringComparison.OrdinalIgnoreCase) &&
                !args[index].EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                message = string.Format("The input test program file {0} is not an igxl test program file at argument {1}, input '-?' to get help.", args[index], index);
                throw new ArgumentException(message);
            }
            if (File.Exists(args[index]))
            {
                return args[index];
            }
            else
            {
                message = string.Format("Test program file {0} not found, input '-?' to get help.", args[index]);
                throw new ArgumentException(message);
            }
        }

        private static string CheckLogPath(int index, string[] args)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing log folder path at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }

            try
            {
                if (!Directory.Exists(args[index]))
                    Directory.CreateDirectory(args[index]);
            }
            catch (Exception ex)
            {
                message = string.Format("Create log path directory error, maybe the log path {0} is not valid: {1}", args[index], ex.Message);
                throw new ArgumentException(message);
            }
            return args[index];
        }
        private static string CheckOutputFilePath(int index, string[] args)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing output file path at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }
            if (!args[index].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                message = string.Format("The output file {0} is not an xlsx excel file at argument {1}, input '-?' to get help.", args[index], index);
                throw new ArgumentException(message);
            }

            //if output folder is not exist, create the folder
            
            try
            {
                string outputDir = new FileInfo(args[index]).DirectoryName;
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);
            }
            catch (Exception ex)
            {
                message = string.Format("Create output directory error, maybe the output path {0} is not valid: {1}", args[index], ex.Message);
                throw new ArgumentException(message);
            }

            //if output file is exist, overwrite it
            if (File.Exists(args[index]))
            {
                File.Delete(args[index]);
            }

            return args[index];
        }

    }
}
