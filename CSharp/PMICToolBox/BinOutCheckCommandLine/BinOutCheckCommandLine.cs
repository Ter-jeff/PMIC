//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
//
// 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path.
// 2022-01-25  Steven Chen    #305	          Support command-line
// 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BinOutCheck;
using Library.Common;

namespace BinOutCheckCommandLine
{
    class BinOutCheckCommandLine
    {
        public static void Help()
        {
            string[] help = new string[]
            {
                "BinOut Check Tool",
                "Command-line arguments:",
                "-tpPath    test program file path(required)     -- IGXL test program file path(.igxl or .xlsm file)",
                "-output    output file path(required)         -- The output file path(.xlsx)",
                "-logPath   log folder path(optional)              -- Log folder path",
                                // 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path. add start
                "-datalogPath   datalog file path(optional)              -- Datalog file path(.txt)",
                // 2022-02-18  Steven Chen    #321	          In CommandLine mode,can't load setting files when program is started in other path. add end
                "Example: -tpPath testProgramPath   -output OutputPath  -logPath logPath",
                "Surround file/folder paths with double quotes if they contain spaces"
            };

            foreach(string s in help)
            {
                Console.WriteLine("{0}", s);
            }
        }
        public static void Main(string[] args)
        {
            try
            {
                // 2022-01-25  Steven Chen    #305	          Support command-line chg start
                //string tpPath = null;
                //string outputFilePath = null;
                //string logPath = null;
                // 2022-01-25  Steven Chen    #305	          Support command-line chg end
                string logFile = null;
                // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add start
                string datalogFile = null;
                GuiArgs guiArgs = new GuiArgs();
                // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut add end
                // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
                //if (!ParseArgList(args, ref tpPath, ref outputFilePath, ref logPath))
                if (!ParseArgList(args, guiArgs))
                    // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
                    return;
                // 2022-01-25  Steven Chen    #305	          Support command-line chg start
                //if (tpPath == null)
                //    throw new ArgumentException("Missing igxl test program file path, input '-?' to get help.");
                //if (outputFilePath == null)
                //    throw new ArgumentException("Missing output file path, input '-?' to get help.");
                //if (logPath != null)
                //    logFile = Path.Combine(logPath, "BinOutCheck.log");
                if (guiArgs.TestPlanPath == null)
                    throw new ArgumentException("Missing igxl test program file path, input '-?' to get help.");
                if (guiArgs.OutputFilePath == null)
                    throw new ArgumentException("Missing output file path, input '-?' to get help.");
                if (guiArgs.LogFilePath != null)
                    logFile = Path.Combine(guiArgs.LogFilePath, "BinOutCheck.log");
                // 2022-01-25  Steven Chen    #305	          Support command-line chg end

                Logger.PrintLog(logFile, "Start to do BinOut check......");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": Start to do BinOut check......");
                try
                {
                    // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
                    //new BinoutCheckMain(tpPath, outputFilePath).WorkFlow();
                    new BinoutCheckMain(guiArgs).WorkFlow();
                    // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end
                }
                catch(Exception e)
                {
                    Logger.PrintLog(logFile, "BinOut check error: " + e.ToString());
                    throw new Exception(e.ToString());
                }
                Logger.PrintLog(logFile, "BinOut check success!");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": BinOut check success!");
            }
            catch(Exception ex)
            {
                Console.Error.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": BinOut check error: {0}", ex.Message);
            }
        }

        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg start
        //private static bool ParseArgList(string[] args, ref string tpPath, ref string outputPath, ref string logPath)
        //{
        //    for (int i=0; i<args.Length; i++)
        //    {
        //        switch (args[i].ToLower())
        //        {
        //            case "-?":
        //                Help();
        //                if (args.Length == 1)
        //                    return false;
        //                break;
        //            case "-tppath":
        //                i += 1;
        //                tpPath = CheckTestProgramPath(i, args);
        //                break;
        //            case "-output":
        //                i += 1;
        //                outputPath = CheckOutputFilePath(i, args);
        //                break;
        //            case "-logpath":
        //                i += 1;
        //                logPath = CheckLogPath(i, args);
        //                break;
        //            default:
        //                break;
        //        }
        //    }
        //    return true;
        //}
        private static bool ParseArgList(string[] args, GuiArgs guiargs)
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
                        guiargs.TestPlanPath = CheckTestProgramPath(i, args);
                        break;
                    case "-datalogpath":
                        i += 1;
                        guiargs.DataLogPath = args[i];
                        break;
                    case "-output":
                        i += 1;
                        guiargs.OutputFilePath = CheckOutputFilePath(i, args);
                        break;
                    case "-logpath":
                        i += 1;
                        guiargs.LogFilePath = CheckLogPath(i, args);
                        break;
                    default:
                        break;
                }
            }
            return true;
        }
        // 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut chg end

        private static string CheckTestProgramPath(int index, string[] args)
        {
            string message;
            if(index >= args.Length)
            {
                message = string.Format("Missing igxl test program file path at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }
            if(!args[index].EndsWith(".igxl",StringComparison.OrdinalIgnoreCase) &&
                !args[index].EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                message = string.Format("The input test program file {0} is not an igxl test program file at argument {1}, input '-?' to get help.", args[index], index);
                throw new ArgumentException(message);
            }
            if (File.Exists(args[index]))
            {
                return args[index];
            }else
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
            }catch(Exception ex)
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
