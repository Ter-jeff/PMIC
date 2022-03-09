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
// 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator
//------------------------------------------------------------------------------ 
using Library;
using Library.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatalogComparatorCommandLine
{
    class DatalogComparatorCommandLine
    {
        public static void Help()
        {
            string[] help = new string[]
            {
                "Datalog Comparator Tool",
                "Command-line arguments:",
                "-basePath    the base datalog path(required)     -- The base datalog file path(.txt)",
                "-compPath    the compare datalog path(required)     -- The compare datalog file path(.txt)",
                "-output    output folder path(required)         -- The output folder path",
                "-logPath   log folder path(optional)              -- Log folder path",
                "Example: -basePath baseDatalogPath   -compPath compareDatalogPath  -output OutputFolderPath",
                "Surround file/folder paths with double quotes if they contain spaces"
            };

            foreach (string s in help)
            {
                Console.WriteLine("{0}", s);
            }
        }

        static void Main(string[] args)
        {
            try
            {
                string logFile = null;
                if (!ParseArgList(args, ref logFile))
                    return;
                if (string.IsNullOrEmpty(CommonData.GetInstance().BaseTxtDatalogPath))
                    throw new ArgumentException("Missing the base datalog file path, input '-?' to get help.");
                if (string.IsNullOrEmpty(CommonData.GetInstance().CompareTxtDatalogPath))
                    throw new ArgumentException("Missing the compare datalog file path, input '-?' to get help.");
                if (string.IsNullOrEmpty(CommonData.GetInstance().OutputPath))
                    throw new ArgumentException("Missing output folder path, input '-?' to get help.");

                if (!string.IsNullOrEmpty(logFile))
                    logFile = Path.Combine(logFile, "DatalogComparator.log");

                Logger.PrintLog(logFile, "Start to do Datalog compare......");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": Start to do Datalog compare......");
                try
                {
                    MainLogic.Instance().MainFlowForCommandLine();
                }
                catch (Exception e)
                {
                    Logger.PrintLog(logFile, "Datalog compare error: " + e.ToString());
                    throw new Exception(e.ToString());
                }
                Logger.PrintLog(logFile, "Datalog compare success!");
                Console.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": Datalog compare success!");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(DateTime.Now.ToLocalTime().ToString() + ": Datalog compare error: {0}", ex.Message);
            }
        }

        private static bool ParseArgList(string[] args, ref string logFilePath)
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
                    case "-basepath":
                        i += 1;
                        CommonData.GetInstance().BaseTxtDatalogPath = CheckDatalogFilePath(i, args, "base");
                        break;
                    case "-comppath":
                        i += 1;
                        CommonData.GetInstance().CompareTxtDatalogPath = CheckDatalogFilePath(i, args, "compare");
                        break;
                    case "-output":
                        i += 1;
                        CommonData.GetInstance().OutputPath = CheckOutputFolderPath(i, args);
                        break;
                    case "-logpath":
                        i += 1;
                        logFilePath = CheckLogPath(i, args);
                        break;
                    default:
                        break;
                }
            }
            return true;
        }
        private static string CheckDatalogFilePath(int index, string[] args, string datalogType)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing {0} datalog file path at argument {1}, input '-?' to get help.", datalogType, index);
                throw new ArgumentException(message);
            }
            if (!args[index].EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
            {
                message = string.Format("The datalog file {0} is not an igxl datalog file at argument {1}, input '-?' to get help.", args[index], index);
                throw new ArgumentException(message);
            }
            if (File.Exists(args[index]))
            {
                return args[index];
            }
            else
            {
                message = string.Format("The {0} datalog file {1} not found, input '-?' to get help.", datalogType, args[index]);
                throw new ArgumentException(message);
            }
        }

        private static string CheckOutputFolderPath(int index, string[] args)
        {
            string message;
            if (index >= args.Length)
            {
                message = string.Format("Missing output folder path at argument {0}, input '-?' to get help.", index);
                throw new ArgumentException(message);
            }

            //if output folder is not exist, create the folder            
            try
            {
                if (!Directory.Exists(args[index]))
                    Directory.CreateDirectory(args[index]);
            }
            catch (Exception ex)
            {
                message = string.Format("Create output directory error, maybe the output path {0} is not valid: {1}", args[index], ex.Message);
                throw new ArgumentException(message);
            }

            return args[index];
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
    }
}
