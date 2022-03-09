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
// Date        Name           Bug#            Notes
//
// 2019 March 28 Oliver Ou                    Initial creation
//
//------------------------------------------------------------------------------ 

using FWFrame;
using FWFrame.InputLoader;
using nWireDefinition.Enums;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace nWireDefinition.InputLoader
{
    public class ReverseCompileFileInputLoader : IInputLoader
    {
        public void Load(Context ctx)
        {
            GUIInfo guiInfo = ctx.Get<GUIInfo>("guiInfo");

            // Show processing status
            Action<int, string> reportStatus = guiInfo.GetParameter<Action<int, string>>("reportStatus");
            reportStatus((int)ProcessPhaseEnum.REVERSE_COMPILE_PAT_FILES, "Reverse compiling pattern files");

            List<string> patternFiles = guiInfo.GetParameter<List<string>>("patternFiles");

            string outputDir = Path.Combine(guiInfo.GetParameter<string>("outputDir"), "Temp");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Check whether IG-XL is installed
            if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("IGXLROOT")))
            {
                throw new FWFrameException("IG-XL is not installed");
            }

            // Reverse Compile one by one
            List<string> atpFiles = new List<string>();
            string commandLine = "/c aprc {0} -output {1} -force";
            foreach (var inputFile in patternFiles)
            {
                string outputFile = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(inputFile) + ".ATP");
                atpFiles.Add(outputFile);

                using (Process p = new Process())
                {
                    p.StartInfo = new ProcessStartInfo("cmd.exe")
                    {
                        Arguments = string.Format(commandLine, WrapFilePath(inputFile), WrapFilePath(outputFile)),
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                    };
                    p.Start();
                    p.WaitForExit();
                    if (p.ExitCode != 0)
                    {
                        throw new FWFrameException("Error occured during reverse compiling");
                    }
                }
            }

            ctx.Add("atpFiles", atpFiles);
        }

        private string WrapFilePath(string filePath)
        {
            return "\"" + filePath + "\"";
        }
    }
}
