//------------------------------------------------------------------------------
// Copyright (C) 2018 Teradyne, Inc. All rights reserved.
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
// 2019 Feb 28 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 

using System;
using System.IO;
using FWFrame.Controller;
using FWFrame.nWireDefinition.Enums;
using FWFrame.nWireDefinition.FileGeneratorMapper;

namespace FWFrame.nWireDefinition.Controllor
{
    public class GenerateProtocolDefinitionController : IController
    {
        public void Process(Context ctx, StreamReader sw = null)
        {
            GUIInfo guiInfo = ctx.Get<GUIInfo>("guiInfo");
            Action<int, string> reportStatus = guiInfo.GetParameter<Action<int, string>>("reportStatus");

            // Generate Files
            IFileGeneratorMapper fileGeneratorMapper = new XmlFileGeneratorMapper(sw);
            fileGeneratorMapper.GetGenerators(ctx).ForEach(x =>
            {
                x.Generate();
            });

            reportStatus((int)ProcessPhaseEnum.COMPLETE, "Mission complete");
        }
    }
}
