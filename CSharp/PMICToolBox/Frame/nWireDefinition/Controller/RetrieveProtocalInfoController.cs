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
using System.Collections.Generic;
using System.IO;
using FWFrame.Controller;
using FWFrame.nWireDefinition.InputModel;

namespace FWFrame.nWireDefinition.Controller
{
    public class RetrieveProtocolInfoController : IController
    {
        public void Process(Context ctx, StreamReader sw = null)
        {
            GUIInfo guiInfo = ctx.Get<GUIInfo>("guiInfo");

            // Return data to front
            Dictionary<string, List<Tuple<string, string>>> protocalInfo = new Dictionary<string, List<Tuple<string, string>>>();
            List<Protocol> protocals = ctx.Get<List<Protocol>>("protocals");
            foreach (Protocol protocal in protocals)
            {
                protocalInfo[protocal.Name] = protocal.Ports.ConvertAll(x => new Tuple<string, string>(x.Name, x.Type));
            }
            guiInfo.AddParameter("protocalInfo", protocalInfo);
        }
    }
}
