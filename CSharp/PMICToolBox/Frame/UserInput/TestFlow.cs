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
// 2018 Feb 28 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 
using System.Collections.Generic;

namespace FWFrame.UserInput
{
    public class TestFlow
    {
        public string Index { get; set; }
        public string CommandName { get; set; }
        public string SerialNumber { get; set; }
        public string Active { get; set; }
        public string Target { get; set; }
        public double InjectFreq { get; set; }
        public double InjectPwr { get; set; }
        public string InjectPin { get; set; }
        public List<string> Params { get; set; }
        public List<string> Efuses { get; set; }

        public TestFlow()
        {
            Params = new List<string>();
            Efuses = new List<string>();
        }
    }
}
