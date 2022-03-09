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
// 2021-12-08  Bruce          #258	          Add No Bin Out Check
//
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BinOutCheck
{
    [Serializable]
    public class NoBinOutResult
    {
        public string FlowSheet { get; set; }
        public string TestItem { get; set; }
        public string Comment { get; set; }
    }
}
