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
// 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet.
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BinOutCheck
{
    public class BinOutConfig
    {
        public string Name { get; set; }
        public string ItemList { get; set; }
        public string Op { get; set; }
        public string Sort { get; set; }
        public string Bin { get; set; }
        public string Result { get; set; }
        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add start
        public bool IsUsed { get; set; } = false;
        // 2021-12-17  Bruce          #266	          Standard Bin：List missing Standard Bin in the BinOut sheet. add end
    }
}
