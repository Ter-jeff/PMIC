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
// 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item)
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.ErrorHandler
{
    class Result
    {
        public string FileName { get; set; }
        public string FuncName { get; set; }

        //public bool ExistOnErrorGoTo { get; set; }
        public bool Modified { get; set; }

        //public bool ExistFuncName { get; set; }
        //public bool ExistErrHandler { get; set; }

        public bool StopKeyword { get; set; }
        public bool MsgBoxKeyword { get; set; }
    }

    // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg start
    //public class OptionExplicitResult
    //{
    //    public string FileName { get; set; }
    //    public bool ExistOptionExplicit { get; set; }
    //}
    public class OptionExplicitResult
    {
        public string FileName { get; set; }
    }
    // 2022-02-10  Steven Chen    #309	          Simplify Output Option Explicit sheet(Change Sheet Name, Remove Result, Remove Pass Item) chg end
}
