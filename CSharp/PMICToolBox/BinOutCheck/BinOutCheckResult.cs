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
// 2021-12-15  Bruce          #261	          New Column For Fail Flag-Symbol Define Of Fail Flag
// 2021-12-08  Bruce          #257	          Change Title Name to Standard BinTable Mismatch
// 2021-11-22  Bruce          #230            Add Check by L-> Soft Bin Number & Fail Stop Table
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
    public class BinOutCheckResult
    {
        public string Name { get; set; }
        public string ItemList { get; set; }
        public string Op { get; set; }
        public string Sort { get; set; }
        public string Bin { get; set; }
        public string Result { get; set; }
        public string FlowName { get; set; }
        public string Fail_flag_and_Bin_define_in_same_flow { get; set; }
        public string Fail_flag_and_BinTable_Sequence_Check { get; set; }
        public string Redundant_Bin { get; set; }
        public string SwBinDuplicate { get; set; }
        // 2021-12-08  Bruce          #257	          Change Title Name to Standard BinTable Mismatch chg start
        //public string FailStop_BinName_Flag_Mismatch { get; set; } = "";
        public string Standard_BinTable_Mismatch { get; set; } = "";
        // 2021-12-08  Bruce          #257	          Change Title Name to Standard BinTable Mismatch chg end
        // 2021-12-15  Bruce          #261	          New Column For Fail Flag-Symbol Define Of Fail Flag add start
        public string Symbol_Define_Of_Fail_Flag { get; set; } = "";
        // 2021-12-15  Bruce          #261	          New Column For Fail Flag-Symbol Define Of Fail Flag add end
    }
}
