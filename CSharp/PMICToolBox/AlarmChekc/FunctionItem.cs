//------------------------------------------------------------------------------
// Copyright (C) 2021 Teradyne, Inc. All rights reserved.
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
// 2022-2-14  Terry Zhang     #312       Initial creation
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AlarmChekc
{
    public class FunctionItem
    {
        private string _functionname = string.Empty;

        public string FunctionName
        {
            get
            {
                return this._functionname;
            }
            set
            {
                this._functionname = value;
            }
        }

        private List<string> _lstfunctioncontent = null;

        public List<string> lstFunctionContent
        {
            get
            {
                return this._lstfunctioncontent;
            }
            set
            {
                this._lstfunctioncontent = value;
            }
        }

        public FunctionItem()
        {
            this._lstfunctioncontent = new List<string>();
        }
    }
}
