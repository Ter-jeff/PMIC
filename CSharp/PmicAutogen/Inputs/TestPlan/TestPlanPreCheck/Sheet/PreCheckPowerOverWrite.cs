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
// 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore. 
//
//------------------------------------------------------------------------------ 

using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckPowerOverWrite : PreCheckBase
    {
        #region Constructor
        public PreCheckPowerOverWrite(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
        {
            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add Start
            IgnoreBlankSheet = true;
            // 2021-06-21  Bruce Qian     #87             T-auotgen , Power Override sheet , if user not put any information then just ignore.  Add End
        }

        #endregion

        #region Member Function
        protected override bool CheckBusiness()
        {
            return true;
        }
        #endregion
    }
}