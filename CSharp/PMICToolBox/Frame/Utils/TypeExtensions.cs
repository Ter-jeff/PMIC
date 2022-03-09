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

using System.Data;

namespace FWFrame.Utils
{
    public static class TypeExtensions
    {
        public static string GetString(this DataRow dataRow, string columnName)
        {
            if (dataRow[columnName] == null)
            {
                return string.Empty;
            }
            else
            {
                return dataRow[columnName].ToString().Trim();
            }
        }

        public static string GetInt(this DataRow dataRow, string columnName)
        {
            if (dataRow[columnName] == null)
            {
                return string.Empty;
            }
            else
            {
                return dataRow[columnName].ToString().Trim();
            }
        }

        public static bool IsNullOrBlank(this string s)
        {
            char[] WhiteSpaceChars = new char[] { (char)0x00, (char)0x01, (char)0x02, (char)0x03, (char)0x04, (char)0x05, 
                                                  (char)0x06, (char)0x07, (char)0x08, (char)0x09, (char)0x0a, (char)0x0b, 
                                                  (char)0x0c, (char)0x0d, (char)0x0e, (char)0x0f, (char)0x10, (char)0x11, 
                                                  (char)0x12, (char)0x13, (char)0x14, (char)0x15, (char)0x16, (char)0x17, 
                                                  (char)0x18, (char)0x19, (char)0x20, (char)0x1a, (char)0x1b, (char)0x1c, 
                                                  (char)0x1d, (char)0x1e, (char)0x1f, (char)0x7f, (char)0x85, (char)0x2028, 
                                                  (char)0x2029 };

            if (s == null || s.Trim(WhiteSpaceChars).Length == 0)
            {
                return true;
            }

            return false;
        }
    }
}
