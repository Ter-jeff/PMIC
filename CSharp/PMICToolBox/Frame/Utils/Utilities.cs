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

using System;
using System.Reflection;

namespace FWFrame.Utils
{
    public class Utilities
    {
        public static bool IsInteger(string s)
        {
            int result = 0;
            return int.TryParse(s, out result);
        }

        public static bool IsNonnegativeInteger(string s)
        {
            int result = 0;
            if (int.TryParse(s, out result))
            {
                if (result >= 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool IsPositiveInteger(string s)
        {
            int result = 0;
            if (int.TryParse(s, out result))
            {
                if (result > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Removes special character including ".", "[", "]".
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static string RemoveSpecialChar(string str)
        {
            return str.Replace(".", "").Replace("[", "").Replace("]", "");
        }

        /// <summary>
        /// bothfix ???
        /// replace "0x" with "&h" and append the string with "&".
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static string ConvertHexFormatWithBothfix(string str)
        {
            return str.Replace("0x", "&h") + "&";
        }

        /// <summary>
        /// replace "0x" with "&h".
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static string ConvertHexFormatWithPrefix(string str)
        {
            return str.Replace("0x", "&h");
        }

        /// <summary>
        /// replace "&h" with "0x".
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static string ReverseConvertHexFormatWithPrefix(string str)
        {
            return str.Replace("&h", "0x");
        }

        /// <summary>
        /// Removes the hexadecimal format prefix("0x") from string.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static string RemoveHexFormatPrefix(string str)
        {
            return str.Replace("0x", "").Replace("&h", "");
        }

        public static T GetInstance<T>(string className, params object[] args)
        {
            Type type =Type.GetType(className);
            return (T)Activator.CreateInstance(type, args);
        }

        public static T GetInstance<T>(string assemblyPath, string className, params object[] args)
        {
            Assembly assembly = Assembly.LoadFile(assemblyPath);
            Type type = assembly.GetType(className);
            return (T)Activator.CreateInstance(type, args);
        }
    }
}
