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
//using FWFrame.Enums;

using System;
using System.IO;
using System.Xml.Serialization;

namespace FWFrame.Utils
{

    /// <summary>
    /// Xml serialization
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XmlSer<T>
    {
        /// <summary>
        /// serialize to xml file
        /// </summary>
        /// <param name="output"></param>
        /// <param name="sysData"></param>
        /// <returns></returns>
        public static void SaveXML(string output, T sysData)
        {
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(T));
                XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
                ns.Add(string.Empty, string.Empty);
                StreamWriter sw = new StreamWriter(output);
                xs.Serialize(sw, sysData, ns);
                sw.Close();
            }
            catch (Exception ex)
            {
                throw new FWFrameException(ex.Message, ex);
            }
        }

        /// <summary>
        /// deserialize from xml file
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        /// 
        public static T LoadXml(string input)
        {
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(T));
                StreamReader sr = new StreamReader(input);
                T sysData = (T)xs.Deserialize(sr);
                sr.Close();
                return sysData;
            }
            catch (Exception ex)
            {
                throw new FWFrameException(ex.Message, ex);
            }
        }

        public static T LoadXml(StreamReader sr)
        {
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(T));
                T sysData = (T)xs.Deserialize(sr);
                sr.Close();
                return sysData;
            }
            catch (Exception ex)
            {
                throw new FWFrameException(ex.Message, ex);
            }
        }
    }
}
