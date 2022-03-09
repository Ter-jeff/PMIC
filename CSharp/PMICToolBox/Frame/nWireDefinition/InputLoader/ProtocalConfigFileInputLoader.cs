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
// 2019 April 22 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 

using System.Collections.Generic;
using System.IO;
using System.Xml;
using FWFrame.InputLoader;
using FWFrame.nWireDefinition.InputModel;
using FWFrame.Properties;

namespace FWFrame.nWireDefinition.InputLoader
{
    public class ProtocolConfigFileInputLoader : IInputLoader
    {
        public void Load(Context ctx, StreamReader sr = null)
        {
            string configFilePath = Settings.Default.ProtocolConfigFilePath;

            // Begin to load file
            List<Protocol> protocals = new List<Protocol>();

            XmlDocument doc = new XmlDocument();
            if (sr != null)
                doc.Load(sr);
            else
                doc.Load(configFilePath);

            XmlElement root = doc.DocumentElement;

            // Get info for version
            XmlNodeList protocalNodes = root.SelectNodes("/Protocols/Protocol");
            foreach (XmlNode protocalNode in protocalNodes)
            {
                Protocol protocal = new Protocol();
                //protocal.Name = protocalNode.Name;
                protocal.Name = protocalNode.Attributes["name"].Value;

                XmlNodeList portNodes = protocalNode.ChildNodes;
                foreach (XmlNode portNode in portNodes)
                {
                    Port port = new Port();
                    port.Name = portNode.Attributes["name"].Value;
                    port.Group = portNode.Attributes["group"].Value;
                    port.Type = portNode.Attributes["type"].Value;
                    port.Description = portNode.Attributes["description"].Value;

                    protocal.Ports.Add(port);
                }

                protocals.Add(protocal);
            }

            ctx.Add("protocals", protocals);
        }
    }
}
