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

using FWFrame.DispatherConfig;
using FWFrame.Utils;
using System.IO;
using System.Linq;

namespace FWFrame
{
    public class DispatherConfigureLoader
    {
        DispatherConfigure _config = null;

        public DispatherConfigureLoader(string configFilePath)
        {
            _config = XmlSer<DispatherConfigure>.LoadXml(configFilePath);
        }

        public DispatherConfigureLoader(StreamReader sr)
        {
            _config = XmlSer<DispatherConfigure>.LoadXml(sr);
        }

        /// <summary>
        /// get chip configure data
        /// </summary>
        /// <returns></returns>
        public ChipConfigure GetChipConfigure(string chipType)
        {
            return _config.ChipConfigure.FirstOrDefault(x => x.ChipType.Equals(chipType));
        }
    }
}
