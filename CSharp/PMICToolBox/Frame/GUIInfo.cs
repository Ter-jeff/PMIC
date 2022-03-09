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
// 2019 Feb 28 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;

namespace FWFrame
{
    public class GUIInfo
    {
        private string chipType = string.Empty;
        private string command = string.Empty;
        private string workFolder = string.Empty;
        private string assemblyName = string.Empty;

        public string ChipType
        {
            get { return chipType; }
            set { chipType = value; }
        }

        public string Command
        {
            get { return command; }
            set { command = value; }
        }

        public string WorkFolder
        {
            get { return workFolder; }
            set { workFolder = value; }
        }

        public string AssemblyFile
        {
            get { return assemblyName; }
            set { assemblyName = value; }
        }

        // Parameters
        Dictionary<string, object> _parameters = new Dictionary<string, object>();

        // Get data from dictionary content
        public T GetParameter<T>(string key)
        {
            if (!_parameters.ContainsKey(key))
            {
                throw new FWFrameException("Can not find data with key=" + key + " in Context");
            }

            if (!(_parameters[key] is T))
            {
                throw new FWFrameException("Can not convert data with key=" + key + " to type=" + typeof(T) + " in Context");
            }

            return (T)_parameters[key];
        }

        // Add data to dictionary content
        public void AddParameter(string key, object value)
        {
            try
            {
                _parameters.Add(key, value);
            }
            catch (ArgumentException ex)
            {
                throw new FWFrameException("An element with the same key=" + key + "already exists in Context", ex);
            }
        }
    }
}
