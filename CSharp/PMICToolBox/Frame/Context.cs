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

using System.Collections.Generic;

namespace FWFrame
{
    public class Context
    {
        Dictionary<string, object> _content = new Dictionary<string, object>();

        // get data from dictionary content
        public T Get<T>(string key)
        {
            if (!_content.ContainsKey(key))
            {
                throw new FWFrameException("Can not find data with key=" + key + " in Context");
            }

            if (!(_content[key] is T))
            {
                throw new FWFrameException("Can not convert data with key=" + key + " to type=" + typeof(T) + " in Context");
            }

            return (T)_content[key];
        }

        // add data to dictionary content
        public void Add(string key, object value)
        {
            if (!_content.ContainsKey(key))
            {
                _content.Add(key, value);
            }
            else
            {
                throw new FWFrameException("An element with the same key=" + key + "already exists in Context");
            }
        }

        // Put data to dictionary content
        public void Put(string key, object value)
        {
            if (!_content.ContainsKey(key))
            {
                _content.Add(key, value);
            }
            else
            {
                _content.Remove(key);
                _content.Add(key, value);
            }
        }

        // Determine whether given key is contained
        public bool ContainsKey(string key)
        {
            return _content.ContainsKey(key);
        }
    }
}
