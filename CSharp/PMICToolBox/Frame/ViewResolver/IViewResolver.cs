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

namespace FWFrame.ViewResolver
{
    public interface IViewResolver
    {
        /// <summary>
        /// Get the path of template file.
        /// </summary>
        /// <param name="templateType">Type of the template.</param>
        /// <param name="name">The name.</param>
        /// <param name="ctx">The context.</param>
        /// <returns></returns>
        string GetTemplateFilePath(string templateType, string name, Context ctx);
    }
}
