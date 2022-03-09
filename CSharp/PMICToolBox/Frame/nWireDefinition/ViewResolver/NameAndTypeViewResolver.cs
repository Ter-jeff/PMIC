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

using System.IO;
using FWFrame.Properties;
using FWFrame.ViewResolver;

namespace FWFrame.nWireDefinition.ViewResolver
{
    public class NameAndTypeViewResolver : IViewResolver
    {
        private string _templateDirectoryPath = Settings.Default.TemplateDirectory;

        public string GetTemplateFilePath(string templateType, string name, Context ctx)
        {
            string filePath = string.Empty;

            switch(templateType)
            {
                case "XMLFILE":
                    filePath = Path.Combine(_templateDirectoryPath, templateType, name) + ".tmp";
                    break;
                default:
                    throw new FWFrameException("Wrong TemplateTypeEnum paremeter");
            }

            if (!File.Exists(filePath))
            {
                throw new FWFrameException("Template file does not exist : " + filePath);
            }

            return filePath;
        }
    }
}
