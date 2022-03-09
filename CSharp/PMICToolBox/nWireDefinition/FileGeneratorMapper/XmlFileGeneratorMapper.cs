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

using FWFrame;
using FWFrame.Utils;
using nWireDefinition.FileGenerator;
using nWireDefinition.FileGeneratorMapper.FileGeneratorMapperConfig;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace nWireDefinition.FileGeneratorMapper
{
    public class XmlFileGeneratorMapper : IFileGeneratorMapper
    {
        Mappings _config = null;

        public XmlFileGeneratorMapper()
        {
            _config = XmlSer<Mappings>.LoadXml(Properties.Settings.Default.FileGeneratorMapperConfigFilePath);
        }

        /// <summary>
        /// get file generators from context
        /// </summary>
        /// <returns></returns>
        public List<IFileGenerator> GetGenerators(Context ctx)
        {
            GUIInfo guiInfo = ctx.Get<GUIInfo>("guiInfo");
            string assemblyFile = Path.Combine(guiInfo.WorkFolder, guiInfo.AssemblyFile);

            List<IFileGenerator> fileGenerators = new List<IFileGenerator>();
            _config.Mapping.First(x => x.GenerationType.Equals(guiInfo.Command)).Generator.ToList()
                .ForEach(x =>
                {
                    fileGenerators.Add(Utilities.GetInstance<IFileGenerator>(assemblyFile, x, ctx));
                });

            return fileGenerators;
        }
    }
}
