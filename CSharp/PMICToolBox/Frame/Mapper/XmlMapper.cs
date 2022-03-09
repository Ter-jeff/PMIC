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
using FWFrame.Generator;
using FWFrame.Mapper.MappingConfig;
using FWFrame.UserInput;
using FWFrame.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace FWFrame.Mapper
{
    public class XmlMapper : IMapper
    {
        Mappings _mappingConfig = null;
        string _generatorNamespace = null;
        Assembly _assembly = null;

        /// <summary>
        /// key=Name of command
        /// value=Name of generator
        /// </summary>
        private Dictionary<string, string> mappingConfig = new Dictionary<string, string>();

        /// <summary>
        /// Cache for generator instances
        /// </summary>
        private Dictionary<string, List<IGenerator>> instanceCache = new Dictionary<string, List<IGenerator>>();

        public XmlMapper(string configFilePath)
        {
            _mappingConfig = XmlSer<Mappings>.LoadXml(configFilePath);
            _generatorNamespace = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Namespace.Split('.').First();
            _assembly = Assembly.GetExecutingAssembly();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<IGenerator> GetGenerators(TestFlow testFlow, Context ctx)
        {
            List<IGenerator> generators = new List<IGenerator>();

            // Check cache
            string funcName = Utilities.GetFuncNameByTestFlow(testFlow, ctx);
            if (instanceCache.Keys.Contains(funcName))
            {
                generators.AddRange(instanceCache[funcName]);
            }
            else
            {
                Mapping generatorMapping = _mappingConfig.Mapping.FirstOrDefault(x => x.BlockType.Equals(ctx.guiInfo.BlockType.ToString(), StringComparison.CurrentCultureIgnoreCase)
                                                                                      && x.Command.Equals(testFlow.CommandName, StringComparison.CurrentCultureIgnoreCase));
                if (generatorMapping != null)
                {
                    List<string> generatorNames = generatorMapping.Generator.ToList();
                    foreach (var generatorName in generatorNames)
                    {
                        Type type = Type.GetType(_generatorNamespace + ".Generator." + ctx.guiInfo.BlockType + "." + generatorName);
                        generators.Add((IGenerator)Activator.CreateInstance(type, new object[] { testFlow, ctx }));
                    }
                }
                else
                {
                    Type type = Type.GetType(_generatorNamespace + ".Generator." + ctx.guiInfo.BlockType + ".GeneralGenerator");
                    generators.Add((IGenerator)Activator.CreateInstance(type, new object[] { testFlow, ctx }));
                }

                instanceCache.Add(funcName, generators);
            }

            return generators;
        }
    }
}
