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

using FWFrame.Controller;
using FWFrame.DispatherConfig;
using FWFrame.InputLoader;
using FWFrame.Interceptor;
using FWFrame.Utils;
using FWFrame.ViewResolver;
using System;
using System.IO;
using System.Linq;

namespace FWFrame
{
    public class Dispatcher
    {
        Context _ctx = new Context();

        public Dispatcher(object guiInfo)
        {
            _ctx.Add("guiInfo", guiInfo);

            // For session
            _ctx.Add("operateTime", DateTime.Now);
        }

        public void Dispatch(StreamReader srDispather = null, StreamReader srProtocol = null, StreamReader srFileGeneratorMapperConfigure = null)
        {
            GUIInfo guiInfo = _ctx.Get<GUIInfo>("guiInfo");

            string dispatherConfigureFilePath = Path.Combine(guiInfo.WorkFolder, "Configure", "DispatherConfigure.xml");
            DispatherConfigureLoader _dispatherConfigureLoader = srDispather == null ?
                new DispatherConfigureLoader(dispatherConfigureFilePath) :
                new DispatherConfigureLoader(srDispather);

            // Get config for chip
            string chipType = guiInfo.ChipType;
            ChipConfigure chipConfigure = _dispatherConfigureLoader.GetChipConfigure(chipType);
            if (chipConfigure == null)
            {
                throw new FWFrameException("Can not found configure for given chip type [" + chipType + "]");
            }

            // Get config for command
            string command = guiInfo.Command;
            ControllerMapping controllerMapping = chipConfigure.ControllerMappings.FirstOrDefault(x => x.Command.Equals(command));
            if (controllerMapping == null)
            {
                throw new FWFrameException("Can not found configure for given command [" + command + "]");
            }

            // Get ViewResolver
            string assemblyFile = Path.Combine(Directory.GetCurrentDirectory(), guiInfo.AssemblyFile);
            _ctx.Add("viewResolver", Utilities.GetInstance<IViewResolver>(controllerMapping.ViewResolver));

            // Get InputLoader and Run
            if (controllerMapping.InputLoaders != null)
            {
                controllerMapping.InputLoaders.ToList().ForEach(x =>
                {
                    IInputLoader inputLoader = Utilities.GetInstance<IInputLoader>(x);
                    inputLoader.Load(_ctx, srProtocol);
                });
            }

            // Get Interceptors and Run
            if (controllerMapping.Interceptors != null)
            {
                controllerMapping.Interceptors.ToList().ForEach(x =>
                {
                    IInterceptor interceptor = Utilities.GetInstance<IInterceptor>(x);
                    interceptor.Intercept(_ctx);
                });
            }

            // Get Controller and Run
            IController controllor = Utilities.GetInstance<IController>(controllerMapping.Controller);
            controllor.Process(_ctx, srFileGeneratorMapperConfigure);
        }


    }
}
