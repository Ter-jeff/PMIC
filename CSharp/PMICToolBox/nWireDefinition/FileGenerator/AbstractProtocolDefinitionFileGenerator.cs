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
using FWFrame.ViewResolver;
using nWireDefinition.InputModel;
using System;
using System.Collections.Generic;

namespace nWireDefinition.FileGenerator
{
    public abstract class AbstractProtocolDefinitionFileGenerator : IFileGenerator
    {
        protected Context ctx = null;
        protected GUIInfo guiInfo = null;
        protected IViewResolver viewResolver = null;
        protected Action<int, string> reportStatus = null;

        protected Dictionary<string, List<Cycle>> dataInfoDic = null;
        protected Dictionary<string, List<Field>> fieldInfoDic = null;
        protected List<Protocol> protocals = null;

        public AbstractProtocolDefinitionFileGenerator(Context ctx)
        {
            this.ctx = ctx;

            this.guiInfo = ctx.Get<GUIInfo>("guiInfo");
            this.viewResolver = ctx.Get<IViewResolver>("viewResolver");
            this.reportStatus = guiInfo.GetParameter<Action<int, string>>("reportStatus");

            this.dataInfoDic = ctx.Get<Dictionary<string, List<Cycle>>>("dataInfoDic");
            this.fieldInfoDic = ctx.Get<Dictionary<string, List<Field>>>("fieldInfoDic");
            this.protocals = ctx.Get<List<Protocol>>("protocals");
        }

        public abstract void Generate();
    }
}
