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

using System.Collections.Generic;
using FWFrame.nWireDefinition.FileGenerator;

namespace FWFrame.nWireDefinition.FileGeneratorMapper
{
    public interface IFileGeneratorMapper
    {
        List<IFileGenerator> GetGenerators(Context ctx);
    }
}
