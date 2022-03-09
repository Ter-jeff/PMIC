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

namespace FWFrame.Model
{
    public class IgxlTestInstancesEntry
    {
        public string Space = string.Empty;
        public string TestName = string.Empty;
        public string Type = string.Empty;
        public string Name = string.Empty;
        public string CalledAs = string.Empty;
        public string CategoryDc = string.Empty;
        public string SelectorDc = string.Empty;
        public string CategoryAc = string.Empty;
        public string SelectorAc = string.Empty;
        public string TimeSets = string.Empty;
        public string EdgeSets = string.Empty;
        public string PinLevels = string.Empty;
        public string MixedSignalTiming = string.Empty;
        public string Overlay = string.Empty;
        public string ArgList = string.Empty;
        public List<string> Args = new List<string>(new string[129]);
        public string Comment = string.Empty;

        public int ArgsCount = 0;

        public override string ToString()
        {
            List<string> items = new List<string>();

            items.Add(Space);
            items.Add(TestName);
            items.Add(Type);
            items.Add(Name);
            items.Add(CalledAs);
            items.Add(CategoryDc);
            items.Add(SelectorDc);
            items.Add(CategoryAc);
            items.Add(SelectorAc);
            items.Add(TimeSets);
            items.Add(EdgeSets);
            items.Add(PinLevels);
            items.Add(MixedSignalTiming);
            items.Add(Overlay);
            items.Add(ArgList);
            items.AddRange(Args);
            items.Add(Comment);

            return string.Join("\t", items);
        }
    }
}
