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
    public class IgxlFlowTableEntry
    {
        public const string OPCODE_TEST = "Test";
        public const string OPCODE_USE_LIMIT = "Use-Limit";
        public const string OPCODE_BINTABLE = "Bintable";
        public const string OPCODE_RETURN = "return";

        public string Space = string.Empty;
        public string Label = string.Empty;
        public string Enable = string.Empty;
        public string Job = string.Empty;
        public string Part = string.Empty;
        public string Env = string.Empty;
        public string Opcode = string.Empty;
        public string Parameter = string.Empty;
        public string TName = string.Empty;
        public string TNum = string.Empty;
        public string LoLim = string.Empty;
        public string HiLim = string.Empty;
        public string Scale = string.Empty;
        public string Units = string.Empty;
        public string Format = string.Empty;
        public string PassBin = string.Empty;
        public string FailBin = string.Empty;
        public string PassSort = string.Empty;
        public string FailSort = string.Empty;
        public string Result = string.Empty;
        public string PassAction = string.Empty;
        public string FailAction = string.Empty;
        public string State = string.Empty;
        public string SpecifierGroup = string.Empty;
        public string SenseGroup = string.Empty;
        public string ConditionGroup = string.Empty;
        public string NameGroup = string.Empty;
        public string SenseDevice = string.Empty;
        public string ConditionDevice = string.Empty;
        public string NameDevice = string.Empty;
        public string Assume = string.Empty;
        public string Sites = string.Empty;
        public string ElapsedTimes = string.Empty;
        public string BackgroundType = string.Empty;
        public string Serialize = string.Empty;
        public string ResourceLock = string.Empty;
        public string FlowStepLocked = string.Empty;
        public string Comment = string.Empty;

        public string voltageType = string.Empty;

        public override string ToString()
        {
            List<string> items = new List<string>();

            items.Add(Space);
            items.Add(Label);
            items.Add(Enable);
            items.Add(Job);
            items.Add(Part);
            items.Add(Env);
            items.Add(Opcode);
            items.Add(Parameter);
            items.Add(TName);
            items.Add(TNum);
            items.Add(LoLim);
            items.Add(HiLim);
            items.Add(Scale);
            items.Add(Units);
            items.Add(Format);
            items.Add(PassBin);
            items.Add(FailBin);
            items.Add(PassSort);
            items.Add(FailSort);
            items.Add(Result);
            items.Add(PassAction);
            items.Add(FailAction);
            items.Add(State);
            items.Add(SpecifierGroup);
            items.Add(SenseGroup);
            items.Add(ConditionGroup);
            items.Add(NameGroup);
            items.Add(SenseDevice);
            items.Add(ConditionDevice);
            items.Add(NameDevice);
            items.Add(Assume);
            items.Add(Sites);
            items.Add(ElapsedTimes);
            items.Add(BackgroundType);
            items.Add(Serialize);
            items.Add(ResourceLock);
            items.Add(FlowStepLocked);
            items.Add(Comment);

            return string.Join("\t", items);
        }
    }
}
