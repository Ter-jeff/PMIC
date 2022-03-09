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
using System.IO;
using System.Linq;
using FWFrame.nWireDefinition.Enums;
using FWFrame.nWireDefinition.InputModel;
using FWFrame.Utils;

namespace FWFrame.nWireDefinition.FileGenerator
{
    public class ProtocolDefinitionFileGenerator : AbstractProtocolDefinitionFileGenerator
    {
        string idle_cycle_data = string.Empty;

        public ProtocolDefinitionFileGenerator(Context ctx) : base(ctx) { }

        public override void Generate()
        {
            // Show processing status
            reportStatus((int)ProcessPhaseEnum.GENERATE_DEFINITION_FILES, "Generate protocal definition file");

            Dictionary<string, object> model = new Dictionary<string, object>();

            string outputDir = guiInfo.GetParameter<string>("outputDir");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // LIST_FUNCTION_NAME
            // LIST_FUNCTION_TYPE
            // LIST_FUNCTION_DESCRIPTION
            string protocalName = guiInfo.GetParameter<string>("protocalName");
            Protocol protocal = protocals.Find(x => x.Name.Equals(protocalName));
            model.Add("LIST_FUNCTION_NAME", protocal.Ports.ConvertAll(x => x.Name));
            model.Add("LIST_FUNCTION_TYPE", protocal.Ports.ConvertAll(x => x.Type));
            model.Add("LIST_FUNCTION_DESCRIPTION", protocal.Ports.ConvertAll(x => x.Description));

            // LIST_FRAME_SEGMENT
            List<string> list_frame_segment = new List<string>();
            List<string> frameNames = guiInfo.GetParameter<List<string>>("frameNames");
            foreach (var frameName in frameNames)
            {
                list_frame_segment.AddRange(GenerateSingleFrame(frameName));
            }
            model.Add("LIST_FRAME_SEGMENT", list_frame_segment);

            // IDLE_CYCLE_DATA
            model.Add("IDLE_CYCLE_DATA", idle_cycle_data);

            string resultFile = Path.Combine(outputDir, string.Join("_", frameNames) + ".xml");
            ModelAndViewRender.RenderToFile(model, viewResolver.GetTemplateFilePath(TemplateTypeEnum.XMLFILE.ToString(), "Protocol", ctx), resultFile);
        }

        private List<string> GenerateSingleFrame(string frameName)
        {
            Dictionary<string, object> model = new Dictionary<string, object>();

            // FRAME_NAME
            model.Add("FRAME_NAME", frameName);

            // LIST_REPEAT_COUNT
            // LIST_CYCLE_DATA
            List<string> list_repeat_count = new List<string>();
            List<string> list_cycle_data = new List<string>();
            foreach (var item in dataInfoDic[frameName])
            {
                list_repeat_count.Add(item.Repeat.ToString());
                list_cycle_data.Add(string.Join(" ", item.Data));
            }
            model.Add("LIST_REPEAT_COUNT", list_repeat_count);
            model.Add("LIST_CYCLE_DATA", list_cycle_data);

            // For IDLE_CYCLE_DATA
            string lastCycle = list_cycle_data.Last();
            if (!string.IsNullOrEmpty(lastCycle))
            {
                idle_cycle_data = list_cycle_data.Last();
            }

            // LIST_FIELD_SEGMENT
            List<string> list_field_segment = new List<string>();
            foreach (Field field in fieldInfoDic[frameName])
            {
                list_field_segment.AddRange(GenerateSingleField(field));
            }
            model.Add("LIST_FIELD_SEGMENT", list_field_segment);

            return ModelAndViewRender.RenderToList(model, viewResolver.GetTemplateFilePath(TemplateTypeEnum.XMLFILE.ToString(), "Frame_Segment", ctx));
        }

        private List<string> GenerateSingleField(Field field)
        {
            Dictionary<string, object> model = new Dictionary<string, object>();

            // FIELD_NAME
            model.Add("FIELD_NAME", field.FieldName);

            // LIST_PORT_NAME
            model.Add("LIST_PORT_NAME", field.PortNames);

            // LIST_CYCLE_INDEX
            model.Add("LIST_CYCLE_INDEX", field.CycleIndice);

            return ModelAndViewRender.RenderToList(model, viewResolver.GetTemplateFilePath(TemplateTypeEnum.XMLFILE.ToString(), "Field_Segment", ctx));
        }
    }
}
