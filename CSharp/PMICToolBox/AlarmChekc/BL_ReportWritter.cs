//------------------------------------------------------------------------------
// Copyright (C) 2021 Teradyne, Inc. All rights reserved.
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
// Date        Name           Task#           Notes
// 2022-2-14  Terry Zhang     #312       Initial creation
//------------------------------------------------------------------------------ 

using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AlarmChekc
{
    public class BL_ReportWritter
    {
        private List<ReportItem> _lstreportitems = null;
        private string _outputfilename = string.Empty;

        public BL_ReportWritter(List<ReportItem> p_lstReportItem,string p_strOutputFileName)
        {
            this._lstreportitems = p_lstReportItem;
            this._outputfilename = p_strOutputFileName;
        }

        public void ExportToFile()
        {
            //export to file
            string l_strOutputFileName = Path.Combine(this._outputfilename, "AlarmCheckReport.xlsx");

            try
            {
                if (File.Exists(l_strOutputFileName))
                {
                    File.Delete(l_strOutputFileName);
                }
            }
            catch(Exception ex)
            {
                throw new Exception("Please close the report file first if it is opened!");
            }

            try
            {
                FileInfo l_FileInfo = new FileInfo(l_strOutputFileName);

                var ep = new ExcelPackage(new FileInfo(l_strOutputFileName));

                ExcelWorksheet l_WorkSheet = ep.Workbook.Worksheets.Add("AlarmCheck");
                ep.Workbook.Worksheets["AlarmCheck"].Cells.LoadFromCollection(this._lstreportitems, true);


                for (var col = 1; col <= l_WorkSheet.Dimension.End.Column; col++)
                {
                    l_WorkSheet.Cells[1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    l_WorkSheet.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    l_WorkSheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.DarkOliveGreen);
                    l_WorkSheet.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                    l_WorkSheet.Cells[1, col].Style.Font.Size = 12;
                    l_WorkSheet.Cells[1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }



                ep.Save();
                ep.Dispose();
            }
            catch
            {
                throw new Exception("export data to excel file failed!");
            }
        }
    }
}
