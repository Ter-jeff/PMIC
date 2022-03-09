using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.Config.ProjectConfig;

namespace PmicAutogen.Inputs.Setting.BinNumber.Reader
{
    public class BinNumberMbistReader
    {
        private const string Category = "category";
        private const string Number = "number";
        private const string SubDig = "Sub dig";
        private const string Column = "Column";
        //private const string Keyword = "Keyword";

        public List<SoftBinMbist> ReadSheet(ExcelWorksheet worksheet)
        {
            List<SoftBinMbist> detailList = new List<SoftBinMbist>();
            try
            {
                for (int i = 3; i <= worksheet.Dimension.End.Row; i++)
                {
                    var range = worksheet.MergedCells[i, 1];
                    int start = i;
                    int end;
                    if (range == null)
                    {
                        end = i;
                    }
                    else
                    {
                        end = new ExcelAddress(range).End.Row;
                        i = end;
                    }

                    var oneCat = ReadOneCat(worksheet, start, end);
                    oneCat.Category = ReadConfigCell(worksheet, start, 1);
                    oneCat.SheetName = worksheet.Name;
                    if (!oneCat.Category.Equals(""))
                    {
                        detailList.Add(oneCat);
                    }
                }
            }
            catch (Exception e)
            {

                throw new Exception(e.Message);
            }
            return detailList;
        }

        private SoftBinMbist ReadOneCat(ExcelWorksheet worksheet, int startRow, int endRow)
        {
            SoftBinMbist categoryDetail = new SoftBinMbist();
            for (int i = 2; i <= worksheet.Dimension.End.Column; i++)
            {
                var range = worksheet.MergedCells[1, i];
                if (range == null)
                    break;
                int start = i;
                int end = new ExcelAddress(range).End.Column;
                i = end;
                var digitalDef = ReadOneDigital(worksheet, startRow, endRow, start, end);
                if (digitalDef != null)
                {
                    categoryDetail.SoftDetails.Add(digitalDef);
                }
            }
            return categoryDetail;
        }

        private List<BinNumberMbistRow> ReadOneDigital(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn)
        {
            List<BinNumberMbistRow> binNumberMbistRows = new List<BinNumberMbistRow>();

            for (int i = startRow; i <= endRow; i++)
            {
                BinNumberMbistRow numDef = new BinNumberMbistRow();
                for (int j = startColumn; j <= endColumn; j++)
                {
                    var columnName = EpplusOperation.GetCellValue(worksheet, 2, j);
                    if (columnName.Equals(Category, StringComparison.OrdinalIgnoreCase))
                    {
                        numDef.Category = ReadConfigCell(worksheet, i, j);
                    }
                    else if (columnName.Equals(Number, StringComparison.OrdinalIgnoreCase))
                    {
                        numDef.Number = ReadConfigCell(worksheet, i, j);
                    }
                    else if (columnName.Equals(Column, StringComparison.OrdinalIgnoreCase))
                    {
                        var name = EpplusOperation.GetCellValue(worksheet, i, j);
                        var condition = EpplusOperation.GetCellValue(worksheet, i, j + 1);
                        if (!name.Equals(""))
                        {
                            numDef.AddCondition(name, condition);
                        }
                        j++;
                    }
                    else if (columnName.Equals(SubDig, StringComparison.OrdinalIgnoreCase))
                    {
                        string subDig = EpplusOperation.GetCellValue(worksheet, i, j);
                        string[] digList = subDig.Split(',');
                        foreach (string s in digList)
                        {
                            numDef.SubDig.Add(s);
                        }
                    }
                }
                if (!numDef.Category.Equals("") && !binNumberMbistRows.Exists(p => p.Category.Equals(numDef.Category)))
                {
                    binNumberMbistRows.Add(numDef);
                }
            }

            return binNumberMbistRows;
        }

        private string ReadConfigCell(ExcelWorksheet worksheet, int row, int column)
        {
            string result = EpplusOperation.GetCellValue(worksheet, row, column).Trim();
            result = ProjectConfigSingleton.Instance().ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, result);
            return result;
        }
    }
}