using PmicAutomation.Utility.OTPRegisterMap.Base;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using DT = System.Data;
using System.Text.RegularExpressions;
using OfficeOpenXml.Style;
using System.Drawing;
using PmicAutomation.Utility.OTPRegisterMap.Input;
using Library.Function.ErrorReport;

namespace PmicAutomation.Utility.OTPRegisterMap.Output
{
    public class WriterOtpRegisterMap
    {
        #region public variable
        //private readonly List<OtpRegisterItem> _otpRegItems;
        //private readonly List<string> _headers;
        //private readonly List<Tuple<int, string>> _oriHeaders;
        //private readonly List<string> _otpVersionList;

        private OtpFileReader _yamlFile = null;
        private List<OtpFileReader> _otpFileList = new List<OtpFileReader>();
        private DT::DataTable _regMapFile = null;
        private string regMapFileName = "";
        #endregion

        public WriterOtpRegisterMap(List<string> headers, List<Tuple<int, string>> oriHeaders, List<OtpRegisterItem> regItems, List<string> versionList)
        {
            //_otpRegItems = regItems;
            //_headers = headers;
            //_oriHeaders = oriHeaders;
            //_otpVersionList = versionList;
        }

        public WriterOtpRegisterMap(OtpFileReader yamlFile, List<OtpFileReader> otpFiles, DT::DataTable regMapFile)
        {
            _yamlFile = yamlFile;
            _otpFileList = otpFiles;
            _regMapFile = regMapFile;
        }

        #region OutPutEFuseBitDef

        public void OutPutResult(string outPutFolder)
        {
            string time = DateTime.Now.ToFileTime().ToString();
            if (_yamlFile != null)
            {
                if(_otpFileList.Count == 0)
                {
                    //Compare yaml and registerMap                    
                    string lStrDiffExcelReport = Path.Combine(outPutFolder, "DiffReport_" + Path.GetFileNameWithoutExtension(_regMapFile.TableName) + "_Yaml" + time + ".xlsx");
                    if (File.Exists(lStrDiffExcelReport)) File.Delete(lStrDiffExcelReport);
                    WriteRegisterMapYamlDiffReport(lStrDiffExcelReport, _regMapFile);
                }
                else
                {                    
                    if(_regMapFile == null)
                    {
                        foreach (OtpFileReader otpFile in _otpFileList)
                        {
                            _yamlFile.MergeOtpToYaml(otpFile);
                        }
                        GenRegisterMapByYamlAndOtps(outPutFolder, "OTP_Register_Map", true);
                    }
                    else
                    {
                        string lStrFileNameTxt = Path.Combine(outPutFolder, Path.GetFileNameWithoutExtension(_regMapFile.TableName) + "_" + time + ".txt");
                        string lStrDiffExcelReport = Path.Combine(outPutFolder, "DiffReport_" + Path.GetFileNameWithoutExtension(_regMapFile.TableName) + "_Yaml" + time + ".xlsx");
                        if (File.Exists(lStrFileNameTxt)) File.Delete(lStrFileNameTxt);
                        if (File.Exists(lStrDiffExcelReport)) File.Delete(lStrDiffExcelReport);

                        DT::DataTable newRegisterTable = GenRegisterMapByRegisterMapAndOtps(lStrFileNameTxt);
                        WriteRegisterMapYamlDiffReport(lStrDiffExcelReport, newRegisterTable);
                    }

                }
            }
            else
            {
                //Add OTP content to RegisterMap file
                string lStrFileNameTxt = Path.Combine(outPutFolder, Path.GetFileNameWithoutExtension(_regMapFile.TableName) + "_" + time + ".txt");                
                if (File.Exists(lStrFileNameTxt)) File.Delete(lStrFileNameTxt);
                GenRegisterMapByRegisterMapAndOtps(lStrFileNameTxt);                
            }
        }
        public void GenRegisterMapByYamlAndOtps(string path, string fileName, bool outputTxt)
        {
            string lStrFileNameXls = Path.Combine(path, fileName + ".xlsx");
            string lStrFileNameCsv = Path.Combine(path, fileName + ".csv");
            string lStrFileNameTxt = Path.Combine(path, fileName + ".txt");

            FileInfo lCsvFile = new FileInfo(lStrFileNameCsv);
            FileInfo lTxtFile = new FileInfo(lStrFileNameTxt);

            Directory.CreateDirectory(path);

            if (File.Exists(lStrFileNameXls))
            {
                File.Delete(lStrFileNameXls);
            }

            if (File.Exists(lStrFileNameCsv))
            {
                File.Delete(lStrFileNameCsv);
            }

            if (File.Exists(lStrFileNameTxt))
            {
                File.Delete(lStrFileNameTxt);
            }

            if (_yamlFile.OtpRows == null) return;
            var excel = new ExcelPackage(new FileInfo(lStrFileNameXls));
            var workbook = excel.Workbook;

            ExcelWorksheet lOtpRegMapSheet = workbook.Worksheets.Add(fileName);
            //foreach (var item in _oriHeaders)
            //{
            //    lOtpRegMapSheet.Cells[1, item.Item1].Value = item.Item2;
            //}
            lOtpRegMapSheet.Cells["A2"].Value = string.Empty; // keep blank row
            int rowIndex = 3;
            for (int i = 0; i < _yamlFile.Headers.Count; i++)
            {
                lOtpRegMapSheet.Cells[rowIndex, i + 1].Value = _yamlFile.Headers[i];
            }
            lOtpRegMapSheet.Cells[rowIndex, _yamlFile.Headers.Count + 1].Value = "END";
            rowIndex++;
            foreach (var otpItem in _yamlFile.OtpRows)
            {
                int index = 1;
                if (otpItem.OtpRegisterName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegisterName.ToUpper();
                if (otpItem.Name != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Name.ToUpper();
                if (otpItem.InstName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstName.ToUpper();
                if (otpItem.InstBase != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstBase;
                if (otpItem.RegName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegName.ToUpper();
                if (otpItem.RegOfs != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegOfs;
                if (otpItem.OtpOwner != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpOwner;
                if (otpItem.DefaultValue != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultValue.ToUpper();
                if (otpItem.Bw != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Bw.ToUpper();
                if (otpItem.Idx != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Idx.ToUpper();
                if (otpItem.Offset != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Offset.ToUpper();
                if (otpItem.OtpB0 != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpB0.ToUpper();
                if (otpItem.OtpA0 != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpA0.ToUpper();
                if (otpItem.OtpRegAdd != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegAdd.ToUpper();
                if (otpItem.OtpRegOfs != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegOfs.ToUpper();
                if (otpItem.DefaultOrReal != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultOrReal;
                lOtpRegMapSheet.Cells[rowIndex, index++].Value = "";
                lOtpRegMapSheet.Cells[rowIndex, index++].Value = "";
                lOtpRegMapSheet.Cells[rowIndex, index++].Value = "";
                lOtpRegMapSheet.Cells[rowIndex, index++].Value = "0";
                //int i = 20;
                foreach (var extraItem in otpItem.OtpExtra)
                {
                    lOtpRegMapSheet.Cells[rowIndex, index++].Value = extraItem;
                    //i++;
                }
                rowIndex++;
            }
            lOtpRegMapSheet.Cells[rowIndex, 1].Value = "END";

            try
            {
                excel.Save();
                excel.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Write error report failed. " + e.Message);
            }

            Application lExcelApp = new Application();
            lExcelApp.DisplayAlerts = false;
            Workbooks lWks = lExcelApp.Workbooks;
            Workbook lWk = null;
            try
            {
                lWk = lWks.Open(lStrFileNameXls);
                lWk.SaveAs(lCsvFile, XlFileFormat.xlCSV);
                if (outputTxt)
                {
                    lWk.SaveAs(lTxtFile, XlFileFormat.xlCurrentPlatformText);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Write error report failed. " + e.Message);
            }
            finally
            {
                if (lWk != null)
                {
                    lWk.Close(false);
                    Marshal.ReleaseComObject(lWk);
                }
                lWk = null;
                Marshal.ReleaseComObject(lWks);
                lWks = null;
                lExcelApp.Quit();
                Marshal.ReleaseComObject(lExcelApp);
                lExcelApp = null;
                GC.Collect();
            }
        }

        public DT::DataTable GenRegisterMapByRegisterMapAndOtps(string outputRegisterFilePath)
        {
            // all index value are stored in memory sight
            int headerRowIndex = -1;
            int insertRowIndex = -1;
            int insertColIndex = -1;
            int registerNameColIndex = -1;

            FetchDataTableInfo(_regMapFile, ref headerRowIndex, ref insertRowIndex, ref insertColIndex, ref registerNameColIndex, "OTP_REGISTER_NAME");
            DT::DataTable targetDT = _regMapFile.Copy();

            // header setting
            int tmpIndex = insertColIndex;
            for (int i = 0; i < _otpFileList.Count; ++i)
            {
                string colName = "Column" + (_regMapFile.Columns.Count + i + 1).ToString();
                targetDT.Columns.Add(new DT.DataColumn(colName));
                targetDT.Rows[headerRowIndex][tmpIndex++] = _otpFileList[i].GetVersionFromOtpFileName();
            }
            targetDT.Rows[headerRowIndex][tmpIndex] = "END";

            // keep original header
            tmpIndex = insertColIndex;
            //foreach (var item in _oriHeaders)
            //{
            //    targetDT.Rows[0][tmpIndex++] = item.Item2;
            //}

            // content setting

            string registerName = "";
            OtpRegisterItem otpRow = null;
            for (int row = insertRowIndex; row < targetDT.Rows.Count - 1; row++)
            {
                registerName = targetDT.Rows[row][registerNameColIndex].ToString().Trim();
                if (string.IsNullOrEmpty(registerName) || registerName.Equals("END"))
                    continue;
                tmpIndex = insertColIndex;
                foreach (OtpFileReader otpFile in _otpFileList)
                {
                    otpRow = otpFile.OtpRows.Find(s => s.OtpRegisterName.Equals(registerName, StringComparison.OrdinalIgnoreCase));
                    if(otpRow != null)
                    {
                        targetDT.Rows[row][tmpIndex] = otpRow.DefaultValue;
                    }
                    tmpIndex++;
                }
            }

            //foreach (var otpItem in _yamlFile.OtpRows)
            //{
            //    tmpIndex = insertColIndex;
            //    foreach (var extraItem in otpItem.OtpExtra)
            //    {
            //        if(insertRowIndex < targetDT.Rows.Count - 1)
            //            targetDT.Rows[insertRowIndex][tmpIndex++] = extraItem;
            //    }
            //    ++insertRowIndex;
            //}

            DumpDataTable2File(targetDT, outputRegisterFilePath);

            return targetDT;
            
        }
        #endregion
        
        private void WriteRegisterMapYamlDiffReport(string reportFile, DT::DataTable regMapDT)
        {
            FileInfo file;
            ExcelPackage epp = null;
            ExcelWorkbook workbook = null;
            try
            {
                file = new FileInfo(reportFile);
                epp = new ExcelPackage();
                workbook = epp.Workbook;
                ExcelWorksheet lOtpRegMapSheet = workbook.Worksheets.Add("RegisterMap_Yaml_Diff");
                OtpRegisterItem otpItem;
                int headRow = -1, index;
                string registerMapValue, otpRegisterName;
                List<string> registerNamelstInRegMap = new List<string>();                

                for (int i = 0; i < regMapDT.Rows.Count; i++)
                {
                    for (int j = 0; j < regMapDT.Columns.Count; j++)
                    {
                        registerMapValue = regMapDT.Rows[i][j].ToString();
                        lOtpRegMapSheet.Cells[i + 1, j + 1].Value = registerMapValue;
                        if (registerMapValue.Equals("OTP_REGISTER_NAME", StringComparison.OrdinalIgnoreCase) && headRow == -1)
                            headRow = i;                       
                    }
                    //Compare register map and yaml value
                    if (headRow > -1 && i > headRow && i < regMapDT.Rows.Count - 1)
                    {
                        otpRegisterName = lOtpRegMapSheet.Cells[i + 1, 1].Value.ToString();
                        otpItem = _yamlFile.OtpRows.Find(s => GetString(s.OtpRegisterName).Equals(otpRegisterName, StringComparison.OrdinalIgnoreCase));
                        if (otpItem == null)
                        {
                            MarkAddedRow(lOtpRegMapSheet, i + 1);
                            continue;
                        }
                        registerNamelstInRegMap.Add(otpRegisterName.ToUpper());

                        index = 1;
                        if (index<= regMapDT.Columns.Count && !GetString(otpItem.OtpRegisterName).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpRegisterName));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.Name).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.Name));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.InstName).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.InstName));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.InstBase).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.InstBase));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.RegName).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.RegName));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.RegOfs).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.RegOfs));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.OtpOwner).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpOwner));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.DefaultValue).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.DefaultValue));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.Bw).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.Bw));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.Idx).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.Idx));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.Offset).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.Offset));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.OtpB0).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpB0));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.OtpA0).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpA0));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.OtpRegAdd).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpRegAdd));
                        if (index <= regMapDT.Columns.Count && !GetString(otpItem.OtpRegOfs).Equals(lOtpRegMapSheet.Cells[i + 1, index++].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                            MarkMismatchCell(lOtpRegMapSheet, i + 1, index - 1, "Yaml:" + GetString(otpItem.OtpRegOfs));                       
                    }
                }

                int rowIndex = regMapDT.Rows.Count + 1;
                for (int i=0; i< _yamlFile.OtpRows.Count; i++)
                {
                    otpItem = _yamlFile.OtpRows[i];
                    if (!registerNamelstInRegMap.Contains(GetString(otpItem.OtpRegisterName)))
                    {
                        index = 1;
                        if (otpItem.OtpRegisterName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegisterName.ToUpper();
                        if (otpItem.Name != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Name.ToUpper();
                        if (otpItem.InstName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstName.ToUpper();
                        if (otpItem.InstBase != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstBase;
                        if (otpItem.RegName != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegName.ToUpper();
                        if (otpItem.RegOfs != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegOfs;
                        if (otpItem.OtpOwner != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpOwner;
                        if (otpItem.DefaultValue != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultValue.ToUpper();
                        if (otpItem.Bw != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Bw.ToUpper();
                        if (otpItem.Idx != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Idx.ToUpper();
                        if (otpItem.Offset != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Offset.ToUpper();
                        if (otpItem.OtpB0 != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpB0.ToUpper();
                        if (otpItem.OtpA0 != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpA0.ToUpper();
                        if (otpItem.OtpRegAdd != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegAdd.ToUpper();
                        if (otpItem.OtpRegOfs != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegOfs.ToUpper();
                        if (otpItem.DefaultOrReal != null) lOtpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultOrReal;
                        MarkMissingRow(lOtpRegMapSheet, rowIndex);
                        rowIndex++;
                    }
                }

                //Marks
                lOtpRegMapSheet.Cells[1, 1].Value = "Marks:";
                lOtpRegMapSheet.Cells[1, 2].Value = "Mismatched Cell";
                lOtpRegMapSheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                lOtpRegMapSheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                lOtpRegMapSheet.Cells[1, 3].Value = "Register Name in register map but not in Yaml";
                lOtpRegMapSheet.Cells[1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                lOtpRegMapSheet.Cells[1, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                lOtpRegMapSheet.Cells[1, 4].Value = "Register Name not in register map but in Yaml";
                lOtpRegMapSheet.Cells[1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                lOtpRegMapSheet.Cells[1, 4].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                epp.SaveAs(file);
                epp.Dispose();
            }catch(Exception ex)
            {
                throw new Exception("Generate diff report error: " + ex.Message);
            }
        }

        private string GetString(string str)
        {
            if (string.IsNullOrEmpty(str))
                return "";
            return str;
        }

        private void MarkMismatchCell(ExcelWorksheet sheet, int row, int column, string comment)
        {
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);
            sheet.Cells[row, column].AddComment(comment, "Teradyne");
        }

        private void MarkMissingRow(ExcelWorksheet sheet, int row)
        {
            sheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Orange);
        }
        private void MarkAddedRow(ExcelWorksheet sheet, int row)
        {
            sheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }
        private void DumpDataTable2File(DT::DataTable table, string filename)
        {
            using (StreamWriter sw = new StreamWriter(filename, false))
            {
                foreach (DT::DataRow row in table.Rows)
                {
                    sw.Write(string.Join("\t", row.ItemArray));
                    sw.WriteLine();
                }
            }
        }

        private void FetchDataTableInfo(DT::DataTable regMapDT, ref int headerRowIndex, ref int insertRowIndex, ref int insertColIndex, ref int headerKeyColIndex, string headerKeyword)
        {
            int rowIndex = -1;
            foreach (DT::DataRow row in regMapDT.Rows)
            {
                ++rowIndex;
                int colIndex = -1;
                foreach (var item in row.ItemArray)
                {
                    if (item == null || item.ToString().Equals(string.Empty)) break;

                    if (headerRowIndex == -1 && item.ToString().ToUpper().Equals(headerKeyword))
                    {
                        headerKeyColIndex = colIndex + 1;
                        headerRowIndex = rowIndex;
                        insertRowIndex = rowIndex + 1;
                    }
                    ++colIndex;
                    if (item.ToString().ToUpper().Equals("END"))
                    {
                        insertColIndex = colIndex;
                        break;
                    }
                }
                if (insertRowIndex != -1 && insertColIndex != -1) break;
            }
        }
    }
}
