using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class WriterOtpRegisterMap
    {
        private readonly OtpFileReader _otpReader;

        public WriterOtpRegisterMap(OtpFileReader reader)
        {
            _otpReader = reader;
        }

        public void OutPutOtpRegisterMap(string path, string sheetName, bool outputPath, bool bypassEditVbt = false)
        {
            var fileNameXls = Path.Combine(path, sheetName + ".xlsx");
            var fileNameCsv = Path.Combine(path, sheetName + ".csv");
            var fileNameTxt = Path.Combine(path, sheetName + ".txt");

            Directory.CreateDirectory(path);

            if (File.Exists(fileNameXls))
                File.Delete(fileNameXls);

            if (File.Exists(fileNameCsv))
                File.Delete(fileNameCsv);

            if (File.Exists(fileNameTxt))
                File.Delete(fileNameTxt);

            if (_otpReader == null || _otpReader.OtProws == null) return;

            var excel = new ExcelPackage(new FileInfo(fileNameXls));
            var eFuseWorkbook = excel.Workbook;

            var otpRegMapSheet = eFuseWorkbook.Worksheets.Add(sheetName);
            var rowIndex = 3;
            otpRegMapSheet.Cells[1, 1].Value = "";

            for (var i = 0; i < _otpReader.FileHeaders.Count; i++)
                otpRegMapSheet.Cells[2, i + 1].Value = _otpReader.FileHeaders[i];


            for (var i = 0; i < _otpReader.Headers.Count; i++)
                otpRegMapSheet.Cells[rowIndex, i + 1].Value = _otpReader.Headers[i];

            otpRegMapSheet.Cells[rowIndex, _otpReader.Headers.Count + 1].Value = "END";
            rowIndex++;

            foreach (var otpItem in _otpReader.OtProws)
            {
                int index = 1;
                if (otpItem.OtpRegisterName != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegisterName.ToUpper();
                if (otpItem.Name != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Name.ToUpper();
                if (otpItem.InstName != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstName.ToUpper();
                if (otpItem.InstBase != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.InstBase;
                if (otpItem.RegName != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegName.ToUpper();
                if (otpItem.RegOfs != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.RegOfs;
                if (otpItem.OtpOwner != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpOwner;
                if (otpItem.DefaultValue != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultValue.ToUpper();
                if (otpItem.Bw != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Bw.ToUpper();
                if (otpItem.Idx != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Idx.ToUpper();
                if (otpItem.Offset != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.Offset.ToUpper();
                if (otpItem.OtpB0 != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpB0.ToUpper();
                if (otpItem.OtpA0 != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpA0.ToUpper();
                if (otpItem.OtpRegAdd != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegAdd.ToUpper();
                if (otpItem.OtpRegOfs != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.OtpRegOfs.ToUpper();
                if (otpItem.DefaultOrReal != null) otpRegMapSheet.Cells[rowIndex, index++].Value = otpItem.DefaultOrReal;
                otpRegMapSheet.Cells[rowIndex, index++].Value = "";
                otpRegMapSheet.Cells[rowIndex, index++].Value = "";
                otpRegMapSheet.Cells[rowIndex, index++].Value = "";
                otpRegMapSheet.Cells[rowIndex, index++].Value = "0";
                //int i = 20;
                foreach (var extraItem in otpItem.OtpExtra)
                {
                    otpRegMapSheet.Cells[rowIndex, index++].Value = extraItem;
                    //i++;
                }
                rowIndex++;
            }

            //otpRegMapSheet.Cells[rowIndex, 1].LoadFromCollection(_otpReader.OtProws);
            //int otpEcidOnlyIndex = _otpReader.Headers.IndexOf("OTP_ECID_ONLY") + 1;
            //foreach (var otpItem in _otpReader.OtProws)
            //{
            //    otpRegMapSheet.Cells[rowIndex, otpEcidOnlyIndex -3].Value = "";
            //    otpRegMapSheet.Cells[rowIndex, otpEcidOnlyIndex - 2].Value = "";
            //    otpRegMapSheet.Cells[rowIndex, otpEcidOnlyIndex -1].Value = "";
            //    var i = otpEcidOnlyIndex;
            //    otpRegMapSheet.Cells[rowIndex, i].Value = "0";
            //    i++;
            //    foreach (var extraItem in otpItem.OtpExtra)
            //    {
            //        otpRegMapSheet.Cells[rowIndex, i].Value = extraItem;
            //        i++;
            //    }

            //    rowIndex++;
            //}

            otpRegMapSheet.Cells[rowIndex, 1].Value = "END";
            otpRegMapSheet.ExportToTxt(fileNameTxt);
            otpRegMapSheet.ExportToTxt(fileNameCsv, ",");


            //temp use for confirm
            var allOwner = _otpReader.OtProws.Select(p => p.OtpOwner).Distinct().ToList();
            if (!bypassEditVbt)
                OutputOtpPossibleOwnerForVbt(allOwner);

            try
            {
                excel.Save();
                excel.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Write error report failed. " + e.Message);
            }

            TestProgram.NonIgxlSheetsList.Add(path, sheetName);
        }

        public void OutputOtpPossibleOwnerForVbt(List<string> owner)
        {
            foreach (var file in Directory.GetFiles(FolderStructure.DirLib))
            {
                var basMain = new BasMain();
                var lines = File.ReadAllLines(file).ToList();
                var targetLine = basMain.SearchContent(lines, new List<string> {"gS_AHBCheckCondition", "const"});
                if (!string.IsNullOrEmpty(targetLine))
                {
                    var regEdit = @"=\s*\w*";
                    if (Regex.IsMatch(targetLine, regEdit, RegexOptions.IgnoreCase))
                    {
                        var newline = targetLine.Split('\'')[0] +
                                      string.Format("\'Remove {0}; Filter by OTP_Owner", string.Join(",", owner));
                        var index = lines.IndexOf(targetLine);
                        lines[index] = newline;
                    }

                    var writer = new StreamWriter(file);
                    writer.WriteLine(lines);
                    break;
                }
            }
        }
    }
}