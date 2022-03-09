using CommonLib.Controls;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OTPFileComparison
{
    public class FilesCompareMain
    {
        private string _outputFolder;
        private List<OTPFileInfo> _otpFiles;
        private Action<string> _updateUIInfo;
        private Action<int> _updateUIProgressBar;

        //Compared Header Name regex
        private string reg_compColDefaultValue = @"_Default_Value$";
        private string reg_compColAhbReadValue = @"^AHB Read Value$";
        private string reg_compColOTPDbValue = @"^OTP DataBase Value$";
        public FilesCompareMain(List<OTPFileInfo> otpFiles, string outputFolder, Action<string> setStatusLabel, Action<int> setProgressBarValue)
        {
            _otpFiles = otpFiles;
            _outputFolder = outputFolder;
            _updateUIInfo = setStatusLabel;
            _updateUIProgressBar = setProgressBarValue;
        }

        public void Compare()
        {
            //Read OTP Files
            _updateUIProgressBar.Invoke(0);
            _updateUIInfo.Invoke("Reading OTP Files....");
            List<OTPFile> otpFileDatalst = new List<OTPFile>();
            OtpFileReader otpFileReader = new OtpFileReader();
            _otpFiles.ForEach(s => otpFileDatalst.Add(otpFileReader.Read(s.FileName)));
            //Comparing and write diff report
            string time = DateTime.Now.ToString("yyMMddHHmmss");
            string reportFile = Path.Combine(_outputFolder, "OTP_Diff_Report_" + time + ".xlsx");
            ExcelPackage epp = new ExcelPackage();
            ExcelWorkbook workbook = epp.Workbook;
            ExcelWorksheet sheet;

            _updateUIProgressBar.Invoke(20);
            _updateUIInfo.Invoke("OTP Files comparing......");
            if (_otpFiles.Exists(s => s.VerticalComparison))
            {
                sheet = workbook.Worksheets.Add("FilesComparingDiffReport");
                int currentRow = 1;
                for(int i =0; i< _otpFiles.Count; i++)
                {
                    if (i>0 && _otpFiles[i].VerticalComparison)
                    {
                        WriteFileToFileCompareResult(sheet, otpFileDatalst[i - 1], otpFileDatalst[i], ref currentRow);
                    }
                }
            }


            _updateUIProgressBar.Invoke(70);
            _updateUIInfo.Invoke("OTP File Self comparing......");
            for (int i = 1; i<=_otpFiles.Count; i++)
            {
                if (!_otpFiles[i-1].HorizontalComparison)
                    continue;
                sheet = workbook.Worksheets.Add("File_" + i + "_SelfCompareDiffReport");
                WriteFileSelfCompareResult(sheet, otpFileDatalst[i - 1]);
            }

            var file = new FileInfo(reportFile);
            epp.SaveAs(file);
            epp.Dispose();

            _updateUIProgressBar.Invoke(100);
            _updateUIInfo.Invoke("Done!");
        }

        /// <summary>
        /// Only compare "AHB Read Value" value in different OTP files
        /// </summary>
        /// <param name="resultSheet"></param>
        /// <param name="baseFile"></param>
        /// <param name="compFile"></param>
        /// <param name="currentRow"></param>
        private void WriteFileToFileCompareResult(ExcelWorksheet resultSheet, OTPFile baseFile, OTPFile compFile, ref int currentRow)
        {           
            int row = currentRow;
            bool hasDiff = false;
            //Compared Header column index
            int baseFileCompCol = FindColIndexByHeaderNameRegex(baseFile.Headers, reg_compColAhbReadValue);
            int compFileCompCol = FindColIndexByHeaderNameRegex(compFile.Headers, reg_compColAhbReadValue);
            if (baseFileCompCol < 0)
                throw new Exception("Can not find header 'AHB Read Value' in OTP file: " + baseFile.FileName);
            if (compFileCompCol < 0)
                throw new Exception("Can not find header 'AHB Read Value' in OTP file: " + compFile.FileName);

            resultSheet.Cells[row++, 1].Value = "Base File Path: " + baseFile.FileName;
            resultSheet.Cells[row++, 1].Value = "Comparing File Path: " + compFile.FileName;
            string baseValue, compValue;
            for (int i = 0; i<baseFile.OTPRows.Count; i++)
            {
                baseValue = baseFile.OTPRows[i].Count >= baseFileCompCol ? baseFile.OTPRows[i][baseFileCompCol - 1] : "";
                compValue = "";
                if (compFile.OTPRows.Count > i && compFile.OTPRows[i].Count >= compFileCompCol)
                    compValue = compFile.OTPRows[i][compFileCompCol - 1];
                if (baseValue.Equals(compValue, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!hasDiff)
                {
                    //Write Header
                    foreach (string header in baseFile.Headers.Keys)
                    {
                        // Header "[XXX]_Default_Value" may be different for each otp file, if this header name is not the same, should show up in header
                        if (Regex.IsMatch(header, reg_compColDefaultValue, RegexOptions.IgnoreCase))
                        {
                            string compHeaderName = compFile.Headers.Keys.ToList().Find(s => Regex.IsMatch(s, reg_compColDefaultValue, RegexOptions.IgnoreCase));
                            if (compHeaderName == null)
                                compHeaderName = "";
                            if(!header.Equals(compHeaderName, StringComparison.OrdinalIgnoreCase))
                            {
                                resultSheet.Cells[row, baseFile.Headers[header]].Value = header + "->" + compHeaderName;
                                MarkRedCell(resultSheet, row, baseFile.Headers[header]);
                            }
                            else
                            {
                                resultSheet.Cells[row, baseFile.Headers[header]].Value = header;
                            }
                        }
                        else
                        {
                            resultSheet.Cells[row, baseFile.Headers[header]].Value = header;
                        }
                    }
                    row++;
                    hasDiff = true;
                }

                for (int col = 1; col <= baseFile.Headers.Count; col++)
                {
                    if (baseFile.OTPRows[i].Count >= col)
                    {
                        resultSheet.Cells[row, col].Value = baseFile.OTPRows[i][col - 1];
                        if (col == baseFileCompCol)
                        {
                            resultSheet.Cells[row, col].Value = baseValue + "->" + compValue;
                            MarkRedCell(resultSheet, row, col);
                        }                        
                    }
                }
                row++;
            }

            if (!hasDiff)
            {
                resultSheet.Cells[row++, 1].Value = "No Diff";
            }

            currentRow = row + 1;
        }

        /// <summary>
        /// Comare "[XXX]_Default_Value" =>"AHB Read Value"
        /// [XXX]_Default_Value" => "OTP DataBase Value"
        /// "AHB Read Value" = > "OTP DataBase Value"
        /// in one OTP file
        /// </summary>
        /// <param name="resultSheet"></param>
        /// <param name="otpFile"></param>
        private void WriteFileSelfCompareResult(ExcelWorksheet resultSheet, OTPFile otpFile)
        {
            //Compared header name regex
            
            //Compared header column index
            int compColDefaultValueIndex = FindColIndexByHeaderNameRegex(otpFile.Headers, reg_compColDefaultValue);
            int compColAhbReadIndex = FindColIndexByHeaderNameRegex(otpFile.Headers, reg_compColAhbReadValue);
            int compColOTPDbValueIndex = FindColIndexByHeaderNameRegex(otpFile.Headers, reg_compColOTPDbValue);
            if (compColDefaultValueIndex < 0 ||
                compColAhbReadIndex<0 ||
                compColAhbReadIndex<0)
            {
                throw new Exception("Can not find header '[XXX]_Default_Value','AHB Read Value' or 'OTP DataBase Value' in OTP File: " + otpFile.FileName);
            }
            string compColDefaultValueColumnName = NumberToExcelColumnName(compColDefaultValueIndex);
            string compColAhbReadColumnName = NumberToExcelColumnName(compColAhbReadIndex);
            string compColOTPDbValueColumnName = NumberToExcelColumnName(compColOTPDbValueIndex);

            resultSheet.Cells[1, 1].Value = "File Path: " + otpFile.FileName;
            //Write Header
            foreach(string header in otpFile.Headers.Keys)
            {
                resultSheet.Cells[3, otpFile.Headers[header]].Value = header;
            }
            resultSheet.Cells[3, otpFile.Headers.Count + 2].Value = compColDefaultValueColumnName + "->" + compColAhbReadColumnName;
            resultSheet.Cells[3, otpFile.Headers.Count + 3].Value = compColDefaultValueColumnName + "->" + compColOTPDbValueColumnName;
            resultSheet.Cells[3, otpFile.Headers.Count + 4].Value = compColAhbReadColumnName + "->" + compColOTPDbValueColumnName;

            //Write Data
            int row = 4;
            foreach(List<string> data in otpFile.OTPRows)
            {
                for(int col = 1; col<= otpFile.Headers.Count; col++)
                {
                    if(data.Count >= col)
                        resultSheet.Cells[row, col].Value = data[col - 1];
                }

                if (data.Count >= otpFile.Headers.Count)
                {
                    resultSheet.Cells[row, otpFile.Headers.Count + 2].Value = data[compColDefaultValueIndex - 1].Equals(data[compColAhbReadIndex - 1], StringComparison.OrdinalIgnoreCase) ? "TRUE" : "FALSE";
                    resultSheet.Cells[row, otpFile.Headers.Count + 3].Value = data[compColDefaultValueIndex - 1].Equals(data[compColOTPDbValueIndex - 1], StringComparison.OrdinalIgnoreCase) ? "TRUE" : "FALSE";
                    resultSheet.Cells[row, otpFile.Headers.Count + 4].Value = data[compColAhbReadIndex - 1].Equals(data[compColOTPDbValueIndex - 1], StringComparison.OrdinalIgnoreCase) ? "TRUE" : "FALSE";
                    if (resultSheet.Cells[row, otpFile.Headers.Count + 2].Value.ToString() == "FALSE")
                        MarkRedCell(resultSheet, row, otpFile.Headers.Count + 2);
                    if (resultSheet.Cells[row, otpFile.Headers.Count + 3].Value.ToString() == "FALSE")
                        MarkRedCell(resultSheet, row, otpFile.Headers.Count + 3);
                    if (resultSheet.Cells[row, otpFile.Headers.Count + 4].Value.ToString() == "FALSE")
                        MarkRedCell(resultSheet, row, otpFile.Headers.Count + 4);
                }
                row++;
            }

        }

        /// <summary>
        /// Find column index by headerName regex
        /// </summary>
        /// <param name="Headers">key:HeaderName, value:Column index</param>
        /// <param name="headerNameRegex"></param>
        /// <returns></returns>
        private int FindColIndexByHeaderNameRegex(Dictionary<string, int> Headers, string headerNameRegex)
        {
            foreach (string key in Headers.Keys.ToList())
            {
                if (Regex.IsMatch(key, headerNameRegex, RegexOptions.IgnoreCase))
                    return Headers[key];
            }
            return -1;
        }
        private string NumberToExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        private void MarkRedCell(ExcelWorksheet sheet, int row, int column)
        {
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);            
        }
    }
}
