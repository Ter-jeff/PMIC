using PmicAutomation.MyControls;
using PmicAutomation.Utility.OTPRegisterMap.Base;
using PmicAutomation.Utility.OTPRegisterMap.Input;
using PmicAutomation.Utility.OTPRegisterMap.Output;
using Library.Function.ErrorReport;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System;
using System.Data;

namespace PmicAutomation.Utility.OTPRegisterMap
{
    public class OtpRegisterMapMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _yamlFilePath;
        private readonly string _otpFilesPath;
        private readonly string _regmapFilePath;
        private readonly string _outputPath;

        //private List<string> _otpFileHeaders;
        //private List<Tuple<int, string>> _otpOriHeaders;
        //private List<OtpRegisterItem> _otpRegisterItems;
        //private List<string> _otpVersionList;

        private OtpFileReader _yamlFile = null;
        private List<OtpFileReader> _otpFileList = new List<OtpFileReader>();
        private DataTable _regMapFile = null;

        public OtpRegisterMapMain(OtpRegisterMapFrom otpFrom)
        {
            _appendText = otpFrom.AppendText;
            _outputPath = otpFrom.FileOpen_OutputPath.ButtonTextBox.Text.Trim();
            _yamlFilePath = otpFrom.FileOpen_Yaml.ButtonTextBox.Text.Trim();
            _otpFilesPath = otpFrom.FilesOpen_Otp.ButtonTextBox.Text.Trim();
            _regmapFilePath = otpFrom.FileOpen_RegMap.ButtonTextBox.Text.Trim();
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            if (!PreCheck())
                return;
            
            ReadFiles();

            GenFiles();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private bool PreCheck()
        {
            //.yaml and .otp
            List<string> otpFiles = _otpFilesPath.Split(',').ToList();
            otpFiles = otpFiles.Select(s => s.Trim()).ToList();

            if (!string.IsNullOrEmpty(_yamlFilePath) && !File.Exists(_yamlFilePath))
            {
                _appendText.Invoke("yaml file is not exist: " + _yamlFilePath, Color.Red);
                return false;
            }

            foreach (string otpFilePath in otpFiles)
            {
                if (!string.IsNullOrEmpty(otpFilePath) && !File.Exists(otpFilePath))
                {
                    _appendText.Invoke("otp file is not exist: " + otpFilePath, Color.Red);
                    return false;
                }
            }

            if (!string.IsNullOrEmpty(_regmapFilePath) && !File.Exists(_regmapFilePath))
            {
                _appendText.Invoke("register map file is not exist: " + _regmapFilePath, Color.Red);
                return false;
            }

            if (string.IsNullOrEmpty(_outputPath) || !Directory.Exists(_outputPath))
            {
                _appendText.Invoke("Output Folder is not exist: " + _outputPath, Color.Red);
                return false;
            }

            return true;
        }

        private void ReadFiles()
        {
            //.yaml and .otp
            List<string> otpFiles = _otpFilesPath.Split(',').ToList();
            otpFiles = otpFiles.Select(s => s.Trim()).ToList();

            if (!string.IsNullOrEmpty(_yamlFilePath))
            {
                _yamlFile = new OtpFileReader(_yamlFilePath);
            }

            foreach(string otpFilePath in otpFiles)
            {
                if (!string.IsNullOrEmpty(otpFilePath))
                {
                    _otpFileList.Add(new OtpFileReader(otpFilePath));
                }
            }

            if (!string.IsNullOrEmpty(_regmapFilePath))
            {
                _regMapFile = Library.Common.Utility.ConvertToDataTable(_regmapFilePath, new char[] { '\t' }, "OTP_REGISTER_NAME");
                _regMapFile.TableName = _regmapFilePath;
            }


            //var yamlReader = new OtpFileReader(_yamlFilePath);
            //foreach (var otpFileName in otpFiles)
            //{
            //    var otpReader = new OtpFileReader(otpFileName);
            //    yamlReader.MergeOtpToYaml(otpReader.OtpRows, otpReader.Version);
            //}

            //_otpFileHeaders = yamlReader.Headers;
            //_otpOriHeaders = yamlReader.OriHeaderInfo;
            //_otpRegisterItems = yamlReader.OtpRows;
            //_otpVersionList = yamlReader.VersionList;
        }

        private void GenFiles()
        {
            GenOtp();

            GenErrorReport();
        }

        private void GenOtp()
        {
            //WriterOtpRegisterMap writerOtpRegisterMap = new WriterOtpRegisterMap(_otpFileHeaders, _otpOriHeaders, _otpRegisterItems, _otpVersionList);

            WriterOtpRegisterMap writerOtpRegisterMap = new WriterOtpRegisterMap(_yamlFile, _otpFileList, _regMapFile);
            writerOtpRegisterMap.OutPutResult(_outputPath);
            //if (!_regmapFilePath.Equals(string.Empty) && File.Exists(_regmapFilePath))
            //    writerOtpRegisterMap.OutPutOtpRegisterMapWithTxtFile(_outputPath, _regmapFilePath);
            //else
            //    writerOtpRegisterMap.OutPutOtpRegisterMap(_outputPath, "OTP_Register_Map", true);
        }

        private void GenErrorReport()
        {
            if (ErrorManager.GetErrorCount() <= 0)
            {
                return;
            }

            _appendText.Invoke("Starting to print error report ...", Color.Red);
            string outputFile = Path.Combine(_outputPath, "Error.xlsx");
            List<string> files = new List<string>();
            ErrorManager.GenErrorReport(outputFile, files);
        }
    }
}