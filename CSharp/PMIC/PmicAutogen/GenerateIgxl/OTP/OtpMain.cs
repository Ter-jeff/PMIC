using CommonLib.Enum;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.GenerateIgxl.OTP.Writer;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.IO;

namespace PmicAutogen.GenerateIgxl.OTP
{
    public class OtpMain : MainBase
    {
        private OtpDefaultRegList _otpDefaultRegList;

        public void WorkFlow()
        {
            try
            {
                Response.Report("Reading OTP Files...", EnumMessageLevel.General, 5);

                Read();

                Response.Report("Generating OTP Files ...", EnumMessageLevel.General, 60);

                WriteOtpFiles();

                SetPinMapByOtpSetup();

                GenOtpSetup();

                Response.Report("OTP Completed!", EnumMessageLevel.General, 100);
            }

            catch (Exception e)
            {
                Response.Report("OTP Function Error.Please sent blow Message to administrator.", EnumMessageLevel.General,
                    0);
                Response.Report("Errors occurs: " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        private static void GenOtpSetup()
        {
            var fileNameTxt = Path.Combine(FolderStructure.DirOtp, PmicConst.OtpSetup + ".txt");
            InputFiles.TestPlanWorkbook.Worksheets[PmicConst.OtpSetup].ExportToTxt(fileNameTxt);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirOtp, PmicConst.OtpSetup);
        }

        public void Read()
        {
            //.yaml and .otp
            if (StaticOtp.OtpFileReader != null)
            {
                // OTP Default Register List
                _otpDefaultRegList = new OtpDefaultRegList();
                if (InputFiles.TestPlanWorkbook != null &&
                    InputFiles.TestPlanWorkbook.Worksheets["otp_default_reglist"] != null)
                    _otpDefaultRegList.Read(InputFiles.TestPlanWorkbook.Worksheets["otp_default_reglist"]);
                else if (StaticOtp.OtpFileReader.OtProws.Count > 0)
                    _otpDefaultRegList.ConvertRegisterMapToDefaultRegList(StaticOtp.OtpFileReader.OtProws);
            }
        }

        protected void WriteOtpFiles()
        {
            var writerOtpRegisterMap = new WriterOtpRegisterMap(StaticOtp.OtpFileReader);
            writerOtpRegisterMap.OutPutOtpRegisterMap(FolderStructure.DirOtp, PmicConst.OtpRegisterMap, true);

            var writerAhbRegisterMap = new WriterAhbRegisterMap(StaticTestPlan.AhbRegisterMapSheet);
            writerAhbRegisterMap.OutPutAhbRegisterMap(FolderStructure.DirOtp, PmicConst.AhbRegisterMap);
            writerAhbRegisterMap.OutPutAhbEnum(FolderStructure.DirOtp);

            var writerOtpSetup = new WriterOtpSetup(StaticOtp.OtpFileReader, StaticTestPlan.AhbRegisterMapSheet);
            writerOtpSetup.OutputOtpSetup(FolderStructure.DirOtp);

            #region Edit Data For IGXL

            //var otpIgxlWorkFlow = new OtpIgxlWorkFlow();
            //var subFlowSheets = otpIgxlWorkFlow.GetSheets();
            //foreach (var subFlowSheet in subFlowSheets)
            //    _igxlSheets.Add(subFlowSheet, FolderStructure.DirOtp);

            //var otpIgxlInstance = new OtpIgxlInstance();
            //var instanceSheets = otpIgxlInstance.GetSheets();
            //foreach (var instanceSheet in instanceSheets)
            //    _igxlSheets.Add(instanceSheet, FolderStructure.DirOtp);

            //var sheet = InputFiles.SettingWorkbook.Worksheets[OtpConst.PatSetsOtp];
            //if (sheet != null)
            //{
            //    var readPatSetSheet = new ReadPatSetSheet();
            //    _igxlSheets.Add(readPatSetSheet.GetSheet(sheet), FolderStructure.DirOtp);
            //}

            #endregion
        }

        private void SetPinMapByOtpSetup()
        {
            var setPinName = new SetJatgPinName();
            setPinName.SetPinMapByOtpSetup();
        }
    }
}