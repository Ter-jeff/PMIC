using AutomationCommon.DataStructure;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.GenerateIgxl.OTP.Writer;
using PmicAutogen.InputPackages;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.IO;
using AutomationCommon.Utility;

namespace PmicAutogen.GenerateIgxl.OTP
{
    public class OtpMain : MainBase
    {
        private OtpDefaultRegList _otpDefaultRegList;

        public void WorkFlow()
        {
            try
            {
                Response.Report("Reading OTP Files...", MessageLevel.General, 5);
                Read();

                Response.Report("Generating OTP Files ...", MessageLevel.General, 60);

                WriteOtpFiles();

                SetPinMapByOtpSetup();
                //GenerateBinTableRows();

                GenOtpSetup();

                Response.Report("OTP Completed!", MessageLevel.General, 100);
            }

            catch (Exception e)
            {
                Response.Report("OTP Function Error.Please sent blow Message to administrator.", MessageLevel.General,
                    0);
                Response.Report("Errors occurs: " + e.Message, MessageLevel.Error, 0);
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

            var writerOTPSetup = new WriterOTPSetup(StaticOtp.OtpFileReader, StaticTestPlan.AhbRegisterMapSheet);
            writerOTPSetup.outputOTPSetup(FolderStructure.DirOtp);

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

        protected void GenerateBinTableRows()
        {
            //Bin_OTP_Type1	F_OTP_Type1	AND	1	1	Pass	T																																																																																	
            //Bin_OTP_Type2	F_OTP_Type2	AND	2	2	Pass	T																																																																																	
            //Bin_OTP_Type3	F_OTP_Type3	AND	3	3	Pass	T																																																																																	
            //Bin_OTP_Type4	F_OTP_Type4	AND	4	4	Pass	T																																																																																	
            //Bin_OTP_LOCKBIT	F_OTP_LOCKBIT	AND	950	8	Fail	T																																																																																	
            //Bin_OTP_CHECK_DefaultReal	F_OTP_CHECK_DefaultReal	AND	951	8	Fail	T																																																																																	
            //Bin_OTP_AHBvsOTP_PreBurn_Comp	F_OTP_AHBvsOTP_PreBurn_Comp	AND	952	8	Fail	T																																																																																	
            //Bin_OTP_AHBvsOTP_AfterBurn_Comp	F_OTP_AHBvsOTP_AfterBurn_Comp	AND	953	8	Fail	T																																																																																	
            //Bin_OTP_Burn_ECID	F_OTP_Burn_ECID	AND	954	8	Fail	T																																																																																	
            //Bin_OTP_Burn_CRC	F_OTP_Burn_CRC	AND	955	8	Fail	T																																																																																	
            //Bin_OTP_CRC_PostBurn	F_OTP_CRC_PostBurn	AND	956	8	Fail	T																																																																																	
            //Bin_OTP_EW_Cnt	Flag_OTP_EW_Cnt	AND	957	8	Fail	T				

            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRow(GenBinTableRow("Bin_OTP_Type1", "F_OTP_Type1", "AND", "1", "1", "Pass",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Type2", "F_OTP_Type2", "AND", "2", "2", "Pass",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Type3", "F_OTP_Type3", "AND", "3", "3", "Pass",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Type4", "F_OTP_Type4", "AND", "4", "4", "Pass",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Init", "F_OTP_Init", "AND", "957", "8", "Fail",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_LOCKBIT", "F_OTP_LOCKBIT", "AND", "950", "8", "Fail",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_CHECK_DefaultReal", "F_OTP_CHECK_DefaultReal", "AND", "951", "8",
                "Fail", new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_AHBvsOTP_PreBurn_Comp", "F_OTP_AHBvsOTP_PreBurn_Comp", "AND", "952",
                "8", "Fail", new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_AHBvsOTP_AfterBurn_Comp", "F_OTP_AHBvsOTP_AfterBurn_Comp", "AND",
                "953", "8", "Fail", new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Burn_ECID", "F_OTP_Burn_ECID", "AND", "954", "8", "Fail",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_Burn_CRC", "F_OTP_Burn_CRC", "AND", "955", "8", "Fail",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_CRC_PostBurn", "F_OTP_CRC_PostBurn", "AND", "956", "8", "Fail",
                new List<string> {"T"}));
            binTable.AddRow(GenBinTableRow("Bin_OTP_EW_Cnt", "Flag_OTP_EW_Cnt", "AND", "957", "8", "Fail",
                new List<string> {"T"}));
        }

        private BinTableRow GenBinTableRow(string name, string itemList, string op, string sort, string bin,
            string result, List<string> items)
        {
            var binRow = new BinTableRow();
            binRow.Name = name;
            binRow.ItemList = itemList;
            binRow.Op = op;
            binRow.Sort = sort;
            binRow.Bin = bin;
            binRow.Result = result;
            binRow.Items.AddRange(items);
            return binRow;
        }

        private void SetPinMapByOtpSetup()
        {
            var setPinName = new SetJATGPinName();
            setPinName.SetPinMapByOTPSetup();
        }
    }
}