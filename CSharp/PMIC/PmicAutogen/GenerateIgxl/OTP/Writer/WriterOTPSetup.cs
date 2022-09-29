using PmicAutogen.GenerateIgxl.OTP.Reader;
using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class WriterOtpSetup
    {
        private readonly AhbRegisterMapSheet _ahbRegSheet;
        private readonly OtpFileReader _otpReader;

        public WriterOtpSetup(OtpFileReader reader, AhbRegisterMapSheet ahbRegSheet)
        {
            _otpReader = reader;
            _ahbRegSheet = ahbRegSheet;
        }

        public void OutputOtpSetup(string path)
        {
            var fileNameOtpSetup = Path.Combine(path, "OTP_SETUP.txt");

            Directory.CreateDirectory(path);

            if (File.Exists(fileNameOtpSetup))
                File.Delete(fileNameOtpSetup);

            if (_otpReader == null || _otpReader.OtProws == null) return;
            var startAddress = CalculateStartAddress(_otpReader.VersionRawDataRow);
            var projectRevStr = ReadProjectRevStr();
            var projectName = LocalSpecs.CurrentProject;
            string otpEventFaultReg, tpVersionMsb, tpVersionLsb, svnVersionMsb, svnVersionLsb;
            ReadRegisterNames(out otpEventFaultReg, out tpVersionMsb, out tpVersionLsb, out svnVersionMsb,
                out svnVersionLsb);
            using (var writer = new StreamWriter(fileNameOtpSetup))
            {
                writer.WriteLine("Variable\tValue\tOTP Offset Address");
                writer.WriteLine("OTP_Start_Addr\t" + startAddress + "\tComment");
                writer.WriteLine("Product_Rev_str\t" + projectRevStr + "\tUsed in the pattern name");
                writer.WriteLine("Product_str\t" + projectName + "\tProject Name");
                writer.WriteLine("OTP_EVENT_FAULT_REG\t" + otpEventFaultReg + "\tPTM Check register");
                writer.WriteLine(
                    "TPVERSION_MSB\t" + tpVersionMsb + "\tRegister used to burn Test Program Version (MSB)");
                writer.WriteLine(
                    "TPVERSION_LSB\t" + tpVersionLsb + "\tRegister used to burn Test Program Version (LSB)");
                writer.WriteLine("SVN_VERSION_MSB\t" + svnVersionMsb + "\tRegister used to burn SVN Version (MSB)");
                writer.WriteLine("SVN_VERSION_LSB\t" + svnVersionLsb + "\tRegister used to burn SVN Version (LSB)");
                var portMapSheet = TestProgram.IgxlWorkBk.PortMapSheets.Count > 0
                    ? TestProgram.IgxlWorkBk.PortMapSheets.Values.ToList()[0]
                    : null;
                if (portMapSheet != null)
                {
                    var nwireJtagPortSetlist = portMapSheet.PortSets.FindAll(s =>
                        s.PortName.Equals("NWIRE_JTAG", StringComparison.OrdinalIgnoreCase));
                    if (nwireJtagPortSetlist != null)
                        foreach (var nwireJtagPortSet in nwireJtagPortSetlist)
                            foreach (var row in nwireJtagPortSet.PortRows)
                                writer.WriteLine("JTAG_{0}_Pin_Name\t{1}\tJTAG {2} Pin", row.FunctionName, row.FunctionPin,
                                    row.FunctionName);
                }

                writer.WriteLine("END");
            }
        }

        public void OutputOtpPossibleOwnerForVbt(List<string> owner)
        {
            foreach (var file in Directory.GetFiles(FolderStructure.DirOtherWaitForClassify))
            {
                var basMain = new BasMain();
                var lines = File.ReadAllLines(file).ToList();
                var targetLine = basMain.SearchContent(lines, new List<string> { "gS_AHBCheckCondition", "const" });
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

        private string ReadProjectRevStr()
        {
            var validPatRow = InputFiles.PatternListMap.PatternListCsvRows.Find(s =>
                s.PatternName.ToUpper().StartsWith("PP_") || s.PatternName.ToUpper().StartsWith("DD_"));
            if (validPatRow != null)
                if (validPatRow.PatternName.Split('_').Length >= 2)
                    return validPatRow.PatternName.Split('_')[1];
            return "";
        }

        private string CalculateStartAddress(string yamlVersionRow)
        {
            var otpMapSizeReg = new Regex(@"otp_map_size[\s]*:[\s]*(?<value>[\d]+)", RegexOptions.IgnoreCase);
            var otpFullSizeReg = new Regex(@"otp_full_size[\s]*:[\s]*(?<value>[\d]+)", RegexOptions.IgnoreCase);
            var otpMapSize = otpMapSizeReg.Match(yamlVersionRow).Groups["value"].ToString();
            var otpFullSize = otpFullSizeReg.Match(yamlVersionRow).Groups["value"].ToString();
            int intOtpMapSize, intotpFullSize;
            if (int.TryParse(otpMapSize, out intOtpMapSize) && int.TryParse(otpFullSize, out intotpFullSize))
                return "&H" + (intotpFullSize - intOtpMapSize).ToString("X");
            return "";
        }

        private void ReadRegisterNames(out string otpEventFaultReg, out string tpVersionMsb, out string tpVersionLsb,
            out string svnVersionMsb, out string svnVersionLsb)
        {
            var canNotFindRegisterName = "Cannot find this register";
            otpEventFaultReg = canNotFindRegisterName;
            tpVersionMsb = canNotFindRegisterName;
            tpVersionLsb = canNotFindRegisterName;
            svnVersionMsb = canNotFindRegisterName;
            svnVersionLsb = canNotFindRegisterName;
            OtpRegister targetRegister = null;
            //TPVERSION_MSB
            targetRegister = _otpReader.OtProws.Find(s =>
                Regex.IsMatch(s.Name.Trim(), @"^MAJOR_[O]?TP_[RE]?VERSION$", RegexOptions.IgnoreCase));
            if (targetRegister != null)
                tpVersionMsb = targetRegister.OtpRegisterName;
            //TPVERSION_LSB
            targetRegister = _otpReader.OtProws.Find(s =>
                Regex.IsMatch(s.Name.Trim(), @"^MINOR_[O]?TP_[RE]?VERSION$", RegexOptions.IgnoreCase));
            if (targetRegister != null)
                tpVersionLsb = targetRegister.OtpRegisterName;
            //SVN_VERSION_MSB
            targetRegister =
                _otpReader.OtProws.Find(s => s.Name.Trim().Equals("LCK1", StringComparison.OrdinalIgnoreCase));
            if (targetRegister == null)
                targetRegister =
                    _otpReader.OtProws.Find(s => s.Name.Trim().Equals("ATE1", StringComparison.OrdinalIgnoreCase));
            if (targetRegister != null)
                svnVersionMsb = targetRegister.OtpRegisterName;
            //SVN_VERSION_LSB
            targetRegister =
                _otpReader.OtProws.Find(s => s.Name.Trim().Equals("LCK0", StringComparison.OrdinalIgnoreCase));
            if (targetRegister != null)
                svnVersionLsb = targetRegister.OtpRegisterName;
            //OTP_EVENT_FAULT_REG
            var targetAhbRegister = _ahbRegSheet.AhbRegRows.Find(s =>
                s.FieldName.Trim().Equals("FLT_OTP_CRC", StringComparison.OrdinalIgnoreCase));
            if (targetAhbRegister == null)
                targetAhbRegister = _ahbRegSheet.AhbRegRows.Find(s =>
                    s.FieldName.Trim().Equals("EVT_CRC", StringComparison.OrdinalIgnoreCase));
            if (targetAhbRegister != null)
                otpEventFaultReg = targetAhbRegister.RegName + "." + targetAhbRegister.FieldName;
        }
    }
}