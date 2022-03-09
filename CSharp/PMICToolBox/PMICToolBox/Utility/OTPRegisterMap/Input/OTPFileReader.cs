using PmicAutomation.Utility.OTPRegisterMap.Base;
using Library.Function.ErrorReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PmicAutomation.Utility.OTPRegisterMap.Input
{
    public class OtpFileReader
    {
        public string Filename { get; set; }
        public string FileNameWithoutExtension { get; set; }
        private List<string> Fields { get; set; }
        private List<string> RowData { get; set; }
        private List<string> Keys { get; set; }
        public readonly List<string> Headers = new List<string>();
        public readonly List<OtpRegisterItem> OtpRows = new List<OtpRegisterItem>();
        public readonly List<string> VersionList = new List<string>();
        public List<Tuple<int, string>> OriHeaderInfo = new List<Tuple<int, string>>();

        public OtpFileReader(string filename)
        {
            Filename = filename;
            FileNameWithoutExtension = Path.GetFileNameWithoutExtension(Filename);
            RowData = GetYamlData();
            Keys = OrderedKeys(RowData);
            Fields = SortByFieldNames(RowData);
            UpdateHeader(Fields);
            foreach (var data in RowData)
            {
                var otpItem = ConvertRawDataToOtpItem(data);
                if (otpItem != null)
                    OtpRows.Add(otpItem);
            }
        }

        private List<string> SortByFieldNames(List<string> data)
        {
            var fieldnames = new List<string>();
            foreach (var line in data)
            {
                if (line.StartsWith("#")) continue;
                if (line.Contains(":") && line.Contains("{"))
                {
                    if (line.Split(':')[0].StartsWith("__")) continue;
                    var content = line.Split('{')[1].Replace("}", "");
                    foreach (var pair in content.Split(','))
                    {
                        fieldnames.Add(pair.Split(':')[0].Trim());
                    }
                    fieldnames.Remove("reg_addr");
                    fieldnames.Remove("desc");
                    fieldnames.Remove("htmldesc");
                    return fieldnames;
                }
            }
            return fieldnames;
        }

        private List<string> GetYamlData()
        {
            var reader = new StreamReader(Filename);
            var contents = new List<string>();
            string line;
            var preInclude = 0;
            var postInclude = 0;
            var linetmp = "";
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains("{"))
                    preInclude = preInclude + line.Count(x => x.Equals('{'));
                if (line.Contains("}"))
                    postInclude = postInclude + line.Count(x => x.Equals('}'));

                if (preInclude != 0)
                {
                    linetmp = linetmp + line;
                    if (preInclude == postInclude)
                    {
                        contents.Add(linetmp);
                        linetmp = "";
                        preInclude = 0;
                        postInclude = 0;
                    }

                }
            }
            return contents;
        }

        private List<string> OrderedKeys(List<string> data)
        {
            var keys = new List<string>();
            foreach (var line in data)
            {
                if (line.StartsWith("#")) continue;
                if (line.Contains(":"))
                {
                    var key = line.Split(':')[0];
                    if (!key.StartsWith("__")) keys.Add(key);
                }
            }
            return keys;
        }

        private OtpRegisterItem ConvertRawDataToOtpItem(string rawData)
        {
            if (rawData.StartsWith("#")) return null;
            if (!rawData.Contains(":") || !rawData.Contains("{")) return null;
            if (rawData.Split(':')[0].StartsWith("__")) return null;
            OtpRegisterItem otpItem = new OtpRegisterItem();
            otpItem.OtpRegisterName = rawData.Split(':')[0].Trim();
            otpItem.DefaultOrReal = "Default";
            var regSplitData = @":\s*{(?<info>.*)}$";

            var property = Regex.Match(rawData, regSplitData, RegexOptions.IgnoreCase).Groups["info"].Value;
            foreach (var pair in property.Split(','))
            {

                if (!pair.Contains(":")) continue;
                var key = pair.Split(':')[0].Trim();
                var data = pair.Split(':')[1].Trim().Replace("\'", "");
                switch (key.ToLower())
                {
                    case "name":
                        otpItem.Name = data;
                        break;
                    case "reg_name":
                        otpItem.RegName = data;
                        break;
                    case "inst_name":
                        otpItem.InstName = data;
                        break;
                    case "inst_base":
                        otpItem.InstBase = data;
                        break;
                    case "reg_ofs":
                        otpItem.RegOfs = data;
                        break;
                    case "otp_owner":
                        otpItem.OtpOwner = data;
                        break;
                    case "value":
                        otpItem.DefaultValue = data;
                        break;
                    case "bw":
                        otpItem.Bw = data;
                        break;
                    case "idx":
                        otpItem.Idx = data;
                        break;
                    case "offset":
                        otpItem.Offset = data;
                        break;
                    case "otp_b0":
                        otpItem.OtpB0 = data;
                        break;
                    case "otp_a0":
                        otpItem.OtpA0 = data;
                        break;
                    case "otpreg_add":
                        otpItem.OtpRegAdd = data;
                        break;
                    case "otpreg_ofs":
                        otpItem.OtpRegOfs = data;
                        break;
                    default:
                        continue;

                }

                if (otpItem.DefaultOrReal.Equals("Default", StringComparison.OrdinalIgnoreCase))
                {
                    if (key.Equals("name", StringComparison.OrdinalIgnoreCase) ||
                        key.Equals("reg_name", StringComparison.OrdinalIgnoreCase))
                        otpItem.DefaultOrReal = UpdateDefaultRealType(data);
                }
            }

            return otpItem;
        }

        private void UpdateHeader(List<string> fields)
        {
            Headers.Clear();
            if (fields.FirstOrDefault(p => p.Equals("OTP_REGISTER_NAME", StringComparison.OrdinalIgnoreCase)) == null)
                Headers.Insert(0, "OTP_REGISTER_NAME");
            foreach (var field in fields)
            {
                if (field.Equals("value", StringComparison.OrdinalIgnoreCase))
                    Headers.Add("DEFAULT VALUE");
                else
                    Headers.Add(field);
            }
            Headers.Add("Default or Real");
            Headers.Add("REAL VALUE");
            Headers.Add("Comment");
            Headers.Add("Different");
            Headers.Add("OTP_ECID_ONLY");
            Headers.Remove("END");
        }

        private string UpdateDefaultRealType(string name)
        {
            if (Regex.IsMatch(name, "CHIP_ID", RegexOptions.IgnoreCase)) return "Real";
            if (name.Equals("CRC", StringComparison.OrdinalIgnoreCase)) return "Real";
            if (Regex.IsMatch(name, @"OTP_SLV_OTP_ATE", RegexOptions.IgnoreCase)) return "Real";
            return "Default";
        }

        /// <summary>
        /// Only for OTP File
        /// </summary>
        /// <returns></returns>
        public string GetVersionFromOtpFileName()
        {
            var infoSubset = FileNameWithoutExtension.Split('_').ToList();
            if (infoSubset.Count < 2)
                MessageBox.Show("OTP file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);

            if (infoSubset.Count >= 2)
            {
                string version = string.Join("_", infoSubset.GetRange(infoSubset.Count - 2, 2));
                if (!Regex.IsMatch(version, @"OTP_[\w]+", RegexOptions.IgnoreCase))
                    MessageBox.Show("OTP file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);
                return version;
            }
            else
            {
                return FileNameWithoutExtension;
            }
        }

        /// <summary>
        /// Only For Yaml File
        /// </summary>
        /// <param name="otpFile"></param>
        public void MergeOtpToYaml(OtpFileReader otpFile)
        {
            string version = otpFile.GetVersionFromOtpFileName();
            Headers.Add(version);
            VersionList.Add(version);

            //var infoSubset = otpFile.Split('_').ToList();
            //if (infoSubset.Count < 2)
            //    MessageBox.Show("OTP file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);

            //if (infoSubset.Count >= 2)
            //{
            //    string version = string.Join("_", infoSubset.GetRange(infoSubset.Count - 2, 2));
            //    if (!Regex.IsMatch(version, @"OTP_[\w]+",RegexOptions.IgnoreCase))
            //        MessageBox.Show("OTP file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);
            //    Headers.Add(version);
            //    VersionList.Add(version);
            //}
            //else
            //{
            //    Headers.Add(otpFile);
            //}

            foreach (var item in otpFile.OtpRows)
            {
                var target = OtpRows.FirstOrDefault(p => p.OtpRegisterName.Equals(item.OtpRegisterName));
                if (target != null)
                    target.OtpExtra.Add(item.DefaultValue);
                else
                {
                    var errMsg = item.OtpRegisterName + "is missing !!!";
                    ErrorManager.AddError(PmicErrorType.MissingRegister, "", 1, errMsg);
                }
            }

        }
    }
}
