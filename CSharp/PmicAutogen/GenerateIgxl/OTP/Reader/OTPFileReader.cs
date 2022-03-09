using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AutomationCommon.DataStructure;
using PmicAutogen.InputPackages;

namespace PmicAutogen.GenerateIgxl.OTP.Reader
{
    public class OtpFileReader
    {
        public List<string> FileHeaders = new List<string>();
        public List<string> Headers = new List<string>();
        public List<OtpRegister> OtProws = new List<OtpRegister>();
        private string _versionRawDataRow = "";

        public OtpFileReader(string filename)
        {
            Filename = filename;
            Version = Path.GetFileNameWithoutExtension(Filename);
            var rows = GetYamlData();
            var fields = OrderedFieldNames(rows);
            UpdateHeader(fields);
            foreach (var row in rows)
            {
                var otpItem = ConvertRawDataToOtpItem(row);
                if (otpItem != null)
                    OtProws.Add(otpItem);
            }
        }

        private string Filename { get; }
        //__VERSION__: {name: APC_SERA, description: 'APC_SERA: APPLE PMU', vendor: apple.com, version: '0.1', otp_instances_nr: 1, otp_map_size: 1024, otp_data_width: 32, otp_full_size: 4096, otp_num_crc: 8, otp_dual_cell: 1, otp_double_bit: 0}
        public string VersionRawDataRow {
            get { return _versionRawDataRow; }
        }
        public string Version { get; set; }



        private List<string> OrderedFieldNames(List<string> rows)
        {
            var fieldNames = new List<string>();
            foreach (var row in rows)
            {
                if (row.StartsWith("#")) continue;
                if (row.Contains(":") && row.Contains("{"))
                {
                    if (row.Split(':')[0].StartsWith("__"))
                    {
                        _versionRawDataRow = row;
                        continue;
                    }
                    var content = row.Split('{')[1].Replace("}", "");
                    foreach (var pair in content.Split(',')) fieldNames.Add(pair.Split(':')[0].Trim());
                    fieldNames.Remove("reg_addr");
                    fieldNames.Remove("desc");
                    fieldNames.Remove("htmldesc");
                    return fieldNames;
                }
            }

            return fieldNames;
        }

        private List<string> GetYamlData()
        {
            var reader = new StreamReader(Filename);
            var contents = new List<string>();
            string line;
            var preInclude = 0;
            var postInclude = 0;
            var lineTemp = "";
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains("{"))
                    preInclude = preInclude + line.Count(x => x.Equals('{'));
                if (line.Contains("}"))
                    postInclude = postInclude + line.Count(x => x.Equals('}'));

                if (preInclude != 0)
                {
                    lineTemp = lineTemp + line;
                    if (preInclude == postInclude)
                    {
                        contents.Add(lineTemp);
                        lineTemp = "";
                        preInclude = 0;
                        postInclude = 0;
                    }
                }
            }

            return contents;
        }

        private OtpRegister ConvertRawDataToOtpItem(string rawData)
        {
            if (rawData.StartsWith("#")) return null;
            if (!rawData.Contains(":") || !rawData.Contains("{")) return null;
            if (rawData.Split(':')[0].StartsWith("__")) return null;
            var otpItem = new OtpRegister();
            otpItem.OtpRegisterName = rawData.Split(':')[0].Trim().ToUpper();
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
                        otpItem.Name = data.ToUpper();
                        break;
                    case "reg_name":
                        otpItem.RegName = data.ToUpper();
                        break;
                    case "inst_name":
                        otpItem.InstName = data.ToUpper();
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
                        otpItem.DefaultValue = data.ToUpper();
                        break;
                    case "bw":
                        otpItem.Bw = data.ToUpper();
                        break;
                    case "idx":
                        otpItem.Idx = data.ToUpper();
                        break;
                    case "offset":
                        otpItem.Offset = data.ToUpper();
                        break;
                    case "otp_b0":
                        otpItem.OtpB0 = data.ToUpper();
                        break;
                    case "otp_a0":
                        otpItem.OtpA0 = data.ToUpper();
                        break;
                    case "otpreg_add":
                        otpItem.OtpRegAdd = data.ToUpper();
                        break;
                    case "otpreg_ofs":
                        otpItem.OtpRegOfs = data.ToUpper();
                        break;
                    default:
                        continue;
                }

                if (otpItem.DefaultOrReal.Equals("Default", StringComparison.OrdinalIgnoreCase))
                    if (key.Equals("name", StringComparison.OrdinalIgnoreCase) ||
                        key.Equals("reg_name", StringComparison.OrdinalIgnoreCase))
                        otpItem.DefaultOrReal = UpdateDefaultRealType(data);
            }

            return otpItem;
        }

        private void UpdateHeader(List<string> fields)
        {
            Headers.Clear();
            if (fields.FirstOrDefault(p => p.Equals("OTP_REGISTER_NAME", StringComparison.OrdinalIgnoreCase)) == null)
                Headers.Insert(0, "OTP_REGISTER_NAME");
            foreach (var field in fields)
                Headers.Add(field.Equals("value", StringComparison.OrdinalIgnoreCase) ? "DEFAULT VALUE" : field);
            Headers.Add("Default or Real");
            Headers.Add("REAL VALUE");
            Headers.Add("Comment");
            Headers.Add("Different");
            Headers.Remove("END");
            FileHeaders.AddRange(Enumerable.Repeat("", Headers.Count));
            Headers.Add("OTP_ECID_ONLY");
            FileHeaders.Add("Default");
        }

        private string UpdateDefaultRealType(string name)
        {
            if (Regex.IsMatch(name, "CHIP_ID", RegexOptions.IgnoreCase)) return "Real";
            if (name.Equals("CRC", StringComparison.OrdinalIgnoreCase)) return "Real";
            if (Regex.IsMatch(name, @"OTP_SLV_OTP_ATE", RegexOptions.IgnoreCase)) return "Real";
            return "Default";
        }

        public void MergeToYaml(List<OtpRegister> otpRegisters, string otpFile)
        {
            var infoSubset = otpFile.Split('_').ToList();
            if (infoSubset.Count < 2)
            {
                //MessageBox.Show("OTP/Yaml file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);
                Response.Report("It is better to make your .OTP input file naming as below : ", MessageLevel.Warning, 0);
                Response.Report("OTP/Yaml file name should be end with string like 'OTP_[version]", MessageLevel.Warning, 0);
            }

            if (infoSubset.Count >= 2)
            {
                string version = string.Join("_", infoSubset.GetRange(infoSubset.Count - 2, 2));
                if(!Regex.IsMatch(version,@"OTP_[\w]+",RegexOptions.IgnoreCase))
                    //MessageBox.Show("OTP/Yaml file name should be end with string like 'OTP_[version]", "Warning", MessageBoxButtons.OK);
                    Response.Report("It is better to make your .OTP input file naming as below : ", MessageLevel.Warning, 0);
                    Response.Report("OTP/Yaml file name should be end with string like 'OTP_[version]", MessageLevel.Warning, 0);
                Headers.Add(version);
            }
            else
            {
                Headers.Add(otpFile);
            }

            FileHeaders.Add(otpFile);
            foreach (var item in otpRegisters)
            {
                var target = OtProws.FirstOrDefault(p => p.OtpRegisterName.Equals(item.OtpRegisterName));
                if (target != null)
                    target.OtpExtra.Add(item.DefaultValue);
            }
        }
    }
}