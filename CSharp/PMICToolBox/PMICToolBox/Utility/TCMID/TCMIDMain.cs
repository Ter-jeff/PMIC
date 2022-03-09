using PmicAutomation.MyControls;
using PmicAutomation.Utility.TCMID.Business;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Util = Library.Common.Utility;
using LIMIT_FILE_TYPE = PmicAutomation.Utility.TCMID.DataStructure.EnumStore.LIMIT_FILE_TYPE;
using System.IO;
using System;
using PmicAutomation.Utility.TCMIDComparator;

namespace PmicAutomation.Utility.TCMID
{
    public class TCMIDMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _inputPath;
        private readonly string _outputPath;
        private DataTable _limitDT;
        protected LIMIT_FILE_TYPE _limit_file_type;
        private readonly string _tpVersion;
        private string _skipItemFile = Properties.Settings.Default.SkipItemConfig;
        private List<string> _skipLines;

        public TCMIDMain(TCMIDForm tcmIDForm)
        {
            _appendText = tcmIDForm.AppendText;
            _inputPath = tcmIDForm.inputPath.ButtonTextBox.Text;
            _outputPath = tcmIDForm.outputPath.ButtonTextBox.Text;
            _limit_file_type = LIMIT_FILE_TYPE.OTHERS;
            _tpVersion = tcmIDForm.GetTPVersion();
            _skipLines = new List<string>();

            if (tcmIDForm.skipItem.Checked)
                SetupSkipItem();
        }

        public TCMIDMain(TCMIDComparatorForm tcmIDComparatorForm)
        {
            _appendText = tcmIDComparatorForm.AppendText;
            _outputPath = tcmIDComparatorForm.outputPath.ButtonTextBox.Text;
            _limit_file_type = LIMIT_FILE_TYPE.OTHERS;
            _tpVersion = tcmIDComparatorForm.GetTPVersion();
            _skipLines = new List<string>();

            if (tcmIDComparatorForm.skipItem.Checked)
                SetupSkipItem();
        }

        public LIMIT_FILE_TYPE Limit_file_type
        {
            get { return _limit_file_type; }
        }

        private void SetupSkipItem()
        {
            if (File.Exists(_skipItemFile))
                _skipLines = File.ReadAllLines(_skipItemFile).ToList();
        }

        public void WorkFlow(List<string> inputFiles, List<TcmIDGenBase> tcmIdObjList = null, bool bCompare = false, bool bGenFlag = true)
        {
            foreach (var inputFile in inputFiles)
            {
                Dictionary<string, int> dicHeaderIndex = new Dictionary<string, int>();

                ReadLimitSheet(inputFile);
                _appendText(string.Format("Read limit file {0} done", inputFile), Color.ForestGreen);

                GetHeaderPosition(dicHeaderIndex);
                _appendText("Get necessary information done", Color.DimGray);

                JudgeLimitType(dicHeaderIndex["idxTestname"]);
                _appendText(string.Format("Limit sheet type: {0}", _limit_file_type), Color.DimGray);

                TcmIDGenBase obj = TcmIDFactory.GetInstance().GetTcmIDObject(_limit_file_type);
                obj.SetParameter(_limit_file_type, _limitDT, inputFile, _outputPath, dicHeaderIndex, _tpVersion, _skipLines);
                obj.Gen(bCompare, bGenFlag);
                if (tcmIdObjList != null)
                    tcmIdObjList.Add(obj);

                _appendText("Generate limit sheet and report done" + Environment.NewLine, Color.DimGray);
            }
        }

        private void JudgeLimitType(int idxTestname)
        {
            List<DataRow> dtList = _limitDT.Rows.Cast<DataRow>().ToList();
            if (dtList.Exists(p => Regex.IsMatch(p[idxTestname].ToString(), @"^open-*\w*_|^short-*\w*_", RegexOptions.IgnoreCase)))
                _limit_file_type = LIMIT_FILE_TYPE.CONTI;
            else if (dtList.Exists(p => Regex.IsMatch(p[idxTestname].ToString(), @"^IDS-*\w*_|^Pull-*\w*_", RegexOptions.IgnoreCase)))
                _limit_file_type = LIMIT_FILE_TYPE.IDS;
            else if (dtList.Exists(p => Regex.IsMatch(p[idxTestname].ToString(), @"^IIL-*\w*_|^IIH-*\w*_", RegexOptions.IgnoreCase)))
                _limit_file_type = LIMIT_FILE_TYPE.LEAKAGE;
            else
                _limit_file_type = LIMIT_FILE_TYPE.OTHERS;
        }

        private void GetHeaderPosition(Dictionary<string, int> dicHeaderIndex)
        {
            List<DataRow> dtList = _limitDT.Rows.Cast<DataRow>().ToList();
            foreach (DataRow dr in dtList)
            {
                int count = 0;
                int index = 0;
                foreach (var item in dr.ItemArray)
                {
                    switch (item.ToString().Trim().ToLower())
                    {
                        case "flowtable":
                            dicHeaderIndex.Add("idxFlowtable", index);
                            ++count;
                            break;
                        case "testname":
                            dicHeaderIndex.Add("idxTestname", index);
                            ++count;
                            break;
                        case "petname":
                            dicHeaderIndex.Add("idxPETname", index);
                            ++count;
                            break;
                        case "scale":
                            dicHeaderIndex.Add("idxScale", index);
                            ++count;
                            break;
                        case "units":
                            if (dtList.Exists(p => p[index].ToString().ToUpper().Equals("CP1")))
                                dicHeaderIndex.Add("idxUnits", index);
                            ++count;
                            break;
                        case "lolim":
                            if (dtList.Exists(p => p[index].ToString().ToUpper().Equals("CP1")))
                                dicHeaderIndex.Add("idxLowlim", index);
                            ++count;
                            break;
                        case "hilim":
                            if (dtList.Exists(p => p[index].ToString().ToUpper().Equals("CP1")))
                                dicHeaderIndex.Add("idxHilim", index);
                            ++count;
                            break;
                    }
                    ++index;
                }
                if (count != 0 && dicHeaderIndex.Count == count)
                    break;
            }
        }

        private void ReadLimitSheet(string inputFile)
        {
            _limitDT = Util.ConvertToDataTable(inputFile, new char[] { '\t' }, "flowtable");
        }
    }
}
