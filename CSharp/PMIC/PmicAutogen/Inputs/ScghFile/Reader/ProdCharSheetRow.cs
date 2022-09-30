using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace PmicAutogen.Inputs.ScghFile.Reader
{
    [Serializable]
    public class ProdCharSheetRow : IProdCharSheetRow
    {
        public List<string> GetInitList()
        {
            return _initList;
        }

        public List<string> GetPayloadList()
        {
            return _payloadList;
        }

        public string PayloadValue
        {
            get
            {
                if (_payloadList.Count == 0)
                    return "";
                return _payloadList.First();
            }
            set { throw new NotImplementedException(); }
        }

        public List<string> GetInitAliasList()
        {
            return _initAliasList;
        }

        public List<string> GetPayloadAliasList()
        {
            return _payloadAliasList;
        }

        public string GetIndexForHardIp()
        {
            var validInitList = InitList.FindAll(s => !string.IsNullOrEmpty(s) && s != "NA" && s != "N/A").ToList();
            var validPayLoadList =
                PayloadList.FindAll(s => !string.IsNullOrEmpty(s) && s != "NA" && s != "N/A").ToList();
            if (InitList.Count > 1 || PayloadList.Count > 1)
                return (string.Join(",", validInitList) + ";" + string.Join(";", validPayLoadList)).Trim(';');
            return PayloadValue;
        }

        public string GetSourceSheet()
        {
            return SourceSheetName;
        }

        public List<string> GetAllPatternList()
        {
            var all = new List<string>();
            all.AddRange(_initList);
            all.AddRange(_payloadList);
            return all;
        }

        public bool IsBist()
        {
            return FlowName.ToUpper().Contains("BIST");
        }

        public string GetDomainByPattern()
        {
            var domain = "";
            //Organization : 'A:HARD_IP,C:CPU,L:GFX,P:HARD_IP,S:SOC,V:HARD_IP'
            //private  readonly Regex RgxOrg = new Regex(@"A|C|L|P|S|V|H", RegexOptions.IgnoreCase | RegexOptions.Compiled); //DP for Dummy Pattern

            foreach (var pattern in PatternList)
            {
                var arr = pattern.Split('_').ToList();
                if (arr.Count > 2)
                {
                    if (arr[2].Equals("A", StringComparison.CurrentCultureIgnoreCase))
                        domain = "HardIP";
                    else if (arr[2].Equals("P", StringComparison.CurrentCultureIgnoreCase))
                        domain = "HardIP";
                    else if (arr[2].Equals("V", StringComparison.CurrentCultureIgnoreCase))
                        domain = "HardIP";
                    else if (arr[2].Equals("H", StringComparison.CurrentCultureIgnoreCase))
                        domain = "HardIP";
                    else if (arr[2].Equals("C", StringComparison.CurrentCultureIgnoreCase))
                        domain = "Cpu";
                    else if (arr[2].Equals("L", StringComparison.CurrentCultureIgnoreCase))
                        domain = "Gfx";
                    else if (arr[2].Equals("S", StringComparison.CurrentCultureIgnoreCase))
                        domain = "Soc";
                }

                if (!string.IsNullOrEmpty(domain))
                    return domain;
            }

            return "";
        }

        public string GetDomainByFlowName()
        {
            var domain = "";
            if (FlowName.ToLower().Contains("Cpu".ToLower()))
                domain = "Cpu";
            else if (FlowName.ToLower().Contains("Gpu".ToLower()) || FlowName.ToLower().Contains("gpu") ||
                     FlowName.ToLower().Contains("gfx"))
                domain = "Gpu";
            else if (FlowName.ToLower().Contains("Soc".ToLower()))
                domain = "Soc";
            else if (FlowName.ToLower().Contains("Spi".ToLower())) domain = "Spi";

            return domain;
        }

        #region Field

        private List<string> _patternList;
        private List<string> _initList;
        private List<string> _payloadList;
        private List<string> _initAliasList;
        private List<string> _payloadAliasList;

        #endregion

        #region Property

        public int RowNum { set; get; }
        public string Block { set; get; }
        public string Mode { set; get; }
        public string Item { set; get; }
        public string Segment { set; get; }
        public string Inits { set; get; }
        public string PayLoads { set; get; }
        public bool IsGenFlow { set; get; }
        public string Application { set; get; }

        public List<string> PatternList
        {
            set { _patternList = value; }
            get { return _patternList; }
        }

        public List<string> InitList
        {
            set { _initList = value; }
            get { return _initList; }
        }

        public List<string> PayloadList
        {
            set { _payloadList = value; }
            get { return _payloadList; }
        }

        public List<string> InitAliasList
        {
            set { _initAliasList = value; }
            get { return _initAliasList; }
        }

        public List<string> PayloadAliasList
        {
            set { _payloadAliasList = value; }
            get { return _payloadAliasList; }
        }

        public string Usage { set; get; }
        public string SupplyVoltage { set; get; }
        public string EnableWord { set; get; }
        public string LevelHVorLv { set; get; }
        public string Comments { set; get; }
        public string PeripheralVoltage { get; set; }
        public string SramVoltage { get; set; }
        public string SourceSheetName { get; set; }
        public string FlowName { set; get; }
        public string Instance { set; get; }

        #endregion

        #region Constructor

        public ProdCharSheetRow(string sourceSheetName = "")
        {
            SourceSheetName = sourceSheetName;
            Block = "";
            Mode = "";
            Item = "";
            Segment = "";
            Inits = "";
            PayLoads = "";
            IsGenFlow = true;
            Application = "";
            _patternList = new List<string>();
            _initList = new List<string>();
            _payloadList = new List<string>();
            _initAliasList = new List<string>();
            _payloadAliasList = new List<string>();
            Usage = "";
            SupplyVoltage = "";
            EnableWord = "";
            LevelHVorLv = "";
            Comments = "";
            PeripheralVoltage = "";
            SramVoltage = "";
        }

        public ProdCharSheetRow(ProdCharSheetRow item)
        {
            SourceSheetName = item.SourceSheetName;
            Block = item.Block;
            Mode = item.Mode;
            Item = item.Item;
            Inits = "";
            PayLoads = "";
            IsGenFlow = true;
            Application = item.Application;
            _initList = item.InitList;
            _patternList = item.PatternList;
            _payloadList = item.PayloadList;
            _initAliasList = item.InitAliasList;
            _payloadAliasList = item.PayloadAliasList;
            Usage = item.Usage;
            SupplyVoltage = item.SupplyVoltage;
            EnableWord = item.EnableWord;
            LevelHVorLv = item.LevelHVorLv;
            Comments = item.Comments;
            PeripheralVoltage = item.PeripheralVoltage;
            SramVoltage = item.SramVoltage;
        }

        #endregion
    }
}