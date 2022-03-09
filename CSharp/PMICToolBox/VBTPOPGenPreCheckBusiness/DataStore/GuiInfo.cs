using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBTPOPGenPreCheckBusiness.DataStore
{
    public class GuiInfo : INotifyPropertyChanged
    {
        private static GuiInfo _instance = null;
        private string _testplanFolder;
        private List<string> _testplanList = null;
        private string _otpRegMap;
        private string _ahbRegMap;
        private string _pinMap;
        private string _output;
        private string _outputFilename;

        private string _progressValue = "0";
        private string _uiStateInfo = string.Empty;

        public event PropertyChangedEventHandler PropertyChanged;

        public static GuiInfo GetInstance()
        {
            if (_instance == null)
                _instance = new GuiInfo();
            return _instance;
        }

        private GuiInfo()
        {
            ClearVariable();
        }

        public void SetOutputFilename(string name)
        {
            _outputFilename = name;
        }

        public void ClearVariable()
        {
            if (_testplanList != null)
                _testplanList.Clear();
            else
                _testplanList = new List<string>();

            _testplanFolder = string.Empty;
            _otpRegMap = string.Empty;
            _ahbRegMap = string.Empty;
            _pinMap = string.Empty;
            _output = string.Empty;
            _outputFilename = string.Empty;
            _progressValue = "0";
            _uiStateInfo = string.Empty;

    }

        public string ProgressValue
        {
            get { return this._progressValue; }
            set
            {
                this._progressValue = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("ProgressValue"));
                }
            }
        }
        public string UISateInfo
        {
            get { return this._uiStateInfo; }
            set
            {
                this._uiStateInfo = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("UISateInfo"));
                }
            }
        }

        public string output
        {
            get { return _output; }
            set { _output = value; }
        }

        public string otpRegMap
        {
            get { return _otpRegMap; }
            set { _otpRegMap = value; }
        }

        public string AhbRegMap
        {
            get { return _ahbRegMap; }
            set { _ahbRegMap = value; }
        }

        public string PinMap
        {
            get { return _pinMap; }
            set { _pinMap = value; }
        }

        public string testplanFolder
        {
            get { return _testplanFolder; }
            set { _testplanFolder = value; }
        }

        public List<string> testplanList
        {
            get { return _testplanList; }
            set { _testplanList = value; }
        }

        public string outputFilename
        {
            get { return _outputFilename; }
        }
    }
}
