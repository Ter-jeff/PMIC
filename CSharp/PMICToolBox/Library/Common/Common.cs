using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using Library.DataStruct;

namespace Library.Common
{
    public class CommonData:INotifyPropertyChanged 
    {
        #region Field
        private static CommonData _instance = null;
        private string _outputPath = string.Empty;
        private string _baseTxtDatalogPath = string.Empty;
        private string _stdfDatalogPath = string.Empty;
        private string _compareTxtDatalogPath = string.Empty;
        private Settings _LogSettings = null;
  
        private string _progressValue = "0";
        private bool _uiEnabled = true;
        private string _uiStateInfo = string.Empty;
        public event PropertyChangedEventHandler PropertyChanged;
        public BackgroundWorker worker = null;

        #endregion

        #region Property
    
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
        public bool UIEnabled
        {
            get { return this._uiEnabled; }
            set
            {
                this._uiEnabled = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("UIEnabled"));
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
        
        public string OutputPath
        {
            get { return this._outputPath; }
            set { this._outputPath = value; }
        }

        public string BaseTxtDatalogPath
        {
            get { return this._baseTxtDatalogPath; }
            set { this._baseTxtDatalogPath = value; }
        }

        public string StdfDatalogPath
        {
            get { return this._stdfDatalogPath; }
            set { this._stdfDatalogPath = value; }
        }

        public string CompareTxtDatalogPath
        {
            get { return this._compareTxtDatalogPath; }
            set { this._compareTxtDatalogPath = value; }
        }

        public Settings LogSettings
        {
            get { return _LogSettings; }
        }
        
        #endregion

        #region Constructor
        private CommonData()
        {
            Init();
        }
        #endregion

        #region Static Function
        public static CommonData GetInstance()
        {
            if (_instance == null)
            {
                _instance = new CommonData();
            }
            return _instance;
        }
        #endregion

        #region Member Function
        
        public void Init()
        {
            this.ProgressValue = "0";
            this.UISateInfo = "";

            _LogSettings = SettingLib.ReadSetting();
            SettingLib.CreateHeaderDataPattern(_LogSettings);
            
        }

        #endregion
    }
}
