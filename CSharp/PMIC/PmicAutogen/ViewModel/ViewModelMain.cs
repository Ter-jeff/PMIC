using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PmicAutogen.ViewModel
{
    public class ViewModelMain : INotifyPropertyChanged
    {
        private static ViewModelMain _instance { get; set; }

        public static ViewModelMain Instance()
        {
            return _instance ?? (_instance = new ViewModelMain());
        }

        private bool _btnRunAutogenIsEnabled;
        private bool _btnSettingIsEnabled;

        private bool _basicIsChecked;
        private bool _scanIsChecked;
        private bool _mbistIsChecked;
        private bool _otpIsChecked;
        private bool _vbtIsChecked;

        private bool _basicIsEnabled;
        private bool _scanIsEnabled;
        private bool _mbistIsEnabled;
        private bool _otpIsEnabled;
        private bool _vbtIsEnabled;

        public ViewModelMain()
        {
        }

        public bool BtnRunAutogenIsEnabled
        {
            get { return _btnRunAutogenIsEnabled; }
            set
            {
                _btnRunAutogenIsEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool BtnSettingIsEnabled
        {
            get { return _btnSettingIsEnabled; }
            set
            {
                _btnSettingIsEnabled = value;
                OnPropertyChanged();
            }
        }

        public bool BasicIsChecked
        {
            get { return _basicIsChecked; }
            set
            {
                _basicIsChecked = value;
                OnPropertyChanged();
            }
        }
        public bool ScanIsChecked
        {
            get { return _scanIsChecked; }
            set
            {
                _scanIsChecked = value;
                OnPropertyChanged();
            }
        }
        public bool MbistIsChecked
        {
            get { return _mbistIsChecked; }
            set
            {
                _mbistIsChecked = value;
                OnPropertyChanged();
            }
        }
        public bool OTPIsChecked
        {
            get { return _otpIsChecked; }
            set
            {
                _otpIsChecked = value;
                OnPropertyChanged();
            }
        }
        public bool VBTIsChecked
        {
            get { return _vbtIsChecked; }
            set
            {
                _vbtIsChecked = value;
                OnPropertyChanged();
            }
        }

        public bool BasicIsEnabled
        {
            get { return _basicIsEnabled; }
            set
            {
                _basicIsEnabled = value;
                OnPropertyChanged();
            }
        }
        public bool ScanIsEnabled
        {
            get { return _scanIsEnabled; }
            set
            {
                _scanIsEnabled = value;
                OnPropertyChanged();
            }
        }
        public bool MbistIsEnabled
        {
            get { return _mbistIsEnabled; }
            set
            {
                _mbistIsEnabled = value;
                OnPropertyChanged();
            }
        }
        public bool OTPIsEnabled
        {
            get { return _otpIsEnabled; }
            set
            {
                _otpIsEnabled = value;
                OnPropertyChanged();
            }
        }
        public bool VBTIsEnabled
        {
            get { return _vbtIsEnabled; }
            set
            {
                _vbtIsEnabled = value;
                OnPropertyChanged();
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public void SetButtonStatusTrue()
        {
            Instance().BasicIsEnabled = true;
            Instance().ScanIsEnabled = true;
            Instance().MbistIsEnabled = true;
            Instance().OTPIsEnabled = true;
            Instance().VBTIsEnabled = true;

            Instance().BasicIsChecked = true;
            Instance().ScanIsChecked = true;
            Instance().MbistIsChecked = true;
            Instance().OTPIsChecked = true;
            Instance().VBTIsChecked = true;
        }
    }
}