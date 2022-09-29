using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AutoTestSystem.UI.Enable
{
    public class EnableRow : INotifyPropertyChanged
    {
        private string _enableWord;

        private bool _select;

        public string EnableWord
        {
            get { return _enableWord; }
            set
            {
                if (_enableWord != value)
                {
                    _enableWord = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool Select
        {
            get { return _select; }
            set
            {
                if (_select != value)
                {
                    _select = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyname = null)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
        }
    }
}