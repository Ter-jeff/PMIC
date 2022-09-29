using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace PmicAutogen.UI
{
    public class VDDRefRow : INotifyPropertyChanged
    {
        private string _domain;
        public string Domain
        {
            get { return _domain; }
            set
            {
                if (_domain != value)
                {
                    _domain = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _voltage;
        public string Voltage
        {
            get { return _voltage; }
            set
            {
                if (_voltage != value)
                {
                    _voltage = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selectRef;
        public string SelectRef
        {
            get { return _selectRef; }
            set
            {
                if (_selectRef != value)
                {
                    _selectRef = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableCollection<string> _refItems;
        public ObservableCollection<string> RefItems
        {
            get { return _refItems; }
            set
            {
                if (_refItems != value)
                {
                    _refItems = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyname = null)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }
    }
}