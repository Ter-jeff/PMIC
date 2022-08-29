using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MyWpf.Model
{
    public class HelpFileRow : INotifyPropertyChanged
    {
        private string _fileName;

        private bool _select;

        public string FileName
        {
            get { return _fileName; }
            set
            {
                if (_fileName != value)
                {
                    _fileName = value;
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