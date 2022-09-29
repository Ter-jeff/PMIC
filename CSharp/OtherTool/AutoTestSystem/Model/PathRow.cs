using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AutoTestSystem.Model
{
    public class PathRow
    {
        private bool _existIni;
        private string _name;

        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool ExistIni
        {
            get { return _existIni; }
            set
            {
                if (_existIni != value)
                {
                    _existIni = value;
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