using IgxlData.IgxlReader;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace IgxlData.IgxlManager
{
    public class SheetTypeRow
    {
        private string _fullname;
        private string _name;
        private string _type;
        private SheetTypes _sheetType;

        public string FullName
        {
            get { return _fullname; }
            set
            {
                if (_fullname != value)
                {
                    _fullname = value;
                    OnPropertyChanged();
                }
            }
        }

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

        public string Type
        {
            get { return _type; }
            set
            {
                if (_type != value)
                {
                    _type = value;
                    OnPropertyChanged();
                }
            }
        }

        public SheetTypes SheetType
        {
            get { return _sheetType; }
            set
            {
                if (_sheetType != value)
                {
                    _sheetType = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}