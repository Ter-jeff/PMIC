using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using IgxlData.IgxlManager;

namespace IgxlEditor.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<SheetTypeRow> _sheetTypeRows;

        private string _editText;

        public MainViewModel()
        {
            SheetTypeRows = new ObservableCollection<SheetTypeRow>();
        }

        public SheetTypeRow SelectRow { get; set; }

        public ObservableCollection<SheetTypeRow> SheetTypeRows
        {
            get { return _sheetTypeRows; }
            set
            {
                if (_sheetTypeRows != value)
                {
                    _sheetTypeRows = value;
                    OnPropertyChanged();
                }
            }
        }

        public string EditText
        {
            get { return _editText; }
            set
            {
                if (_editText != value)
                {
                    _editText = value;
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