using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace MyWpf.Controls
{
    /// <summary>
    /// Interaction logic for HelpWindow.xaml
    /// </summary>
    public partial class SheetListWindow : Window
    {
        private readonly ObservableCollection<SheetRow> _itemsSource;

        public SheetListWindow(ObservableCollection<SheetRow> sheetRows)
        {
            _itemsSource = sheetRows;
            InitializeComponent();

            DataGrid.ItemsSource = sheetRows;
        }

        public List<string> SelectItems
        {
            get { return _itemsSource.Where(x => x.Select).Select(x => x.SheetName).ToList(); }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class SheetRow : INotifyPropertyChanged
    {
        private string _sheetName;

        public string SheetName
        {
            get { return _sheetName; }
            set
            {
                if (_sheetName != value)
                {
                    _sheetName = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _select;

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
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }
    }
}
