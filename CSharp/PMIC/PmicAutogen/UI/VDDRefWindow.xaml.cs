using PmicAutogen.Inputs.TestPlan.Reader;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace PmicAutogen.UI
{
    public partial class VDDRefWindow : Window
    {
        private readonly Dictionary<string, string> _domainDic = new Dictionary<string, string>();
        private readonly Dictionary<string, VddLevelsRow> _vddPinDic = new Dictionary<string, VddLevelsRow>();
        private ObservableCollection<VDDRefRow> vddRefRows;

        public VDDRefWindow(Dictionary<string, string> domainDic, Dictionary<string, VddLevelsRow> vddPinDic)
        {
            InitializeComponent();

            _domainDic = domainDic;
            _vddPinDic = vddPinDic;

            vddRefRows = new ObservableCollection<VDDRefRow>();
            for (var i = 0; i < _domainDic.Count; i++)
            {
                var pinInfo = _domainDic.ElementAt(i);
                var items = new ObservableCollection<string>();
                items.Add("No Reference");
                _vddPinDic.Keys.ToList().ForEach(x => items.Add(x));
                vddRefRows.Add(new VDDRefRow()
                {
                    Domain = pinInfo.Key,
                    Voltage = pinInfo.Value,
                    RefItems = items,
                    SelectRef = "No Reference",
                });
            }
            DataGrid.ItemsSource = vddRefRows;
        }

        public Dictionary<string, VddLevelsRow> RefVddPins { get; set; }

        private bool CheckMapping(string pinVolt, VddLevelsRow vddLevelRow)
        {
            if (vddLevelRow.Nv != vddLevelRow.Lv ||
                vddLevelRow.Nv != vddLevelRow.Hv ||
                (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
                return true;

            if (vddLevelRow.Nv != pinVolt) return false;
            return true;
        }

        public void Click_Ok()
        {
            RefVddPins = new Dictionary<string, VddLevelsRow>();
            foreach (var row in vddRefRows)
            {
                if (row.SelectRef != "No Reference")
                {
                    var ioPin = row.Domain;
                    var vddLevelRow = _vddPinDic[row.SelectRef];
                    if (vddLevelRow.Nv != vddLevelRow.Lv ||
                        vddLevelRow.Nv != vddLevelRow.Hv ||
                        (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                        (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
                        RefVddPins.Add(ioPin.ToString(), vddLevelRow);
                }
            }
            DialogResult = true;
        }

        private void RefVdd_Click(object sender, RoutedEventArgs e)
        {
            Click_Ok();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = (VDDRefRow)DataGrid.CurrentItem;
            if (item == null)
                return;
            var pin = item.SelectRef;
            if (pin == "No Reference")
                return;

            var pinVolt = item.Voltage;
            var vddLevelRow = _vddPinDic[pin];

            if (!CheckMapping(pinVolt, vddLevelRow))
            {
                var msg = string.Format("Voltage Not Match!\n  IO Pin {0} : {1}\n Vdd Pin {2} : {3}.",
                    item.Domain, pinVolt, pin, vddLevelRow.Nv);
                MessageBox.Show(msg, "Warnning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
