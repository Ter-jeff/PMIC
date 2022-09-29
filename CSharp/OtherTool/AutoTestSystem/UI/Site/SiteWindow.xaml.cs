using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AutoTestSystem.UI.Site
{
    public partial class SiteWindow
    {
        public SiteWindow()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void ListBox_KeyDown(object sender, KeyEventArgs e)
        {
            var list = sender as ListBox;
            var letter = e.Key.ToString();
            if (list != null)
            {
                var index = list.Items.SourceCollection.Cast<SiteRow>().ToList()
                    .FindIndex(x => x.Site.StartsWith(letter, StringComparison.CurrentCultureIgnoreCase));
                if (index == -1)
                    return;
                list.SelectedIndex = index;
                list.ScrollIntoView(list.Items[index]);
            }
            e.Handled = true;
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in ListBox.ItemsSource)
            {
                var enableRow = (SiteRow)item;
                enableRow.Select = false;
            }
        }
    }
}
