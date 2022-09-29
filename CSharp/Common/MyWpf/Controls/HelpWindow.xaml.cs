using Microsoft.WindowsAPICodePack.Dialogs;
using MyWpf.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;

namespace MyWpf.Controls
{
    /// <summary>
    /// Interaction logic for HelpWindow.xaml
    /// </summary>
    public partial class HelpWindow : Window
    {
        public ObservableCollection<HelpFileRow> HelpFileRows = new ObservableCollection<HelpFileRow>();
        private Dictionary<string, Stream> _resources = new Dictionary<string, Stream>();

        public HelpWindow()
        {
            InitializeComponent();
        }

        public HelpWindow Download(string name, string[] resourceNames, Assembly assembly)
        {
            var resourceSelectList = new Dictionary<string, string>();
            foreach (var resourceName in resourceNames)
            {
                if (!resourceName.StartsWith(name))
                    continue;
                var selectName = resourceName.Replace(name, "").TrimStart('.');
                resourceSelectList.Add(selectName, resourceName);
                _resources.Add(selectName, assembly.GetManifestResourceStream(resourceName));
            }

            HelpFileRows = new ObservableCollection<HelpFileRow>();
            foreach (var resourceName in resourceSelectList)
                HelpFileRows.Add(new HelpFileRow() { FileName = resourceName.Key, Select = true });
            DataGrid.ItemsSource = HelpFileRows;
            return this;
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;

            if (dialog.ShowDialog() == CommonFileDialogResult.Cancel) return;

            foreach (HelpFileRow item in DataGrid.Items)
            {
                if (item.Select)
                {
                    var fileName = item.FileName;
                    if (_resources.ContainsKey(fileName))
                    {
                        using (var resource = _resources[fileName])
                        {
                            var file = Path.Combine(dialog.FileName, fileName);
                            using (Stream output = File.OpenWrite(file))
                                resource.CopyTo(output);
                        }
                    }
                }
            }
            Process.Start("explorer.exe", dialog.FileName);
            Close();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
