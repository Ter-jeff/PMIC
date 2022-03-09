using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    public partial class MyDownloadForm : Form
    {
        private Dictionary<string, string> _resourceSelectList=new Dictionary<string, string>();
        private Dictionary<string, Stream> _resources=new Dictionary<string, Stream>();

        public MyDownloadForm()
        {
            InitializeComponent();
            Font = new Font("Microsoft Sans Serif", 9F);
        }

        public MyDownloadForm Download(string name, string[] resourceNames, Assembly assembly)
        {
            foreach (var resourceName in resourceNames)
            {
                if (!resourceName.StartsWith(name))
                    continue;
                var selectName = resourceName.Replace(name, "").TrimStart('.');
                _resourceSelectList.Add(selectName, resourceName);
                _resources.Add(selectName, assembly.GetManifestResourceStream(resourceName));
            }

            foreach (var resourceName in _resourceSelectList)
                checkedListBox.Items.Add(resourceName.Key, true);
            checkedListBox.Sorted = true;
            return this;
        }

        private void SaveFolder_Click(object sender, EventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog
            {
                Description = @"Select Output Directory"
            };

            if (folderBrowserDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            foreach (var checkedItems in checkedListBox.CheckedItems)
            {
                if (_resources.ContainsKey(checkedItems.ToString()))
                {
                    using (var resource = _resources[checkedItems.ToString()])
                    {
                        int cnt = checkedItems.ToString().Split('.').Length;
                        var file = Path.Combine(folderBrowserDialog.SelectedPath,
                            string.Join(".", checkedItems.ToString().Split('.').ToList().GetRange(cnt - 2, 2)));
                        using (Stream output = File.OpenWrite(file))
                        {
                            resource.CopyTo(output);
                        }
                    }
                }
            }
            Process.Start("explorer.exe", folderBrowserDialog.SelectedPath);
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
