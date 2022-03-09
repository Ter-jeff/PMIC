using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace AutomationCommon.Controls
{
    public partial class MyDownloadForm : Form
    {

        public MyDownloadForm()
        {
            InitializeComponent();
        }

        public MyDownloadForm Download(string resourceName)
        {
            GetResourceName(resourceName);
            return this;
        }

        private void GetResourceName(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (!resourceName.StartsWith(name)) continue;
                var file = resourceName.Replace(name, "").TrimStart('.');
                checkedListBox.Items.Add(file, true);
            }
            checkedListBox.Sorted = true;
        }

        public MyDownloadForm DownloadContains(string resourceName)
        {
            GetResourceNameContains(resourceName);
            return this;
        }

        private void GetResourceNameContains(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (!resourceName.ToLower().Contains(name.ToLower())) continue;
                var file = resourceName.Substring(resourceName.ToLower().IndexOf(name, StringComparison.CurrentCultureIgnoreCase)+ name.Length);
                checkedListBox.Items.Add(file, true);
            }
            checkedListBox.Sorted = true;
        }

        private void SaveFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
            {
                Description = @"Select Output Directory"
            };

            if (folderBrowserDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var checkedItems in checkedListBox.CheckedItems)
            {
                foreach (var resourceName in resourceNames)
                {
                    if (resourceName.Contains(checkedItems.ToString()))
                    {
                        using (var resource = assembly.GetManifestResourceStream(resourceName))
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
            }
            Process.Start("explorer.exe", folderBrowserDialog.SelectedPath);
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
