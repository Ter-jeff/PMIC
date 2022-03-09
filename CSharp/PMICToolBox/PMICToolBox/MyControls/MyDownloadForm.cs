using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace PmicAutomation.MyControls
{
    public partial class MyDownloadForm : Form
    {
        private Assembly _ass;

        public MyDownloadForm()
        {
            InitializeComponent();
        }

        public MyDownloadForm Download(string resourceName)
        {
            GetResourceName(resourceName);
            return this;
        }

        public MyDownloadForm DownloadFromAssembly(string resourceName,Assembly ass)
        {
            _ass = ass;
            GetResourceName(resourceName);
            return this;
        }

        private void GetResourceName(string name)
        {
            var assembly = _ass==null?Assembly.GetExecutingAssembly():_ass;
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (!resourceName.StartsWith(name)) continue;
                var file = resourceName.Replace(name, "").TrimStart('.');
                checkedListBox.Items.Add(file, true);
            }
            checkedListBox.Sorted = true;
        }

        public MyDownloadForm DownloadContains(List<string> nameList)
        {
            GetResourceNameContains(nameList);
            return this;
        }

        private void GetResourceNameContains(List<string> nameList)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            List<string> resourceList = resourceNames.ToList();
            List<string> targetList = new List<string>();
            nameList.ForEach(p => {
                targetList.Add(resourceList.Find(s => s.IndexOf(p, StringComparison.OrdinalIgnoreCase) != -1));
            });
            foreach (var obj in targetList)
            {
                var items = obj.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                var file = string.Join(".", items[items.Count()-2], items[items.Count()-1]);
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

            var assembly = _ass == null ? Assembly.GetExecutingAssembly() : _ass;
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
