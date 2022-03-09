using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CommonLib.Controls
{
    public partial class MyForm : Form
    {
        public delegate void RichTextBoxAppend(string message, Color color);
        public string DefaultPath;
        public DateTime StartTime;
        private decimal _delaycount;

        public MyForm()
        {
            InitializeComponent();
            Font = new Font("Microsoft Sans Serif", 9F);
        }

        protected string FileSelect(object sender, EnumFileFilter filter)
        {
            using (var openFileDialog = new OpenFileDialog { Filter = new FileFilter().GetFilter(filter), Multiselect = false })
            {
                if (!string.IsNullOrEmpty(DefaultPath))
                    openFileDialog.FileName = DefaultPath;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ((Control)sender).Parent.Text = openFileDialog.FileNames.First();
                    DefaultPath = Path.GetDirectoryName(openFileDialog.FileNames.First());
                    return openFileDialog.FileNames.First();
                }
                return null;
            }
        }

        protected List<string> FileMultiSelect(object sender, EnumFileFilter filter)
        {
            using (var openFileDialog = new OpenFileDialog { Filter = new FileFilter().GetFilter(filter), Multiselect = true })
            {
                if (!string.IsNullOrEmpty(DefaultPath))
                    openFileDialog.FileName = DefaultPath;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ((Control)sender).Parent.Text = string.Join(",", openFileDialog.FileNames);
                    DefaultPath = Path.GetDirectoryName(openFileDialog.FileNames.First());
                    return openFileDialog.FileNames.ToList();
                }
                return null;
            }
        }

        protected string PathSelect(object sender)
        {
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = DefaultPath,
                Title = @"Select Directory of Setting Folder"
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                ((Control)sender).Parent.Focus();
                return null;
            }

            ((Control)sender).Parent.Focus();
            ((Control)sender).Parent.Text = folderBrowserDialog.FileName;
            DefaultPath = folderBrowserDialog.FileName;
            return folderBrowserDialog.FileName;
        }

        protected void Download(string templatePath)
        {
            var folderBrowserDialog = new FolderBrowserDialog
            {
                Description = @"Select Output Directory"
            };

            if (DefaultPath != "")
            {
                folderBrowserDialog.SelectedPath = DefaultPath;
            }

            if (folderBrowserDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            foreach (var dirPath in Directory.GetDirectories(templatePath, "*",
                SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(templatePath, folderBrowserDialog.SelectedPath));
            }

            foreach (var newPath in Directory.GetFiles(templatePath, "*.*",
                SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(templatePath, folderBrowserDialog.SelectedPath), true);
            }

            Process.Start("explorer.exe", folderBrowserDialog.SelectedPath);
        }

        protected bool IsOpened(string filePath)
        {
            if (!File.Exists(filePath)) return false;
            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        protected void Reset()
        {
            richTextBox.Clear();
            ToolStripStatusLabel.Text = @"Status";
            ProcessTimeToolStripStatusLabel.Text = @"Process Time";
        }

        protected void CalculateTimeStart()
        {
            StartTime = DateTime.Now;
            ToolStripStatusLabel.Text = @"Status";
            ProcessTimeToolStripStatusLabel.Text = @"Process Time";
            ToolStripProgressBar.Value = 0;
            myStatus.Refresh();
        }

        protected void CalculateTimeStop()
        {
            ProcessTimeToolStripStatusLabel.Text = (DateTime.Now - StartTime).ToString(@"hh\:mm\:ss");
            ToolStripProgressBar.Value = 0;
            myStatus.Refresh();
        }

        public void ProgressBarIncrement()
        {
            if (ToolStripProgressBar.Value == ToolStripProgressBar.Maximum)
            {
                _delaycount++;
                if (_delaycount > 1)
                {
                    ToolStripProgressBar.Value = 0;
                    _delaycount = 0;
                }
            }
            ToolStripProgressBar.Increment(1);
            myStatus.Refresh();
        }

        public void SetStatusLabel(string text)
        {
            ToolStripStatusLabel.Text = text;
            myStatus.Refresh();
        }

        public void SetProgressBarValue(int value)
        {
            if (ToolStripProgressBar.ProgressBar != null)
            {
                ToolStripProgressBar.ProgressBar.Value = value;
                ToolStripProgressBar.ProgressBar.Refresh();
            }
            myStatus.Refresh();
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
            richTextBox.Refresh();
        }
    }
}
