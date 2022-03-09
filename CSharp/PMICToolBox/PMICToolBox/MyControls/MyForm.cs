using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PmicAutomation.MyControls
{
    public enum FileFilter
    {
        Excel,
        TemplateFile,
        BasFile,
        PaFile,
        XmlFile,
        IgxlFile,
        YamlFile,
        OtpFile
    }

    public class MyForm : Form
    {
        protected PmicAutomation.MyControls.MyStatus MyStatus;
        protected DateTime StartTime;
        public delegate void RichTextBoxAppend(string message, Color color);
        protected string DefaultPath;

        public MyForm()
        {
            InitializeComponent();
        }

        private string GetFilter(FileFilter filter)
        {
            if (filter == FileFilter.Excel)
            {
                return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";
            }

            if (filter == FileFilter.TemplateFile)
            {
                return "(*.tmp)|*.tmp";
            }

            if (filter == FileFilter.BasFile)
            {
                return "(*.bas,*.txt)|*.bas;*.txt";
            }

            if (filter == FileFilter.PaFile)
            {
                return "(*.csv,*.xls*)|*.csv;*.xlsx;*.xlsm";
            }

            if (filter == FileFilter.XmlFile)
            {
                return "(*.xml)|*.xml";
            }

            if (filter == FileFilter.IgxlFile)
            {
                return "(*.igxl)|*.igxl";
            }
            if (filter == FileFilter.YamlFile)
            {
                return "(*.yaml)|*.yaml";
            }
            if (filter == FileFilter.OtpFile)
            {
                return "(*.otp)|*.otp";
            }
            return "";
        }

        protected string FileDialog(object sender, FileFilter filter, bool multiSelect = false)
        {
            OpenFileDialog dialog = new OpenFileDialog { Filter = GetFilter(filter), Multiselect = multiSelect };
            if (!string.IsNullOrEmpty(DefaultPath))
            {
                dialog.FileName = DefaultPath;
            }

            if (dialog.ShowDialog() != DialogResult.OK)
            {
                return null;
            }
            if (!dialog.FileNames.Any())
            {
                return null;
            }

            DefaultPath = Path.GetDirectoryName(dialog.FileNames.First());
            var fileNames = multiSelect ?
               string.Join(",", dialog.FileNames) :
               dialog.FileNames.First();
            ((Control)sender).Parent.Text = fileNames;
            return fileNames;
        }

        protected string PathDialog(object sender, bool isFolderPicker, bool multiSelect = false)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = DefaultPath;
            dialog.IsFolderPicker = isFolderPicker;
            dialog.Multiselect = multiSelect;
            if (dialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                return null;
            }
            DefaultPath = dialog.FileNames.First();
            var fileNames = multiSelect ?
                string.Join(",", dialog.FileNames) :
                dialog.FileNames.First();
            ((Control)sender).Parent.Text = fileNames;
            return fileNames;
        }

        private void InitializeComponent()
        {
            this.MyStatus = new PmicAutomation.MyControls.MyStatus();
            this.SuspendLayout();
            // 
            // MyStatus
            // 
            this.MyStatus.Location = new System.Drawing.Point(0, 231);
            this.MyStatus.Name = "MyStatus";
            this.MyStatus.Size = new System.Drawing.Size(538, 22);
            this.MyStatus.TabIndex = 0;
            this.MyStatus.Text = "MyStatus";
            // 
            // MyForm
            // 
            this.ClientSize = new System.Drawing.Size(538, 253);
            this.Controls.Add(this.MyStatus);
            this.Name = "MyForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        protected void CalculateTime()
        {
            MyStatus.LabelProcessTime.Text = (DateTime.Now - StartTime).ToString(@"hh\:mm\:ss");
            MyStatus.LabelStatus.Text = "Done!";
            MyStatus.ProgressBar.Value =  0;
        }

    }
}