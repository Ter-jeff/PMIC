using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace CurrentProfileTool.Controls
{
    public enum FileFilter
    {
        Excel,
        TemplateFile,
        BasFile,
        PaFile,
        XmlFile,
        PatternCsv,
        TestProgram,
        Txt,
        Igxl
    }

    public class MyForm : Form
    {
        public delegate void RichTextBoxAppend(string message, Color color);
        protected string DefaultPath;
        protected DateTime StartTime;

        protected string GetFilter(FileFilter filter)
        {
            if (filter == FileFilter.TestProgram)
            {
                return "(*.igxl*,*.xls*)|*.xls*;*.igxl*";
            }

            if (filter == FileFilter.Igxl)
            {
                return "(*.igxl*)|*.igxl*";
            }

            if (filter == FileFilter.PatternCsv)
            {
                return "(*pattern*.csv)|*pattern*.csv";
            }

            if (filter == FileFilter.Excel)
            {
                return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";
            }

            if (filter == FileFilter.TemplateFile)
            {
                return "(*.tmp)|*.tmp";
            }

            if (filter == FileFilter.Txt)
            {
                return "(*.txt)|*.txt";
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

            return "";
        }

        protected string FileSelect(object sender, FileFilter filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = GetFilter(filter), Multiselect = false };
            if (!string.IsNullOrEmpty(DefaultPath))
            {
                openFileDialog.FileName = DefaultPath;
            }

            openFileDialog.ShowDialog();
            if (!openFileDialog.FileNames.Any())
            {
                return null;
            }

            ((Control)sender).Parent.Text = openFileDialog.FileNames.First();
            DefaultPath = Path.GetDirectoryName(openFileDialog.FileNames.First());
            return openFileDialog.FileNames.First();
        }

        protected List<string> FileMultiSelect(object sender, FileFilter filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = GetFilter(filter), Multiselect = true };
            if (!string.IsNullOrEmpty(DefaultPath))
            {
                openFileDialog.FileName = DefaultPath;
            }

            openFileDialog.ShowDialog();
            if (!openFileDialog.FileNames.Any())
            {
                return null;
            }

            ((Control)sender).Parent.Text = string.Join(",", openFileDialog.FileNames);
            DefaultPath = Path.GetDirectoryName(openFileDialog.FileNames.First());
            return openFileDialog.FileNames.ToList();
        }

        protected string PathSelect(object sender)
        {
            CommonOpenFileDialog folderBrowserDialog = new CommonOpenFileDialog
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
            DefaultPath = Path.GetDirectoryName(folderBrowserDialog.FileName);
            return folderBrowserDialog.FileName;
        }

        protected void Download(string templatePath)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
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

            foreach (string dirPath in Directory.GetDirectories(templatePath, "*",
                SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(templatePath, folderBrowserDialog.SelectedPath));
            }

            foreach (string newPath in Directory.GetFiles(templatePath, "*.*",
                SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(templatePath, folderBrowserDialog.SelectedPath), true);
            }

            Process.Start("explorer.exe", folderBrowserDialog.SelectedPath);
        }
    }
}