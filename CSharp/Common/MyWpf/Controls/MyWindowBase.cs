using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace MyWpf.Controls
{
    public class MyWindowBase : Window
    {
        private string _initialDirectory;
        protected DateTime StartTime;

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            DragMove();
        }

        protected string FileSelect(object sender, EnumFileFilter filter)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = new FileFilter().GetFilter(filter),
                Multiselect = false
            };

            if (!string.IsNullOrEmpty(_initialDirectory))
                openFileDialog.InitialDirectory = _initialDirectory;
            var text = ((TextBoxButton)sender).Text;
            if (string.IsNullOrEmpty(openFileDialog.FileName) && !string.IsNullOrEmpty(text))
                openFileDialog.InitialDirectory = Path.GetDirectoryName(text);

            if ((bool)openFileDialog.ShowDialog())
            {
                var name = openFileDialog.FileNames.First();
                _initialDirectory = Path.GetDirectoryName(name);
                Focus();
                ((TextBoxButton)sender).Focus();
                ((TextBoxButton)sender).Text = name;
                return name;
            }
            Focus();
            ((TextBoxButton)sender).Focus();
            return null;
        }

        protected string PathSelect(object sender)
        {
            if (!string.IsNullOrEmpty(((TextBoxButton)sender).Text))
                _initialDirectory = ((TextBoxButton)sender).Text;
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = _initialDirectory,
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                ((TextBoxButton)sender).Focus();
                return null;
            }

            ((TextBoxButton)sender).Focus();
            ((TextBoxButton)sender).Text = folderBrowserDialog.FileName;
            _initialDirectory = folderBrowserDialog.FileName;
            return folderBrowserDialog.FileName;
        }

        public class FileFilter
        {
            public string GetFilter(EnumFileFilter filter)
            {
                if (filter == EnumFileFilter.TestProgram)
                {
                    return "(*.igxl*,*.xls*)|*.xls*;*.igxl*";
                }

                if (filter == EnumFileFilter.Igxl)
                {
                    return "(*.igxl*)|*.igxl*";
                }

                if (filter == EnumFileFilter.PatternCsv)
                {
                    return "(*pattern*.csv)|*pattern*.csv";
                }

                if (filter == EnumFileFilter.Excel)
                {
                    return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";
                }

                if (filter == EnumFileFilter.IdsDistribution)
                {
                    return @"IDS Distribution" + @"(IDS_Distribution.txt)|IDS_Distribution.txt*";
                }

                if (filter == EnumFileFilter.TestPlan)
                {
                    return @"TestPlan" + @"(*test*plan*.xlsx*)|*test*plan*.xlsx*";
                }

                if (filter == EnumFileFilter.BinCut)
                {
                    return @"*Bin*Cut*" + @"(*.txt,*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx";
                }

                if (filter == EnumFileFilter.BinCutPost)
                {
                    return @"*Post*Bin*Cut*" + @"(*.txt,*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx";
                }

                if (filter == EnumFileFilter.TemplateFile)
                {
                    return "(*.tmp)|*.tmp";
                }

                if (filter == EnumFileFilter.Txt)
                {
                    return "(*.txt)|*.txt";
                }

                if (filter == EnumFileFilter.BasFile)
                {
                    return "(*.bas)|*.bas";
                }

                if (filter == EnumFileFilter.PaFile)
                {
                    return "(*.csv,*.xls*)|*.csv;*.xlsx;*.xlsm";
                }

                if (filter == EnumFileFilter.XmlFile)
                {
                    return "(*.xml)|*.xml";
                }

                if (filter == EnumFileFilter.BdFile)
                {
                    return "(*.csv)|*.csv";
                }

                return "";
            }
        }
    }
}
