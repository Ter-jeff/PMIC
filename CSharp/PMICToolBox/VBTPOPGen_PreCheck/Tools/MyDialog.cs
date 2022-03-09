using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VBTPOPGenPreCheckBusiness.DataStore;

namespace VBTPOPGen_PreCheck.Tools
{
    public class MyDialog
    {
        private static MyDialog _instance = null;
        private static string DefaultPath;

        private MyDialog()
        {
            DefaultPath = string.Empty;
        }

        public static MyDialog GetInstance()
        {
            if (_instance == null)
                _instance = new MyDialog();
            return _instance;
        }

        public string GetFilter(FileFilter filter)
        {
            if (filter == FileFilter.Excel)
                return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";

            if (filter == FileFilter.TemplateFile)
                return "(*.tmp)|*.tmp";

            if (filter == FileFilter.BasFile)
                return "(*.bas,*.txt)|*.bas;*.txt";

            if (filter == FileFilter.PaFile)
                return "(*.csv,*.xls*)|*.csv;*.xlsx;*.xlsm";

            if (filter == FileFilter.XmlFile)
                return "(*.xml)|*.xml";

            if (filter == FileFilter.IgxlFile)
                return "(*.igxl)|*.igxl";

            return "";
        }

        public string FileDialog(FileFilter filter, bool multiSelect = false)
        {
            OpenFileDialog dialog = new OpenFileDialog { Filter = GetFilter(filter), Multiselect = multiSelect };
            if (!string.IsNullOrEmpty(DefaultPath))
                dialog.FileName = DefaultPath;

            dialog.ShowDialog();

            if (!dialog.FileNames.Any())
                return null;

            DefaultPath = Path.GetDirectoryName(dialog.FileNames.First());
            var fileNames = multiSelect ? string.Join(",", dialog.FileNames) : dialog.FileNames.First();
            return fileNames;
        }

        public string PathDialog(bool isFolderPicker, bool multiSelect = false)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = DefaultPath;
            dialog.IsFolderPicker = isFolderPicker;
            dialog.Multiselect = multiSelect;

            if (dialog.ShowDialog() == CommonFileDialogResult.Cancel)
                return null;

            DefaultPath = dialog.FileNames.First();
            var fileNames = multiSelect ? string.Join(",", dialog.FileNames) : dialog.FileNames.First();
            return fileNames;
        }
    }
}
