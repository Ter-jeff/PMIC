using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace MyWpf.Controls
{
    public class MyWindow : Window
    {
        private const uint WS_EX_CONTEXTHELP = 0x00000400;
        private const uint WS_MINIMIZEBOX = 0x00020000;
        private const uint WS_MAXIMIZEBOX = 0x00010000;
        private const int GWL_STYLE = -16;
        private const int GWL_EXSTYLE = -20;
        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOMOVE = 0x0002;
        private const int SWP_NOZORDER = 0x0004;
        private const int SWP_FRAMECHANGED = 0x0020;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_CONTEXTHELP = 0xF180;
        protected string _initialDirectory;

        public event EventHandler HelpButtonClick;

        [DllImport("user32.dll")]
        private static extern uint GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hwnd, int index, uint newStyle);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int width,
            int height, uint flags);

        protected override void OnSourceInitialized(EventArgs e)
        {
            var brushConverter = new BrushConverter();
            Background = (Brush)brushConverter.ConvertFromString("#AFE5FF");
            base.OnSourceInitialized(e);
            var hwnd = new WindowInteropHelper(this).Handle;
            var styles = GetWindowLong(hwnd, GWL_STYLE);
            styles &= 0xFFFFFFFF ^ (WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
            SetWindowLong(hwnd, GWL_STYLE, styles);
            styles = GetWindowLong(hwnd, GWL_EXSTYLE);
            styles |= WS_EX_CONTEXTHELP;
            SetWindowLong(hwnd, GWL_EXSTYLE, styles);
            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
            ((HwndSource)PresentationSource.FromVisual(this)).AddHook(HelpHook);
        }

        private IntPtr HelpHook(IntPtr hwnd,
            int msg,
            IntPtr wParam,
            IntPtr lParam,
            ref bool handled)
        {
            if (msg == WM_SYSCOMMAND &&
                ((int)wParam & 0xFFF0) == SC_CONTEXTHELP)
            {
                if (HelpButtonClick != null)
                    HelpButtonClick(this, EventArgs.Empty);
                handled = true;
            }

            return IntPtr.Zero;
        }

        protected string FileSelect(object sender, EnumFileFilter filter)
        {
            var text = ((TextBoxButton)sender).Text;
            if (_initialDirectory == null && !string.IsNullOrEmpty(text))
                _initialDirectory = Path.GetDirectoryName(text);
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = false,
                InitialDirectory = _initialDirectory,
                Title = @"Select Directory of Setting Folder"
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Cancel)
            {
                ((TextBoxButton)sender).Focus();
                return null;
            }

            ((TextBoxButton)sender).Focus();
            ((TextBoxButton)sender).Text = folderBrowserDialog.FileName;
            _initialDirectory = Path.GetDirectoryName(folderBrowserDialog.FileName);
            return folderBrowserDialog.FileName;
        }

        protected string PathSelect(object sender)
        {
            if (_initialDirectory == null)
                _initialDirectory = ((TextBoxButton)sender).Text;
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                InitialDirectory = _initialDirectory,
                Title = @"Select Directory of Setting Folder"
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
                if (filter == EnumFileFilter.TestProgram) return "(*.igxl,*.xls*)|*.xls*;*.igxl";

                if (filter == EnumFileFilter.Igxl) return "(*.igxl)|*.igxl";

                if (filter == EnumFileFilter.Dlex) return "(*.dlex)|*.dlex";

                if (filter == EnumFileFilter.Bat) return "(*.bat)|*.bat";

                if (filter == EnumFileFilter.PatternCsv) return "(*pattern*.csv)|*pattern*.csv";

                if (filter == EnumFileFilter.Excel) return "(*.xlsx,*.xlsm)|*.xlsx;*.xlsm";

                if (filter == EnumFileFilter.IdsDistribution)
                    return @"IDS Distribution" + @"(IDS_Distribution.txt)|IDS_Distribution.txt*";

                if (filter == EnumFileFilter.TestPlan) return @"TestPlan" + @"(*test*plan*.xlsx*)|*test*plan*.xlsx*";

                if (filter == EnumFileFilter.BinCut)
                    return @"*Bin*Cut*" + @"(*.txt,*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx";

                if (filter == EnumFileFilter.BinCutPost)
                    return @"*Post*Bin*Cut*" + @"(*.txt,*Bin*Cut*.xlsx)|*.txt;*Bin*Cut*.xlsx";

                if (filter == EnumFileFilter.TemplateFile) return "(*.tmp)|*.tmp";

                if (filter == EnumFileFilter.Txt) return "(*.txt)|*.txt";

                if (filter == EnumFileFilter.BasFile) return "(*.bas,*.txt)|*.bas;*.txt";

                if (filter == EnumFileFilter.PaFile) return "(*.csv,*.xls*)|*.csv;*.xlsx;*.xlsm";

                if (filter == EnumFileFilter.XmlFile) return "(*.xml)|*.xml";

                if (filter == EnumFileFilter.BdFile) return "(*.csv)|*.csv";

                return "";
            }
        }
    }
}