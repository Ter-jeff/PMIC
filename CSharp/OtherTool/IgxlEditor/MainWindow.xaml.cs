using IgxlData.IgxlManager;
using IgxlData.Zip;
using IgxlEditor.ViewModel;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using MyWpf.Controls;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Ribbon;
using System.Windows.Data;
using unvell.ReoGrid.IO;

namespace IgxlEditor
{
    public partial class MainWindow : RibbonWindow
    {
        private readonly MainViewModel _mainViewModel;
        private string _testProgram;

        public MainWindow()
        {
            //Style style = Application.Current.FindResource(typeof(HeaderedContentControl)) as Style;

            //var sb = new System.Text.StringBuilder();
            //var writer = System.Xml.XmlWriter.Create(sb, new System.Xml.XmlWriterSettings
            //{
            //    Indent = true,
            //    ConformanceLevel = System.Xml.ConformanceLevel.Fragment,
            //    OmitXmlDeclaration = true
            //});
            //var mgr = new System.Windows.Markup.XamlDesignerSerializationManager(writer);
            //mgr.XamlWriterMode = System.Windows.Markup.XamlWriterMode.Expression;
            //System.Windows.Markup.XamlWriter.Save(style, mgr);
            //string styleString = sb.ToString();

            InitializeComponent();

            _mainViewModel = new MainViewModel();

            MyListView.ItemsSource = _mainViewModel.SheetTypeRows;

            _mainViewModel.SheetTypeRows = new ObservableCollection<SheetTypeRow>();
            var igxlManagerMain = new IgxlManagerMain();
            _testProgram = @"C:\01.Jeffli\Sample.igxl";
            var sheetTypeRows = igxlManagerMain.GetIgxlSheetTypeRows(_testProgram);
            foreach (var sheetTypeRow in sheetTypeRows)
                _mainViewModel.SheetTypeRows.Add(sheetTypeRow);
            ICollectionView view = CollectionViewSource.GetDefaultView(_mainViewModel.SheetTypeRows);
            view.GroupDescriptions.Add(new PropertyGroupDescription("Type"));
            view.SortDescriptions.Add(new SortDescription("Type", ListSortDirection.Ascending));
            view.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
            MyListView.ItemsSource = view;

            MyEditText.Owner = MySheet.EditTextBox.Owner;
        }

        private void SearchTextBoxButton_OnTextChanged(object sender, EventArgs e)
        {
            var text = ((TextBoxButton)sender).Text;
            if (text == null)
                return;
            MyListView.ItemsSource = _mainViewModel.SheetTypeRows.Where(x => x.Name.ToUpper().Contains(text.ToUpper()));
        }

        private void MyListView_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_testProgram == null)
                return;
            var igxlManagerMain = new IgxlManagerMain();
            var row = (SheetTypeRow)(((ListBox)sender).SelectedItem);
            if (row == null)
                return;
            var name = row.Name.Replace(" ", "%20");
            var zipEntry = igxlManagerMain.GetIgxlSheets(_testProgram, name);
            MySheet.Load(zipEntry.OpenReader(), FileFormat.IGXL, row.Name);
        }

        private void FileOpen_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = new MyWindow.FileFilter().GetFilter(EnumFileFilter.Igxl),
                Multiselect = false
            };

            if ((bool)(openFileDialog.ShowDialog()))
            {
                _testProgram = openFileDialog.FileNames.First();
                _mainViewModel.SheetTypeRows = new ObservableCollection<SheetTypeRow>();
                var igxlManagerMain = new IgxlManagerMain();
                var sheetTypeRows = igxlManagerMain.GetIgxlSheetTypeRows(_testProgram);
                foreach (var sheetTypeRow in sheetTypeRows)
                    _mainViewModel.SheetTypeRows.Add(sheetTypeRow);
                ICollectionView view = CollectionViewSource.GetDefaultView(_mainViewModel.SheetTypeRows);
                view.GroupDescriptions.Add(new PropertyGroupDescription("Type"));
                view.SortDescriptions.Add(new SortDescription("Type", ListSortDirection.Ascending));
                view.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
                MyListView.ItemsSource = view;
            }
        }

        private void FileSave_Click(object sender, RoutedEventArgs e)
        {
            FileSave();
        }

        private void FileSave()
        {
            using (var zip = new ZipFile(_testProgram))
            {
                foreach (var sheet in MySheet.Worksheets)
                {
                    var content = sheet.ExportAsTxt();
                    var sheetName = sheet.Name.Replace(" ", "%20") + ".txt";
                    zip.AddUpdateEntry(sheetName, content);
                }

                zip.Save();
            }
        }

        private void Paste_OnClick(object sender, RoutedEventArgs e)
        {
            MySheet.ActiveWorksheet.Paste();
        }

        private void Cut_OnClick(object sender, RoutedEventArgs e)
        {
            MySheet.ActiveWorksheet.Cut();
        }

        private void Copy_OnClick(object sender, RoutedEventArgs e)
        {
            MySheet.ActiveWorksheet.Copy();
        }

        private void FileOpen_Import(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = new MyWindow.FileFilter().GetFilter(EnumFileFilter.Txt),
                Multiselect = true
            };

            if ((bool)(openFileDialog.ShowDialog()))
            {
                var igxlManagerMain = new IgxlManagerMain();
                foreach (var file in openFileDialog.FileNames)
                {
                    var name = Path.GetFileNameWithoutExtension(file);
                    var content = string.Join(Environment.NewLine, File.ReadAllLines(file));
                    igxlManagerMain.ImportTxt(_testProgram, name, content);
                    _mainViewModel.SheetTypeRows.Add(igxlManagerMain.GetIgxlSheetTypeRowsByTxt(file));
                }
            }
        }

        private void FileOpen_Export(object sender, RoutedEventArgs e)
        {
            var folderBrowserDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = @"Select Export Directory"
            };

            if (folderBrowserDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                FileSave();
                var igxlManagerMain = new IgxlManagerMain();
                igxlManagerMain.ExportTxt(_testProgram, folderBrowserDialog.FileName);
            }
        }
    }
}