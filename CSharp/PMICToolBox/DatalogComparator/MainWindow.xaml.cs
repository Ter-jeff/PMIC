using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using Library.Output;
using Library.Common;
using System.Text.RegularExpressions;
using Library.DataStruct;
using Library.Input;
using Library;

namespace DatalogComparator
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private System.ComponentModel.BackgroundWorker _bgWorkerForHardWorkCheck = null;
        private CommonData data = CommonData.GetInstance();
        
        public MainWindow()
        {
            InitializeComponent();
            this.grid.DataContext = data;            
        }

        private void SelctOutputFolderBut_Click(object sender, RoutedEventArgs e)
        {
            const string descriOutput = "Select Output Folder";
            FolderBrowserDialog folderBroDialog = new FolderBrowserDialog();
            System.Windows.Controls.Button srcButton = (System.Windows.Controls.Button)sender;
            folderBroDialog.Description = descriOutput;
            folderBroDialog.RootFolder = Environment.SpecialFolder.Desktop;
            folderBroDialog.ShowDialog();
            if (!String.IsNullOrEmpty(folderBroDialog.SelectedPath))
            {
                this.OutputPath.Text = folderBroDialog.SelectedPath;
                this.OutputPathToolTip.Text = folderBroDialog.SelectedPath;
            }
        }

        private void SelctDatalogFileBut_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = (System.Windows.Controls.Button)sender;
            String buttonName = button.Name;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.RestoreDirectory = true;
            dialog.Title = "Select TXT Datalog File";
            dialog.Filter = "File(*.txt)|*.txt";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                switch (buttonName)
                {
                    case "SelectTxtlogBut":
                        this.txtdatalogPath.Text = dialog.FileName;
                        this.txtdatalogPathToolTip.Text = dialog.FileName;
                        break;
                    case "SelectReftxtlogBut":
                        this.reftxtdatalogPath.Text = dialog.FileName;
                        this.reftxtdatalogPathToolTip.Text = dialog.FileName;
                        break;
                }                
            }
        }

        private void RunTestProgram(object sender, EventArgs e)
        {

            CommonData.GetInstance().worker = _bgWorkerForHardWorkCheck;
            MainLogic.Instance().MainFlow();
        }

        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            CommonData.GetInstance().ProgressValue = (e.ProgressPercentage.ToString());
        }
        private void WorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            CommonData.GetInstance().UIEnabled = true;
            this.Cursor = null;
        }
        private void RunBut_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = System.Windows.Input.Cursors.Wait;
            this.grid.IsEnabled = false;
            
            CommonData.GetInstance().OutputPath = this.OutputPath.Text;
            CommonData.GetInstance().BaseTxtDatalogPath = this.txtdatalogPath.Text;
            CommonData.GetInstance().CompareTxtDatalogPath = this.reftxtdatalogPathToolTip.Text;

            if (!Directory.Exists(this.OutputPath.Text))
            {
                System.Windows.Forms.MessageBox.Show("Output Folder is not exist!", "Error", MessageBoxButtons.OK);
                this.grid.IsEnabled = true;
                this.Cursor = System.Windows.Input.Cursors.Arrow;
                return;
            }

            _bgWorkerForHardWorkCheck = new System.ComponentModel.BackgroundWorker();
            _bgWorkerForHardWorkCheck.WorkerReportsProgress = true;
            _bgWorkerForHardWorkCheck.DoWork += new System.ComponentModel.DoWorkEventHandler(RunTestProgram);
            _bgWorkerForHardWorkCheck.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(ProgressChanged);
            _bgWorkerForHardWorkCheck.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(WorkerCompleted);
            _bgWorkerForHardWorkCheck.RunWorkerAsync();
             this.grid.IsEnabled = true;
             this.Cursor = System.Windows.Input.Cursors.Arrow;
        }

        private void RunTest_Click(object sender, RoutedEventArgs e)
        {

            DatalogReader datalogReader = new DatalogReader();
            CommonData.GetInstance().CompareTxtDatalogPath = this.reftxtdatalogPath.Text;
            CommonData.GetInstance().BaseTxtDatalogPath = this.txtdatalogPath.Text;
            CommonData.GetInstance().OutputPath = this.OutputPath.Text;
            //List<TestInstanceItem> datalogInstancelst = datalogReader.Read(CommonData.GetInstance().RefTxtDatalogPath);

            //compare datalog
            List<InstanceCompareResult> compareResult = MainLogic.Instance().CompareDatalog(CommonData.GetInstance().BaseTxtDatalogPath,
                CommonData.GetInstance().CompareTxtDatalogPath);
            //generate output report
            string currentTimeStr = DateTime.Now.ToString("yyyymmddhhmmss");
            string reportFIlePath = CommonData.GetInstance().OutputPath + "\\DiffReport_" + currentTimeStr + ".xlsx";
            new ReportWriter().GenerateDiffReport(compareResult, MainLogic.Instance().logDiffResultlst, reportFIlePath);
        }
}
}
