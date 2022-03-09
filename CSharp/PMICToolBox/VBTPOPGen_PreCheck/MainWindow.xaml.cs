using VBTPOPGen_PreCheck.Tools;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using VBTPOPGenPreCheckBusiness.Business;
using VBTPOPGenPreCheckBusiness.DataStore;

namespace VBTPOPGen_PreCheck
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MyDialog dialog = null;
        private GuiInfo guiInfo = null;
        private Action<string> _downLoadEvent;
        public MainWindow()
        {
            InitializeComponent();
            dialog = MyDialog.GetInstance();
            guiInfo = GuiInfo.GetInstance();
            this.grid.DataContext = guiInfo;
        }

        private void btnTestplan_Click(object sender, RoutedEventArgs e)
        {
            string path = dialog.PathDialog(true, false);
            if (path != null)
            {
                guiInfo.testplanFolder = path;
                tbTestplan.Text = path;

                guiInfo.testplanList = Directory.GetFiles(path)
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                    .Where(s => !s.Contains("~$")).ToList();
            }
        }

        private void btnOTPRegMap_Click(object sender, RoutedEventArgs e)
        {
            var file = dialog.FileDialog(FileFilter.BasFile, false);
            if (file != null)
            {
                guiInfo.otpRegMap = file;
                tbOTPRegMap.Text = file;
            }
        }

        private void btnAHBRegMap_Click(object sender, RoutedEventArgs e)
        {
            var file = dialog.FileDialog(FileFilter.BasFile, false);
            if (file != null)
            {
                guiInfo.AhbRegMap = file;
                tbAhbRegMap.Text = file;
            }
        }

        private void btnPinMap_Click(object sender, RoutedEventArgs e)
        {
            var file = dialog.FileDialog(FileFilter.BasFile, false);
            if (file != null)
            {
                guiInfo.PinMap = file;
                tbPinMap.Text = file;
            }
        }

        private void btnOutput_Click(object sender, RoutedEventArgs e)
        {
            string path = dialog.PathDialog(true, false);
            if (path != null)
            {
                guiInfo.output = path;
                tbOutput.Text = path;
            }
        }

        private bool CheckIOStatus()
        {
            if (!guiInfo.testplanList.Any())
            {
                MessageBox.Show("Can not find any TestPlan file!");
                return false;
            }
            if (guiInfo.otpRegMap.Equals(string.Empty))
                return false;
            if (!Directory.Exists(guiInfo.output))
                Directory.CreateDirectory(guiInfo.output);
            return true;
        }

        private async void btnGo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                guiInfo.ProgressValue = "0";
                guiInfo.UISateInfo = "";
                if (!CheckIOStatus())
                    return;

                btnGo.IsEnabled = false;
                await Task.Run(() =>
                {
                    CheckProcess cp = new CheckProcess(guiInfo);
                    cp.WorkFlow();
                });
                btnGo.IsEnabled = true;                
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                btnGo.IsEnabled = true;
                guiInfo.ProgressValue = "0";
                guiInfo.UISateInfo = "";
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            _downLoadEvent(resourceName);
        }

        public void SetDownLoadEvent(Action<string> inputEvent)
        {
            _downLoadEvent = inputEvent;
        }
    }
}
