using PmicAutomation.MyControls;
using PmicAutomation.Utility.AHBEnum;
using PmicAutomation.Utility.nWire;
using PmicAutomation.Utility.PA;
using PmicAutomation.Utility.PatSetsAll;
using PmicAutomation.Utility.Relay;
using PmicAutomation.Utility.VbtGenerator;
using PmicAutomation.Utility.VbtGenToolTemplate;
using PmicAutomation.Utility.ErrorHandler;
using PmicAutomation.Utility.TCMID;
using PmicAutomation.Utility.TCMIDComparator;
using PinNameRemoval;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using PmicAutomation.Utility.OTPRegisterMap;
using System.Collections.Generic;
using ProfileTool_PMIC;
using System.Reflection;
using BinOutCheck;
using PmicAutomation.Utility.DatalogComparator;
using AlarmChekc;

namespace Automation
{
    public partial class PmicMainForm : Form
    {
        public PmicMainForm()
        {
            InitializeComponent();

            HelpButtonClicked += PmicMainForm_HelpButtonClicked;
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            Text = "PMIC Application ToolBox : Version - " + version;

            //Hide nWireDefinition tool (or others) when toolbox started.
            checkBox_ShowAll.CheckState = CheckState.Unchecked;
            nWireDefinition.Visible = false;
            ribbonButton_nWireDefinition.Visible = false;
        }

        private void CheckBoxShowAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_ShowAll.CheckState == CheckState.Unchecked)
            {
                nWireDefinition.Visible = false;
                ribbonButton_nWireDefinition.Visible = false;
            }
            else
            {
                nWireDefinition.Visible = true;
                ribbonButton_nWireDefinition.Visible = true;
            }
        }

        private void RibbonButton_PaParser_Click(object sender, EventArgs e)
        {
            new PaParser().Show();
        }

        private void RibbonButton_VbtGenerator_Click(object sender, EventArgs e)
        {
            new VbtGeneratorFrom().Show();
        }

        private void RibbonButton_Ahb_Enum_Click(object sender, EventArgs e)
        {
            new AhbEnum().Show();
        }

        private void RibbonButton_TemplateGenerator_Click(object sender, EventArgs e)
        {
            new VbtGenToolGenerator().Show();
        }

        private void RibbonButton_Relay_Click(object sender, EventArgs e)
        {
            new Relay().Show();
        }

        private void RibbonButton_PinNameRemoval_Click(object sender, EventArgs e)
        {
            new PinNameRemovalForm().Show();
        }

        private void RibbonButton_nWireDefinition_Click(object sender, EventArgs e)
        {
            new MainForm().Show();
        }

        private void RibbonButton_PatSetsAll_Click(object sender, EventArgs e)
        {
            new PatSetsAllForm().Show();
        }

        private void RibbonButton_OTP_Click(object sender, EventArgs e)
        {
            new OtpRegisterMapFrom().Show();
        }


        private void PmicMainForm_HelpButtonClicked(Object sender, EventArgs e)
        {
            List<string> nameList = new List<string>();
            if (this.ribbon.ActiveTab.Text.Equals("Main", StringComparison.OrdinalIgnoreCase))
            {
                nameList.Add(".Utility.PA.UserManual.");
                nameList.Add(".Utility.VbtGenerator.UserManual.");
                nameList.Add(".Utility.PatSetsAll.UserManual.");
                nameList.Add(".Utility.OTPRegisterMap.UserManual.");
                nameList.Add(".Utility.ErrorHandler.UserManual.");
                nameList.Add(".Utility.TCMID.UserManual.");
                nameList.Add(".UserManual.PMIC_BinCheck_Tool");
                nameList.Add(".UserManual.Datalog");
            }
            else if (this.ribbon.ActiveTab.Text.Equals("Sub", StringComparison.OrdinalIgnoreCase))
            {
                nameList.Add(".Utility.AHBEnum.UserManual.");
                //nameList.Add(".Utility.VbtGenToolTemplate.UserManual.");
                nameList.Add(".Utility.Relay.UserManual.");
                nameList.Add(".UserManual.nWire");
                nameList.Add(".UserManual.PinNameRemoval");
            }
            new MyDownloadForm().DownloadContains(nameList).Show();
        }

        private void ribbonButton_ExceptionHandling_Click(object sender, EventArgs e)
        {
            new ErrorHandlerForm().Show();
        }

        private void ribbonButton_DatalogComparator_Click(object sender, EventArgs e)
        {
            new DatalogComparatorForm().Show();
        }

        private void RibbonButton_GenTCMID_Click(object sender, EventArgs e)
        {
            new TCMIDForm().Show();
        }

        private void RibbonButton_TCMIDComparator_Click(object sender, EventArgs e)
        {
            new TCMIDComparatorForm().Show();
        }

        private void ribbonButton_ProfileTool_Click(object sender, EventArgs e)
        {
            new ProfileToolForm().Show();
        }

        private void ribbonButton_clbist_Click(object sender,EventArgs e)
        {
            var temp = new CLBistDataConverter.frmMain();
            temp.SetDownLoadEvent((res)=> {Assembly ass=Assembly.GetAssembly(typeof(CLBistDataConverter.DataStructures.CLBistSite)); new MyDownloadForm().DownloadFromAssembly(res + ".Template.",ass).Show(); });
            temp.Show();
        }

        private void ribbonButton_vbtpopprecheck_Click(object sender, EventArgs e)
        {
            //var temp = new frmMain();
            //temp.SetDownLoadEvent((res) => { Assembly ass = Assembly.GetAssembly(typeof(CLBistDataConverter.DataStructures.CLBistSite)); new MyDownloadForm().DownloadFromAssembly(res + ".Template.", ass).Show(); });
            //temp.Show();
            VBTPOPGen_PreCheck.MainWindow window = new VBTPOPGen_PreCheck.MainWindow();
            window.SetDownLoadEvent((res) => { Assembly ass = Assembly.GetAssembly(typeof(VBTPOPGen_PreCheck.MainWindow)); new MyDownloadForm().DownloadFromAssembly(res + ".Template.", ass).Show(); });
            window.ShowDialog();
        }

        private void ribbonButton_binoutcheck_Click(object sender, EventArgs e)
        {
            BinOutCheckForm form = new BinOutCheckForm();
            form.SetDownLoadEvent((res) => { Assembly ass = Assembly.GetAssembly(typeof(BinOutCheckForm)); new MyDownloadForm().DownloadFromAssembly(res + ".Template.", ass).Show(); });
            form.Show();
        }
        private void ribbonButton_otpFilesComparing_Click(object sender, EventArgs e)
        {
            OTPFileComparison.frmMain form = new OTPFileComparison.frmMain();
            form.SetDownLoadEvent((res) => { Assembly ass = Assembly.GetAssembly(typeof(OTPFileComparison.frmMain)); new MyDownloadForm().DownloadFromAssembly(res + ".Template.", ass).Show(); });
            form.Show();
        }

        private void AlarmCheck_Click(object sender, EventArgs e)
        {
            new AlarmCheck().Show();
        }
    }
}