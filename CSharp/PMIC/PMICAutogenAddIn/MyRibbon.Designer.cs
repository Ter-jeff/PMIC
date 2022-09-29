using System.Reflection;

namespace PMICAutogenAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            Assembly assembly = Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            this.group_Pmic.Label = version;
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_Autogen = this.Factory.CreateRibbonTab();
            this.group_Pmic = this.Factory.CreateRibbonGroup();
            this.button_Validate = this.Factory.CreateRibbonButton();
            this.button_Autogen = this.Factory.CreateRibbonButton();
            this.button_back = this.Factory.CreateRibbonButton();
            this.button_Help = this.Factory.CreateRibbonButton();
            this.button_History = this.Factory.CreateRibbonButton();
            this.tab_Autogen.SuspendLayout();
            this.group_Pmic.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Autogen
            // 
            this.tab_Autogen.Groups.Add(this.group_Pmic);
            this.tab_Autogen.Label = global::PMICAutogenAddIn.Properties.Resources.TabAutogen;
            this.tab_Autogen.Name = "tab_Autogen";
            // 
            // group_Pmic
            // 
            this.group_Pmic.Items.Add(this.button_Validate);
            this.group_Pmic.Items.Add(this.button_Autogen);
            this.group_Pmic.Items.Add(this.button_back);
            this.group_Pmic.Items.Add(this.button_Help);
            this.group_Pmic.Items.Add(this.button_History);
            this.group_Pmic.Name = "group_Pmic";
            // 
            // button_Validate
            // 
            this.button_Validate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Validate.Description = "Validate PMIC test plan11";
            this.button_Validate.Image = global::PMICAutogenAddIn.Properties.Resources.bell;
            this.button_Validate.Label = "Validate";
            this.button_Validate.Name = "button_Validate";
            this.button_Validate.ScreenTip = "Validate";
            this.button_Validate.ShowImage = true;
            this.button_Validate.SuperTip = "Validate PMIC test plan";
            this.button_Validate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Validate_Click);
            // 
            // button_Autogen
            // 
            this.button_Autogen.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Autogen.Image = global::PMICAutogenAddIn.Properties.Resources.play_button;
            this.button_Autogen.Label = "Autogen";
            this.button_Autogen.Name = "button_Autogen";
            this.button_Autogen.ScreenTip = "Autogen";
            this.button_Autogen.ShowImage = true;
            this.button_Autogen.SuperTip = "Auto generate test program";
            this.button_Autogen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Autogen_Click);
            // 
            // button_back
            // 
            this.button_back.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_back.Image = global::PMICAutogenAddIn.Properties.Resources.back;
            this.button_back.Label = "Back";
            this.button_back.Name = "button_back";
            this.button_back.ScreenTip = "Back (Ctrl+Shift+Q)";
            this.button_back.ShowImage = true;
            this.button_back.SuperTip = "Go back last sheet";
            this.button_back.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Back_Click);
            // 
            // button_Help
            // 
            this.button_Help.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Help.Image = global::PMICAutogenAddIn.Properties.Resources.info;
            this.button_Help.Label = "Help";
            this.button_Help.Name = "button_Help";
            this.button_Help.ScreenTip = "Help";
            this.button_Help.ShowImage = true;
            this.button_Help.SuperTip = "PMIC autogen user manual";
            this.button_Help.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Help_Click);
            // 
            // button_History
            // 
            this.button_History.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_History.Image = global::PMICAutogenAddIn.Properties.Resources.calendar;
            this.button_History.Label = "History";
            this.button_History.Name = "button_History";
            this.button_History.ScreenTip = "History";
            this.button_History.ShowImage = true;
            this.button_History.SuperTip = "Version change history";
            this.button_History.Visible = false;
            this.button_History.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_History_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_Autogen);
            this.tab_Autogen.ResumeLayout(false);
            this.tab_Autogen.PerformLayout();
            this.group_Pmic.ResumeLayout(false);
            this.group_Pmic.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        public Microsoft.Office.Tools.Ribbon.RibbonTab tab_Autogen;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_Pmic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Validate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Autogen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Help;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_History;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_back;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
