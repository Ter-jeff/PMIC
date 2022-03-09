using System.Reflection;

namespace PMICAutogenAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
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
            this.button_Help = this.Factory.CreateRibbonButton();
            this.tab_Autogen.SuspendLayout();
            this.group_Pmic.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_Autogen
            // 
            this.tab_Autogen.Groups.Add(this.group_Pmic);
            this.tab_Autogen.Label = "Autogen";
            this.tab_Autogen.Name = "tab_Autogen";
            // 
            // group_Pmic
            // 
            this.group_Pmic.Items.Add(this.button_Validate);
            this.group_Pmic.Items.Add(this.button_Autogen);
            this.group_Pmic.Items.Add(this.button_Help);
            this.group_Pmic.Label = "V17.0.0.0";
            this.group_Pmic.Name = "group_Pmic";
            // 
            // button_Validate
            // 
            this.button_Validate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Validate.Image = global::PMICAutogenAddIn.Properties.Resources.alarm;
            this.button_Validate.Label = "Validate";
            this.button_Validate.Name = "button_Validate";
            this.button_Validate.ShowImage = true;
            this.button_Validate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Validate_Click);
            // 
            // button_Autogen
            // 
            this.button_Autogen.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Autogen.Image = global::PMICAutogenAddIn.Properties.Resources.play_button;
            this.button_Autogen.Label = "Autogen";
            this.button_Autogen.Name = "button_Autogen";
            this.button_Autogen.ShowImage = true;
            this.button_Autogen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Autogen_Click);
            // 
            // button_Help
            // 
            this.button_Help.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Help.Image = global::PMICAutogenAddIn.Properties.Resources.info;
            this.button_Help.Label = "Help";
            this.button_Help.Name = "button_Help";
            this.button_Help.ShowImage = true;
            this.button_Help.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Help_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
