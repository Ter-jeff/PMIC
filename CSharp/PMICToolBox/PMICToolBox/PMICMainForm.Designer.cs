using System.Diagnostics;

namespace Automation
{
    partial class PmicMainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PmicMainForm));
            this.ribbon = new System.Windows.Forms.Ribbon();
            this.ribbonTab_Utility = new System.Windows.Forms.RibbonTab();
            this.ribbonPane_PaParser = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_PaParser = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_VbtGenerator = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_VbtGenerator = new System.Windows.Forms.RibbonButton();
            this.PatSetsAll = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_PatSets_All = new System.Windows.Forms.RibbonButton();
            this.OTPRegisterMap = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_OTPRegisterMap = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel1 = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_ExceptionHandling = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel2 = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_DataComparator = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_vbtpopprecheck = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_vbtpopprecheck = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_binoutcheck = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_binoutcheck = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_otpFilesComparing = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_otpFilesComparing = new System.Windows.Forms.RibbonButton();
            this.ribbonTab_Tool = new System.Windows.Forms.RibbonTab();
            this.ribbonPanel_Ahb_Enum = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_Ahb_Enum = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_TemplateGenerator = new System.Windows.Forms.RibbonPanel();
            this.ribbonPanel_Relay = new System.Windows.Forms.RibbonPanel();
            this.RibbonButton_Relay = new System.Windows.Forms.RibbonButton();
            this.PinNameRemoval = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_PinNameRemoval = new System.Windows.Forms.RibbonButton();
            this.nWireDefinition = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_nWireDefinition = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_clbist = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_clbist = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_GenTCMID = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_GenTCMID = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_TCMIDComparator = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_TCMIDComparator = new System.Windows.Forms.RibbonButton();
            this.ribbonPanel_ProfileTool = new System.Windows.Forms.RibbonPanel();
            this.ribbonButton_ProfileTool = new System.Windows.Forms.RibbonButton();
            this.ribbonButton1 = new System.Windows.Forms.RibbonButton();
            this.checkBox_ShowAll = new System.Windows.Forms.CheckBox();
            this.ribbonPanel_AlarmCheck = new System.Windows.Forms.RibbonPanel();
            this.AlarmCheck = new System.Windows.Forms.RibbonButton();
            this.SuspendLayout();
            // 
            // ribbon
            // 
            this.ribbon.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ribbon.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.ribbon.Location = new System.Drawing.Point(0, 0);
            this.ribbon.Margin = new System.Windows.Forms.Padding(0);
            this.ribbon.Minimized = false;
            this.ribbon.Name = "ribbon";
            // 
            // 
            // 
            this.ribbon.OrbDropDown.BorderRoundness = 8;
            this.ribbon.OrbDropDown.Location = new System.Drawing.Point(0, 0);
            this.ribbon.OrbDropDown.Name = "";
            this.ribbon.OrbDropDown.TabIndex = 0;
            this.ribbon.OrbStyle = System.Windows.Forms.RibbonOrbStyle.Office_2013;
            this.ribbon.OrbVisible = false;
            // 
            // 
            // 
            this.ribbon.QuickAccessToolbar.DropDownButtonVisible = false;
            this.ribbon.QuickAccessToolbar.FlashEnabled = true;
            this.ribbon.QuickAccessToolbar.ShowFlashImage = true;
            this.ribbon.RibbonTabFont = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.ribbon.Size = new System.Drawing.Size(800, 180);
            this.ribbon.TabIndex = 1;
            this.ribbon.Tabs.Add(this.ribbonTab_Utility);
            this.ribbon.Tabs.Add(this.ribbonTab_Tool);
            this.ribbon.TabsMargin = new System.Windows.Forms.Padding(5, 26, 20, 0);
            this.ribbon.TabSpacing = 4;
            this.ribbon.Text = "ribbon1";
            // 
            // ribbonTab_Utility
            // 
            this.ribbonTab_Utility.Name = "ribbonTab_Utility";
            this.ribbonTab_Utility.Panels.Add(this.ribbonPane_PaParser);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel_VbtGenerator);
            this.ribbonTab_Utility.Panels.Add(this.PatSetsAll);
            this.ribbonTab_Utility.Panels.Add(this.OTPRegisterMap);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel1);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel2);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel_vbtpopprecheck);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel_binoutcheck);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel_otpFilesComparing);
            this.ribbonTab_Utility.Panels.Add(this.ribbonPanel_AlarmCheck);
            this.ribbonTab_Utility.Text = "Main";
            // 
            // ribbonPane_PaParser
            // 
            this.ribbonPane_PaParser.Items.Add(this.ribbonButton_PaParser);
            this.ribbonPane_PaParser.Name = "ribbonPane_PaParser";
            this.ribbonPane_PaParser.Text = "";
            // 
            // ribbonButton_PaParser
            // 
            this.ribbonButton_PaParser.Image = global::PmicAutomation.Properties.Resources._002_pantone;
            this.ribbonButton_PaParser.LargeImage = global::PmicAutomation.Properties.Resources._002_pantone;
            this.ribbonButton_PaParser.Name = "ribbonButton_PaParser";
            this.ribbonButton_PaParser.SmallImage = global::PmicAutomation.Properties.Resources._002A_pantone;
            this.ribbonButton_PaParser.Text = "PA parser";
            this.ribbonButton_PaParser.ToolTip = "Parse PA to generate pin map and channel map";
            this.ribbonButton_PaParser.Click += new System.EventHandler(this.RibbonButton_PaParser_Click);
            // 
            // ribbonPanel_VbtGenerator
            // 
            this.ribbonPanel_VbtGenerator.Items.Add(this.ribbonButton_VbtGenerator);
            this.ribbonPanel_VbtGenerator.Name = "ribbonPanel_VbtGenerator";
            this.ribbonPanel_VbtGenerator.Text = "VBT";
            // 
            // ribbonButton_VbtGenerator
            // 
            this.ribbonButton_VbtGenerator.Image = global::PmicAutomation.Properties.Resources._005_rgb;
            this.ribbonButton_VbtGenerator.LargeImage = global::PmicAutomation.Properties.Resources._005_rgb;
            this.ribbonButton_VbtGenerator.Name = "ribbonButton_VbtGenerator";
            this.ribbonButton_VbtGenerator.SmallImage = global::PmicAutomation.Properties.Resources._005A_rgb;
            this.ribbonButton_VbtGenerator.Text = "Gen DC Support";
            this.ribbonButton_VbtGenerator.ToolTip = "Generate VBT by template";
            this.ribbonButton_VbtGenerator.Click += new System.EventHandler(this.RibbonButton_VbtGenerator_Click);
            // 
            // PatSetsAll
            // 
            this.PatSetsAll.Items.Add(this.ribbonButton_PatSets_All);
            this.PatSetsAll.Name = "PatSetsAll";
            this.PatSetsAll.Text = "";
            // 
            // ribbonButton_PatSets_All
            // 
            this.ribbonButton_PatSets_All.Image = global::PmicAutomation.Properties.Resources._017_compass;
            this.ribbonButton_PatSets_All.LargeImage = global::PmicAutomation.Properties.Resources._017_compass;
            this.ribbonButton_PatSets_All.Name = "ribbonButton_PatSets_All";
            this.ribbonButton_PatSets_All.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_PatSets_All.SmallImage")));
            this.ribbonButton_PatSets_All.Text = "PatSets All";
            this.ribbonButton_PatSets_All.Click += new System.EventHandler(this.RibbonButton_PatSetsAll_Click);
            // 
            // OTPRegisterMap
            // 
            this.OTPRegisterMap.Items.Add(this.ribbonButton_OTPRegisterMap);
            this.OTPRegisterMap.Name = "OTPRegisterMap";
            this.OTPRegisterMap.Text = "";
            // 
            // ribbonButton_OTPRegisterMap
            // 
            this.ribbonButton_OTPRegisterMap.Image = global::PmicAutomation.Properties.Resources._018_photo_film;
            this.ribbonButton_OTPRegisterMap.LargeImage = global::PmicAutomation.Properties.Resources._018_photo_film;
            this.ribbonButton_OTPRegisterMap.Name = "ribbonButton_OTPRegisterMap";
            this.ribbonButton_OTPRegisterMap.SmallImage = global::PmicAutomation.Properties.Resources._018A_photo_film;
            this.ribbonButton_OTPRegisterMap.Text = "OTP Register Map";
            this.ribbonButton_OTPRegisterMap.Click += new System.EventHandler(this.RibbonButton_OTP_Click);
            // 
            // ribbonPanel1
            // 
            this.ribbonPanel1.Items.Add(this.ribbonButton_ExceptionHandling);
            this.ribbonPanel1.Name = "ribbonPanel1";
            this.ribbonPanel1.Text = "";
            // 
            // ribbonButton_ExceptionHandling
            // 
            this.ribbonButton_ExceptionHandling.Image = ((System.Drawing.Image)(resources.GetObject("ribbonButton_ExceptionHandling.Image")));
            this.ribbonButton_ExceptionHandling.LargeImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_ExceptionHandling.LargeImage")));
            this.ribbonButton_ExceptionHandling.Name = "ribbonButton_ExceptionHandling";
            this.ribbonButton_ExceptionHandling.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_ExceptionHandling.SmallImage")));
            this.ribbonButton_ExceptionHandling.Text = "Check Exception Handling";
            this.ribbonButton_ExceptionHandling.Click += new System.EventHandler(this.ribbonButton_ExceptionHandling_Click);
            // 
            // ribbonPanel2
            // 
            this.ribbonPanel2.Items.Add(this.ribbonButton_DataComparator);
            this.ribbonPanel2.Name = "ribbonPanel2";
            this.ribbonPanel2.Text = "";
            // 
            // ribbonButton_DataComparator
            // 
            this.ribbonButton_DataComparator.Image = ((System.Drawing.Image)(resources.GetObject("ribbonButton_DataComparator.Image")));
            this.ribbonButton_DataComparator.LargeImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_DataComparator.LargeImage")));
            this.ribbonButton_DataComparator.Name = "ribbonButton_DataComparator";
            this.ribbonButton_DataComparator.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_DataComparator.SmallImage")));
            this.ribbonButton_DataComparator.Text = "Datalog Comparator";
            this.ribbonButton_DataComparator.Click += new System.EventHandler(this.ribbonButton_DatalogComparator_Click);
            // 
            // ribbonPanel_vbtpopprecheck
            // 
            this.ribbonPanel_vbtpopprecheck.Items.Add(this.ribbonButton_vbtpopprecheck);
            this.ribbonPanel_vbtpopprecheck.Name = "ribbonPanel_vbtpopprecheck";
            this.ribbonPanel_vbtpopprecheck.Text = "";
            // 
            // ribbonButton_vbtpopprecheck
            // 
            this.ribbonButton_vbtpopprecheck.Image = global::PmicAutomation.Properties.Resources._025_sketchbook;
            this.ribbonButton_vbtpopprecheck.LargeImage = global::PmicAutomation.Properties.Resources._025_sketchbook;
            this.ribbonButton_vbtpopprecheck.Name = "ribbonButton_vbtpopprecheck";
            this.ribbonButton_vbtpopprecheck.SmallImage = global::PmicAutomation.Properties.Resources._025_sketchbook;
            this.ribbonButton_vbtpopprecheck.Text = "VBTPOP Gen PreCheck";
            this.ribbonButton_vbtpopprecheck.Click += new System.EventHandler(this.ribbonButton_vbtpopprecheck_Click);
            // 
            // ribbonPanel_binoutcheck
            // 
            this.ribbonPanel_binoutcheck.Items.Add(this.ribbonButton_binoutcheck);
            this.ribbonPanel_binoutcheck.Name = "ribbonPanel_binoutcheck";
            this.ribbonPanel_binoutcheck.Text = "";
            // 
            // ribbonButton_binoutcheck
            // 
            this.ribbonButton_binoutcheck.Image = global::PmicAutomation.Properties.Resources._010_target1;
            this.ribbonButton_binoutcheck.LargeImage = global::PmicAutomation.Properties.Resources._010_target1;
            this.ribbonButton_binoutcheck.Name = "ribbonButton_binoutcheck";
            this.ribbonButton_binoutcheck.SmallImage = global::PmicAutomation.Properties.Resources._010_target1;
            this.ribbonButton_binoutcheck.Text = "BinOut Check";
            this.ribbonButton_binoutcheck.Click += new System.EventHandler(this.ribbonButton_binoutcheck_Click);
            // 
            // ribbonPanel_otpFilesComparing
            // 
            this.ribbonPanel_otpFilesComparing.Items.Add(this.ribbonButton_otpFilesComparing);
            this.ribbonPanel_otpFilesComparing.Name = "ribbonPanel_otpFilesComparing";
            this.ribbonPanel_otpFilesComparing.Text = "";
            // 
            // ribbonButton_otpFilesComparing
            // 
            this.ribbonButton_otpFilesComparing.Image = global::PmicAutomation.Properties.Resources._013_creative;
            this.ribbonButton_otpFilesComparing.LargeImage = global::PmicAutomation.Properties.Resources._013_creative;
            this.ribbonButton_otpFilesComparing.Name = "ribbonButton_otpFilesComparing";
            this.ribbonButton_otpFilesComparing.SmallImage = global::PmicAutomation.Properties.Resources._013_creative;
            this.ribbonButton_otpFilesComparing.Text = "OTP Files Comparator";
            this.ribbonButton_otpFilesComparing.Click += new System.EventHandler(this.ribbonButton_otpFilesComparing_Click);
            // 
            // ribbonTab_Tool
            // 
            this.ribbonTab_Tool.Name = "ribbonTab_Tool";
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_Ahb_Enum);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_TemplateGenerator);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_Relay);
            this.ribbonTab_Tool.Panels.Add(this.PinNameRemoval);
            this.ribbonTab_Tool.Panels.Add(this.nWireDefinition);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_clbist);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_GenTCMID);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_TCMIDComparator);
            this.ribbonTab_Tool.Panels.Add(this.ribbonPanel_ProfileTool);
            this.ribbonTab_Tool.Text = "Sub";
            // 
            // ribbonPanel_Ahb_Enum
            // 
            this.ribbonPanel_Ahb_Enum.Items.Add(this.ribbonButton_Ahb_Enum);
            this.ribbonPanel_Ahb_Enum.Name = "ribbonPanel_Ahb_Enum";
            this.ribbonPanel_Ahb_Enum.Text = "VBT";
            // 
            // ribbonButton_Ahb_Enum
            // 
            this.ribbonButton_Ahb_Enum.Image = global::PmicAutomation.Properties.Resources._008_pencil;
            this.ribbonButton_Ahb_Enum.LargeImage = global::PmicAutomation.Properties.Resources._008_pencil;
            this.ribbonButton_Ahb_Enum.Name = "ribbonButton_Ahb_Enum";
            this.ribbonButton_Ahb_Enum.SmallImage = global::PmicAutomation.Properties.Resources._008A_pencil;
            this.ribbonButton_Ahb_Enum.Text = "Gen AHB";
            this.ribbonButton_Ahb_Enum.ToolTip = "Generate AHB Enum by AHB AHB_register_map";
            this.ribbonButton_Ahb_Enum.Click += new System.EventHandler(this.RibbonButton_Ahb_Enum_Click);
            // 
            // ribbonPanel_TemplateGenerator
            // 
            this.ribbonPanel_TemplateGenerator.Name = "ribbonPanel_TemplateGenerator";
            this.ribbonPanel_TemplateGenerator.Text = null;
            this.ribbonPanel_TemplateGenerator.Visible = false;
            // 
            // ribbonPanel_Relay
            // 
            this.ribbonPanel_Relay.Items.Add(this.RibbonButton_Relay);
            this.ribbonPanel_Relay.Name = "ribbonPanel_Relay";
            this.ribbonPanel_Relay.Text = "";
            // 
            // RibbonButton_Relay
            // 
            this.RibbonButton_Relay.Image = global::PmicAutomation.Properties.Resources._011_measurent;
            this.RibbonButton_Relay.LargeImage = global::PmicAutomation.Properties.Resources._011_measurent;
            this.RibbonButton_Relay.Name = "RibbonButton_Relay";
            this.RibbonButton_Relay.SmallImage = global::PmicAutomation.Properties.Resources._011A_measurent;
            this.RibbonButton_Relay.Text = "Relay";
            this.RibbonButton_Relay.ToolTip = "Generate relay setting by Component_Pin_Report";
            this.RibbonButton_Relay.Click += new System.EventHandler(this.RibbonButton_Relay_Click);
            // 
            // PinNameRemoval
            // 
            this.PinNameRemoval.Items.Add(this.ribbonButton_PinNameRemoval);
            this.PinNameRemoval.Name = "PinNameRemoval";
            this.PinNameRemoval.Text = "";
            // 
            // ribbonButton_PinNameRemoval
            // 
            this.ribbonButton_PinNameRemoval.Image = global::PmicAutomation.Properties.Resources._016_grid;
            this.ribbonButton_PinNameRemoval.LargeImage = global::PmicAutomation.Properties.Resources._016_grid;
            this.ribbonButton_PinNameRemoval.Name = "ribbonButton_PinNameRemoval";
            this.ribbonButton_PinNameRemoval.SmallImage = global::PmicAutomation.Properties.Resources._016A_grid;
            this.ribbonButton_PinNameRemoval.Text = "PinName Removal";
            this.ribbonButton_PinNameRemoval.Click += new System.EventHandler(this.RibbonButton_PinNameRemoval_Click);
            // 
            // nWireDefinition
            // 
            this.nWireDefinition.Items.Add(this.ribbonButton_nWireDefinition);
            this.nWireDefinition.Name = "nWireDefinition";
            this.nWireDefinition.Text = "";
            // 
            // ribbonButton_nWireDefinition
            // 
            this.ribbonButton_nWireDefinition.Image = global::PmicAutomation.Properties.Resources._031_pantone_1;
            this.ribbonButton_nWireDefinition.LargeImage = global::PmicAutomation.Properties.Resources._031_pantone_1;
            this.ribbonButton_nWireDefinition.Name = "ribbonButton_nWireDefinition";
            this.ribbonButton_nWireDefinition.SmallImage = global::PmicAutomation.Properties.Resources._031A_pantone_1;
            this.ribbonButton_nWireDefinition.Text = "nWire Definition";
            this.ribbonButton_nWireDefinition.Click += new System.EventHandler(this.RibbonButton_nWireDefinition_Click);
            // 
            // ribbonPanel_clbist
            // 
            this.ribbonPanel_clbist.Items.Add(this.ribbonButton_clbist);
            this.ribbonPanel_clbist.Name = "ribbonPanel_clbist";
            this.ribbonPanel_clbist.Text = "";
            // 
            // ribbonButton_clbist
            // 
            this.ribbonButton_clbist.Image = global::PmicAutomation.Properties.Resources._043_laptop;
            this.ribbonButton_clbist.LargeImage = global::PmicAutomation.Properties.Resources._043_laptop;
            this.ribbonButton_clbist.Name = "ribbonButton_clbist";
            this.ribbonButton_clbist.SmallImage = global::PmicAutomation.Properties.Resources._043_laptop;
            this.ribbonButton_clbist.Text = "CLVR Datalog Converter";
            this.ribbonButton_clbist.Click += new System.EventHandler(this.ribbonButton_clbist_Click);
            // 
            // ribbonPanel_GenTCMID
            // 
            this.ribbonPanel_GenTCMID.Items.Add(this.ribbonButton_GenTCMID);
            this.ribbonPanel_GenTCMID.Name = "ribbonPanel_GenTCMID";
            this.ribbonPanel_GenTCMID.Text = "";
            // 
            // ribbonButton_GenTCMID
            // 
            this.ribbonButton_GenTCMID.Image = ((System.Drawing.Image)(resources.GetObject("ribbonButton_GenTCMID.Image")));
            this.ribbonButton_GenTCMID.LargeImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_GenTCMID.LargeImage")));
            this.ribbonButton_GenTCMID.Name = "ribbonButton_GenTCMID";
            this.ribbonButton_GenTCMID.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_GenTCMID.SmallImage")));
            this.ribbonButton_GenTCMID.Text = "Gen TCMID";
            this.ribbonButton_GenTCMID.Click += new System.EventHandler(this.RibbonButton_GenTCMID_Click);
            // 
            // ribbonPanel_TCMIDComparator
            // 
            this.ribbonPanel_TCMIDComparator.Items.Add(this.ribbonButton_TCMIDComparator);
            this.ribbonPanel_TCMIDComparator.Name = "ribbonPanel_TCMIDComparator";
            this.ribbonPanel_TCMIDComparator.Text = "";
            // 
            // ribbonButton_TCMIDComparator
            // 
            this.ribbonButton_TCMIDComparator.Image = ((System.Drawing.Image)(resources.GetObject("ribbonButton_TCMIDComparator.Image")));
            this.ribbonButton_TCMIDComparator.LargeImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_TCMIDComparator.LargeImage")));
            this.ribbonButton_TCMIDComparator.Name = "ribbonButton_TCMIDComparator";
            this.ribbonButton_TCMIDComparator.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_TCMIDComparator.SmallImage")));
            this.ribbonButton_TCMIDComparator.Text = "TCMID Comparator";
            this.ribbonButton_TCMIDComparator.Click += new System.EventHandler(this.RibbonButton_TCMIDComparator_Click);
            // 
            // ribbonPanel_ProfileTool
            // 
            this.ribbonPanel_ProfileTool.Items.Add(this.ribbonButton_ProfileTool);
            this.ribbonPanel_ProfileTool.Name = "ribbonPanel_ProfileTool";
            this.ribbonPanel_ProfileTool.Text = "";
            // 
            // ribbonButton_ProfileTool
            // 
            this.ribbonButton_ProfileTool.Image = global::PmicAutomation.Properties.Resources._040_layers;
            this.ribbonButton_ProfileTool.LargeImage = global::PmicAutomation.Properties.Resources._040_layers;
            this.ribbonButton_ProfileTool.Name = "ribbonButton_ProfileTool";
            this.ribbonButton_ProfileTool.SmallImage = ((System.Drawing.Image)(resources.GetObject("ribbonButton_ProfileTool.SmallImage")));
            this.ribbonButton_ProfileTool.Text = "ProfileTool";
            this.ribbonButton_ProfileTool.Click += new System.EventHandler(this.ribbonButton_ProfileTool_Click);
            // 
            // ribbonButton1
            // 
            this.ribbonButton1.Image = global::PmicAutomation.Properties.Resources._031_pantone_1;
            this.ribbonButton1.LargeImage = global::PmicAutomation.Properties.Resources._031_pantone_1;
            this.ribbonButton1.Name = "ribbonButton1";
            this.ribbonButton1.SmallImage = global::PmicAutomation.Properties.Resources._031A_pantone_1;
            this.ribbonButton1.Text = "nWire Definition";
            // 
            // checkBox_ShowAll
            // 
            this.checkBox_ShowAll.AutoSize = true;
            this.checkBox_ShowAll.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.checkBox_ShowAll.Location = new System.Drawing.Point(770, 13);
            this.checkBox_ShowAll.Name = "checkBox_ShowAll";
            this.checkBox_ShowAll.Size = new System.Drawing.Size(15, 14);
            this.checkBox_ShowAll.TabIndex = 2;
            this.checkBox_ShowAll.UseVisualStyleBackColor = false;
            this.checkBox_ShowAll.CheckStateChanged += new System.EventHandler(this.CheckBoxShowAll_CheckedChanged);
            // 
            // ribbonPanel_AlarmCheck
            // 
            this.ribbonPanel_AlarmCheck.Items.Add(this.AlarmCheck);
            this.ribbonPanel_AlarmCheck.Name = "ribbonPanel_AlarmCheck";
            this.ribbonPanel_AlarmCheck.Text = "";
            // 
            // AlarmCheck
            // 
            this.AlarmCheck.Image = global::PmicAutomation.Properties.Resources._039_memory;
            this.AlarmCheck.LargeImage = global::PmicAutomation.Properties.Resources._039_memory;
            this.AlarmCheck.Name = "AlarmCheck";
            this.AlarmCheck.SmallImage = ((System.Drawing.Image)(resources.GetObject("AlarmCheck.SmallImage")));
            this.AlarmCheck.Text = "AlarmCheck";
            this.AlarmCheck.Click += new System.EventHandler(this.AlarmCheck_Click);
            // 
            // PmicMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(800, 207);
            this.Controls.Add(this.checkBox_ShowAll);
            this.Controls.Add(this.ribbon);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.HelpButton = true;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PmicMainForm";
            this.ShowIcon = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Ribbon ribbon;
        private System.Windows.Forms.RibbonTab ribbonTab_Utility;
        private System.Windows.Forms.RibbonPanel ribbonPane_PaParser;
        private System.Windows.Forms.RibbonButton ribbonButton_PaParser;
        private System.Windows.Forms.RibbonPanel ribbonPanel_VbtGenerator;
        private System.Windows.Forms.RibbonButton ribbonButton_VbtGenerator;
        private System.Windows.Forms.RibbonPanel ribbonPanel_Ahb_Enum;
        private System.Windows.Forms.RibbonButton ribbonButton_Ahb_Enum;
        private System.Windows.Forms.RibbonPanel ribbonPanel_TemplateGenerator;
        //private System.Windows.Forms.RibbonButton ribbonButton_TemplateGenerator;
        private System.Windows.Forms.RibbonPanel ribbonPanel_Relay;
        private System.Windows.Forms.RibbonButton RibbonButton_Relay;
        private System.Windows.Forms.RibbonPanel PinNameRemoval;
        private System.Windows.Forms.RibbonButton ribbonButton_PinNameRemoval;
        private System.Windows.Forms.RibbonPanel nWireDefinition;
        private System.Windows.Forms.RibbonButton ribbonButton_nWireDefinition;
        private System.Windows.Forms.RibbonPanel PatSetsAll;
        private System.Windows.Forms.RibbonButton ribbonButton_PatSets_All;
        private System.Windows.Forms.RibbonPanel OTPRegisterMap;
        private System.Windows.Forms.RibbonButton ribbonButton_OTPRegisterMap;
        private System.Windows.Forms.RibbonPanel ribbonPanel1;
        private System.Windows.Forms.RibbonButton ribbonButton_ExceptionHandling;
        private System.Windows.Forms.RibbonPanel ribbonPanel2;
        private System.Windows.Forms.RibbonButton ribbonButton_DataComparator;
        private System.Windows.Forms.RibbonTab ribbonTab_Tool;
        private System.Windows.Forms.RibbonPanel ribbonPanel_GenTCMID;
        private System.Windows.Forms.RibbonButton ribbonButton_GenTCMID;
        private System.Windows.Forms.RibbonButton ribbonButton1;
        private System.Windows.Forms.RibbonPanel ribbonPanel_ProfileTool;
        private System.Windows.Forms.RibbonButton ribbonButton_ProfileTool;
        private System.Windows.Forms.CheckBox checkBox_ShowAll;
        private System.Windows.Forms.RibbonPanel ribbonPanel_TCMIDComparator;
        private System.Windows.Forms.RibbonButton ribbonButton_TCMIDComparator;
        private System.Windows.Forms.RibbonPanel ribbonPanel_clbist;
        private System.Windows.Forms.RibbonButton ribbonButton_clbist;
        private System.Windows.Forms.RibbonPanel ribbonPanel_vbtpopprecheck;
        private System.Windows.Forms.RibbonButton ribbonButton_vbtpopprecheck;
        private System.Windows.Forms.RibbonPanel ribbonPanel_binoutcheck;
        private System.Windows.Forms.RibbonButton ribbonButton_binoutcheck;
        private System.Windows.Forms.RibbonPanel ribbonPanel_otpFilesComparing;
        private System.Windows.Forms.RibbonButton ribbonButton_otpFilesComparing;
        private System.Windows.Forms.RibbonPanel ribbonPanel_AlarmCheck;
        private System.Windows.Forms.RibbonButton AlarmCheck;
    }
}