using CommonLib.Controls;
using System.Windows.Forms;

namespace ProfileTool_PMIC
{
    partial class ProfileToolForm
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
            this.myStatus = new System.Windows.Forms.StatusStrip();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.checkBox_Excluding_By_Job = new System.Windows.Forms.CheckBox();
            this.checkBox_PowerPinOnly = new System.Windows.Forms.CheckBox();
            this.Button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxCurrent = new System.Windows.Forms.CheckBox();
            this.checkBoxVoltage = new System.Windows.Forms.CheckBox();
            this.groupBox = new System.Windows.Forms.GroupBox();
            this.radioButtonByFlow = new System.Windows.Forms.RadioButton();
            this.radioButtonByInstance = new System.Windows.Forms.RadioButton();
            this.ComboBox_ChanMap = new System.Windows.Forms.ComboBox();
            this.label_ChanMap = new System.Windows.Forms.Label();
            this.FileOpen_OutputPath1 = new CommonLib.Controls.MyFileOpen();
            this.FileOpen_CorePowerPins = new CommonLib.Controls.MyFileOpen();
            this.FileOpen_TestProgram = new CommonLib.Controls.MyFileOpen();
            this.FileOpen_ExecutionProfile = new CommonLib.Controls.MyFileOpen();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.checkBox_Legend = new System.Windows.Forms.CheckBox();
            this.checkBox_MultiPins = new System.Windows.Forms.CheckBox();
            this.checkBox_Power = new System.Windows.Forms.CheckBox();
            this.groupBoxChartGroupby = new System.Windows.Forms.GroupBox();
            this.radioButtonIndividual = new System.Windows.Forms.RadioButton();
            this.radioButtonMerge = new System.Windows.Forms.RadioButton();
            this.groupBoxFilter = new System.Windows.Forms.GroupBox();
            this.checkBoxOnlyLast = new System.Windows.Forms.CheckBox();
            this.labelSec = new System.Windows.Forms.Label();
            this.labelPulseWidth = new System.Windows.Forms.Label();
            this.labelStdev = new System.Windows.Forms.Label();
            this.labelLoopCount = new System.Windows.Forms.Label();
            this.textBoxStdev = new System.Windows.Forms.TextBox();
            this.textBoxPulseWidth = new System.Windows.Forms.TextBox();
            this.textBoxLoopCount = new System.Windows.Forms.TextBox();
            this.groupBoxChartCount = new System.Windows.Forms.GroupBox();
            this.textBoxChartCount = new System.Windows.Forms.TextBox();
            this.radioButtonAuto = new System.Windows.Forms.RadioButton();
            this.radioButtonManual = new System.Windows.Forms.RadioButton();
            this.Button2 = new System.Windows.Forms.Button();
            this.FileOpen_OutputPath2 = new CommonLib.Controls.MyFileOpen();
            this.FileOpen_ProfilePath2 = new CommonLib.Controls.MyFileOpen();
            this.FileOpen_ProfilePath1 = new CommonLib.Controls.MyFileOpen();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBoxChartGroupby.SuspendLayout();
            this.groupBoxFilter.SuspendLayout();
            this.groupBoxChartCount.SuspendLayout();
            this.SuspendLayout();
            // 
            // richTextBox
            // 
            this.richTextBox.Location = new System.Drawing.Point(30, 0);
            this.richTextBox.Margin = new System.Windows.Forms.Padding(30, 30, 30, 30);
            this.richTextBox.Size = new System.Drawing.Size(814, 273);
            // 
            // panel2
            // 
            this.panel2.Location = new System.Drawing.Point(0, 462);
            this.panel2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel2.Padding = new System.Windows.Forms.Padding(30, 0, 30, 30);
            this.panel2.Size = new System.Drawing.Size(874, 303);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel1.Padding = new System.Windows.Forms.Padding(30, 30, 30, 30);
            this.panel1.Size = new System.Drawing.Size(874, 462);
            // 
            // myStatus
            // 
            this.myStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.myStatus.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.myStatus.Location = new System.Drawing.Point(0, 560);
            this.myStatus.Name = "myStatus";
            this.myStatus.Size = new System.Drawing.Size(750, 22);
            this.myStatus.TabIndex = 8;
            this.myStatus.Text = "myStatus";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabControl1.Location = new System.Drawing.Point(30, 30);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(814, 402);
            this.tabControl1.TabIndex = 17;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.checkBox_Excluding_By_Job);
            this.tabPage1.Controls.Add(this.checkBox_PowerPinOnly);
            this.tabPage1.Controls.Add(this.Button1);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.groupBox);
            this.tabPage1.Controls.Add(this.ComboBox_ChanMap);
            this.tabPage1.Controls.Add(this.label_ChanMap);
            this.tabPage1.Controls.Add(this.FileOpen_OutputPath1);
            this.tabPage1.Controls.Add(this.FileOpen_CorePowerPins);
            this.tabPage1.Controls.Add(this.FileOpen_TestProgram);
            this.tabPage1.Controls.Add(this.FileOpen_ExecutionProfile);
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tabPage1.Size = new System.Drawing.Size(806, 371);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Test Program Modify";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // checkBox_Excluding_By_Job
            // 
            this.checkBox_Excluding_By_Job.AutoSize = true;
            this.checkBox_Excluding_By_Job.Checked = true;
            this.checkBox_Excluding_By_Job.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Excluding_By_Job.Location = new System.Drawing.Point(571, 245);
            this.checkBox_Excluding_By_Job.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBox_Excluding_By_Job.Name = "checkBox_Excluding_By_Job";
            this.checkBox_Excluding_By_Job.Size = new System.Drawing.Size(143, 22);
            this.checkBox_Excluding_By_Job.TabIndex = 0;
            this.checkBox_Excluding_By_Job.Text = "Excluding By Job";
            this.checkBox_Excluding_By_Job.UseVisualStyleBackColor = true;
            // 
            // checkBox_PowerPinOnly
            // 
            this.checkBox_PowerPinOnly.AutoSize = true;
            this.checkBox_PowerPinOnly.Location = new System.Drawing.Point(388, 245);
            this.checkBox_PowerPinOnly.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBox_PowerPinOnly.Name = "checkBox_PowerPinOnly";
            this.checkBox_PowerPinOnly.Size = new System.Drawing.Size(124, 22);
            this.checkBox_PowerPinOnly.TabIndex = 0;
            this.checkBox_PowerPinOnly.Text = "PowerPinOnly";
            this.checkBox_PowerPinOnly.UseVisualStyleBackColor = true;
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(571, 282);
            this.Button1.Margin = new System.Windows.Forms.Padding(0);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(206, 48);
            this.Button1.TabIndex = 20;
            this.Button1.Text = "Run";
            this.Button1.Click += new System.EventHandler(this.button_run1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxCurrent);
            this.groupBox1.Controls.Add(this.checkBoxVoltage);
            this.groupBox1.Location = new System.Drawing.Point(315, 280);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(206, 48);
            this.groupBox1.TabIndex = 24;
            this.groupBox1.TabStop = false;
            // 
            // checkBoxCurrent
            // 
            this.checkBoxCurrent.AutoSize = true;
            this.checkBoxCurrent.Checked = true;
            this.checkBoxCurrent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxCurrent.Location = new System.Drawing.Point(13, 16);
            this.checkBoxCurrent.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBoxCurrent.Name = "checkBoxCurrent";
            this.checkBoxCurrent.Size = new System.Drawing.Size(79, 22);
            this.checkBoxCurrent.TabIndex = 1;
            this.checkBoxCurrent.Text = "Current";
            this.checkBoxCurrent.UseVisualStyleBackColor = true;
            // 
            // checkBoxVoltage
            // 
            this.checkBoxVoltage.AutoSize = true;
            this.checkBoxVoltage.Location = new System.Drawing.Point(105, 16);
            this.checkBoxVoltage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBoxVoltage.Name = "checkBoxVoltage";
            this.checkBoxVoltage.Size = new System.Drawing.Size(79, 22);
            this.checkBoxVoltage.TabIndex = 0;
            this.checkBoxVoltage.Text = "Voltage";
            this.checkBoxVoltage.UseVisualStyleBackColor = true;
            // 
            // groupBox
            // 
            this.groupBox.Controls.Add(this.radioButtonByFlow);
            this.groupBox.Controls.Add(this.radioButtonByInstance);
            this.groupBox.Location = new System.Drawing.Point(39, 280);
            this.groupBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox.Name = "groupBox";
            this.groupBox.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox.Size = new System.Drawing.Size(230, 48);
            this.groupBox.TabIndex = 24;
            this.groupBox.TabStop = false;
            // 
            // radioButtonByFlow
            // 
            this.radioButtonByFlow.AutoSize = true;
            this.radioButtonByFlow.Checked = true;
            this.radioButtonByFlow.Location = new System.Drawing.Point(15, 16);
            this.radioButtonByFlow.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonByFlow.Name = "radioButtonByFlow";
            this.radioButtonByFlow.Size = new System.Drawing.Size(78, 22);
            this.radioButtonByFlow.TabIndex = 1;
            this.radioButtonByFlow.TabStop = true;
            this.radioButtonByFlow.Text = "ByFlow";
            this.radioButtonByFlow.UseVisualStyleBackColor = true;
            // 
            // radioButtonByInstance
            // 
            this.radioButtonByInstance.AutoSize = true;
            this.radioButtonByInstance.Location = new System.Drawing.Point(100, 16);
            this.radioButtonByInstance.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonByInstance.Name = "radioButtonByInstance";
            this.radioButtonByInstance.Size = new System.Drawing.Size(101, 22);
            this.radioButtonByInstance.TabIndex = 0;
            this.radioButtonByInstance.Text = "ByInstance";
            this.radioButtonByInstance.UseVisualStyleBackColor = true;
            // 
            // ComboBox_ChanMap
            // 
            this.ComboBox_ChanMap.FormattingEnabled = true;
            this.ComboBox_ChanMap.Location = new System.Drawing.Point(153, 240);
            this.ComboBox_ChanMap.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ComboBox_ChanMap.Name = "ComboBox_ChanMap";
            this.ComboBox_ChanMap.Size = new System.Drawing.Size(170, 26);
            this.ComboBox_ChanMap.TabIndex = 23;
            // 
            // label_ChanMap
            // 
            this.label_ChanMap.AutoSize = true;
            this.label_ChanMap.Location = new System.Drawing.Point(35, 242);
            this.label_ChanMap.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_ChanMap.Name = "label_ChanMap";
            this.label_ChanMap.Size = new System.Drawing.Size(95, 18);
            this.label_ChanMap.TabIndex = 22;
            this.label_ChanMap.Text = "Channel Map";
            // 
            // FileOpen_OutputPath1
            // 
            this.FileOpen_OutputPath1.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_OutputPath1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_OutputPath1.LebalText = "OutputPath";
            this.FileOpen_OutputPath1.Location = new System.Drawing.Point(4, 160);
            this.FileOpen_OutputPath1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_OutputPath1.Name = "FileOpen_OutputPath1";
            this.FileOpen_OutputPath1.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_OutputPath1.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_OutputPath1.TabIndex = 19;
            this.FileOpen_OutputPath1.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_Output1_Click);
            // 
            // FileOpen_CorePowerPins
            // 
            this.FileOpen_CorePowerPins.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_CorePowerPins.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_CorePowerPins.LebalText = "Core Power Pins(*.txt)";
            this.FileOpen_CorePowerPins.Location = new System.Drawing.Point(4, 108);
            this.FileOpen_CorePowerPins.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_CorePowerPins.Name = "FileOpen_CorePowerPins";
            this.FileOpen_CorePowerPins.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_CorePowerPins.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_CorePowerPins.TabIndex = 25;
            this.FileOpen_CorePowerPins.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_CorePowerList_Click);
            // 
            // FileOpen_TestProgram
            // 
            this.FileOpen_TestProgram.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_TestProgram.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_TestProgram.LebalText = "Test Program";
            this.FileOpen_TestProgram.Location = new System.Drawing.Point(4, 56);
            this.FileOpen_TestProgram.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_TestProgram.Name = "FileOpen_TestProgram";
            this.FileOpen_TestProgram.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_TestProgram.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_TestProgram.TabIndex = 17;
            this.FileOpen_TestProgram.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_TestProgarm_Click);
            // 
            // FileOpen_ExecutionProfile
            // 
            this.FileOpen_ExecutionProfile.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_ExecutionProfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_ExecutionProfile.LebalText = "Execution Profile(*.txt)";
            this.FileOpen_ExecutionProfile.Location = new System.Drawing.Point(4, 4);
            this.FileOpen_ExecutionProfile.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_ExecutionProfile.Name = "FileOpen_ExecutionProfile";
            this.FileOpen_ExecutionProfile.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_ExecutionProfile.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_ExecutionProfile.TabIndex = 18;
            this.FileOpen_ExecutionProfile.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_ExecutionProfile_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.checkBox_Legend);
            this.tabPage2.Controls.Add(this.checkBox_MultiPins);
            this.tabPage2.Controls.Add(this.checkBox_Power);
            this.tabPage2.Controls.Add(this.groupBoxChartGroupby);
            this.tabPage2.Controls.Add(this.groupBoxFilter);
            this.tabPage2.Controls.Add(this.groupBoxChartCount);
            this.tabPage2.Controls.Add(this.Button2);
            this.tabPage2.Controls.Add(this.FileOpen_OutputPath2);
            this.tabPage2.Controls.Add(this.FileOpen_ProfilePath2);
            this.tabPage2.Controls.Add(this.FileOpen_ProfilePath1);
            this.tabPage2.Location = new System.Drawing.Point(4, 27);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tabPage2.Size = new System.Drawing.Size(806, 371);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Data Analysis";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // checkBox_Legend
            // 
            this.checkBox_Legend.AutoSize = true;
            this.checkBox_Legend.Checked = true;
            this.checkBox_Legend.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Legend.Location = new System.Drawing.Point(570, 222);
            this.checkBox_Legend.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBox_Legend.Name = "checkBox_Legend";
            this.checkBox_Legend.Size = new System.Drawing.Size(199, 22);
            this.checkBox_Legend.TabIndex = 28;
            this.checkBox_Legend.Text = "Show Legend in MultiPins";
            this.checkBox_Legend.UseVisualStyleBackColor = true;
            // 
            // checkBox_MultiPins
            // 
            this.checkBox_MultiPins.AutoSize = true;
            this.checkBox_MultiPins.Checked = true;
            this.checkBox_MultiPins.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_MultiPins.Location = new System.Drawing.Point(570, 194);
            this.checkBox_MultiPins.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBox_MultiPins.Name = "checkBox_MultiPins";
            this.checkBox_MultiPins.Size = new System.Drawing.Size(162, 22);
            this.checkBox_MultiPins.TabIndex = 27;
            this.checkBox_MultiPins.Text = "Gen MultiPins Chart";
            this.checkBox_MultiPins.UseVisualStyleBackColor = true;
            // 
            // checkBox_Power
            // 
            this.checkBox_Power.AutoSize = true;
            this.checkBox_Power.Location = new System.Drawing.Point(570, 169);
            this.checkBox_Power.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBox_Power.Name = "checkBox_Power";
            this.checkBox_Power.Size = new System.Drawing.Size(126, 22);
            this.checkBox_Power.TabIndex = 4;
            this.checkBox_Power.Text = "ChartByPower";
            this.checkBox_Power.UseVisualStyleBackColor = true;
            this.checkBox_Power.CheckedChanged += new System.EventHandler(this.checkBox_Power_CheckedChanged);
            // 
            // groupBoxChartGroupby
            // 
            this.groupBoxChartGroupby.Controls.Add(this.radioButtonIndividual);
            this.groupBoxChartGroupby.Controls.Add(this.radioButtonMerge);
            this.groupBoxChartGroupby.Location = new System.Drawing.Point(19, 172);
            this.groupBoxChartGroupby.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxChartGroupby.Name = "groupBoxChartGroupby";
            this.groupBoxChartGroupby.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxChartGroupby.Size = new System.Drawing.Size(224, 60);
            this.groupBoxChartGroupby.TabIndex = 25;
            this.groupBoxChartGroupby.TabStop = false;
            this.groupBoxChartGroupby.Text = "Chart Group by";
            // 
            // radioButtonIndividual
            // 
            this.radioButtonIndividual.AutoSize = true;
            this.radioButtonIndividual.Checked = true;
            this.radioButtonIndividual.Location = new System.Drawing.Point(26, 24);
            this.radioButtonIndividual.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonIndividual.Name = "radioButtonIndividual";
            this.radioButtonIndividual.Size = new System.Drawing.Size(88, 22);
            this.radioButtonIndividual.TabIndex = 1;
            this.radioButtonIndividual.TabStop = true;
            this.radioButtonIndividual.Text = "Individual";
            this.radioButtonIndividual.UseVisualStyleBackColor = true;
            this.radioButtonIndividual.CheckedChanged += new System.EventHandler(this.radioButtonIndividual_CheckedChanged);
            // 
            // radioButtonMerge
            // 
            this.radioButtonMerge.AutoSize = true;
            this.radioButtonMerge.Location = new System.Drawing.Point(141, 24);
            this.radioButtonMerge.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonMerge.Name = "radioButtonMerge";
            this.radioButtonMerge.Size = new System.Drawing.Size(71, 22);
            this.radioButtonMerge.TabIndex = 1;
            this.radioButtonMerge.TabStop = true;
            this.radioButtonMerge.Text = "Merge";
            this.radioButtonMerge.UseVisualStyleBackColor = true;
            // 
            // groupBoxFilter
            // 
            this.groupBoxFilter.Controls.Add(this.checkBoxOnlyLast);
            this.groupBoxFilter.Controls.Add(this.labelSec);
            this.groupBoxFilter.Controls.Add(this.labelPulseWidth);
            this.groupBoxFilter.Controls.Add(this.labelStdev);
            this.groupBoxFilter.Controls.Add(this.labelLoopCount);
            this.groupBoxFilter.Controls.Add(this.textBoxStdev);
            this.groupBoxFilter.Controls.Add(this.textBoxPulseWidth);
            this.groupBoxFilter.Controls.Add(this.textBoxLoopCount);
            this.groupBoxFilter.Location = new System.Drawing.Point(19, 241);
            this.groupBoxFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxFilter.Name = "groupBoxFilter";
            this.groupBoxFilter.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxFilter.Size = new System.Drawing.Size(707, 60);
            this.groupBoxFilter.TabIndex = 25;
            this.groupBoxFilter.TabStop = false;
            this.groupBoxFilter.Text = "Filter Rule";
            // 
            // checkBoxOnlyLast
            // 
            this.checkBoxOnlyLast.AutoSize = true;
            this.checkBoxOnlyLast.Checked = true;
            this.checkBoxOnlyLast.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxOnlyLast.Location = new System.Drawing.Point(566, 25);
            this.checkBoxOnlyLast.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.checkBoxOnlyLast.Name = "checkBoxOnlyLast";
            this.checkBoxOnlyLast.Size = new System.Drawing.Size(92, 22);
            this.checkBoxOnlyLast.TabIndex = 4;
            this.checkBoxOnlyLast.Text = "Only Last";
            this.checkBoxOnlyLast.UseVisualStyleBackColor = true;
            // 
            // labelSec
            // 
            this.labelSec.AutoSize = true;
            this.labelSec.Location = new System.Drawing.Point(355, 28);
            this.labelSec.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelSec.Name = "labelSec";
            this.labelSec.Size = new System.Drawing.Size(34, 18);
            this.labelSec.TabIndex = 3;
            this.labelSec.Text = "Sec";
            // 
            // labelPulseWidth
            // 
            this.labelPulseWidth.AutoSize = true;
            this.labelPulseWidth.Location = new System.Drawing.Point(136, 28);
            this.labelPulseWidth.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelPulseWidth.Name = "labelPulseWidth";
            this.labelPulseWidth.Size = new System.Drawing.Size(82, 18);
            this.labelPulseWidth.TabIndex = 3;
            this.labelPulseWidth.Text = "Filter Width";
            // 
            // labelStdev
            // 
            this.labelStdev.AutoSize = true;
            this.labelStdev.Location = new System.Drawing.Point(402, 28);
            this.labelStdev.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelStdev.Name = "labelStdev";
            this.labelStdev.Size = new System.Drawing.Size(89, 18);
            this.labelStdev.TabIndex = 3;
            this.labelStdev.Text = "Spec(Stdev)";
            // 
            // labelLoopCount
            // 
            this.labelLoopCount.AutoSize = true;
            this.labelLoopCount.Location = new System.Drawing.Point(13, 28);
            this.labelLoopCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelLoopCount.Name = "labelLoopCount";
            this.labelLoopCount.Size = new System.Drawing.Size(69, 18);
            this.labelLoopCount.TabIndex = 3;
            this.labelLoopCount.Text = "Loop Cnt";
            // 
            // textBoxStdev
            // 
            this.textBoxStdev.Location = new System.Drawing.Point(500, 24);
            this.textBoxStdev.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxStdev.Name = "textBoxStdev";
            this.textBoxStdev.Size = new System.Drawing.Size(37, 24);
            this.textBoxStdev.TabIndex = 2;
            this.textBoxStdev.Text = "6";
            this.textBoxStdev.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxChartCount_KeyPress);
            // 
            // textBoxPulseWidth
            // 
            this.textBoxPulseWidth.Location = new System.Drawing.Point(230, 24);
            this.textBoxPulseWidth.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxPulseWidth.Name = "textBoxPulseWidth";
            this.textBoxPulseWidth.Size = new System.Drawing.Size(121, 24);
            this.textBoxPulseWidth.TabIndex = 2;
            this.textBoxPulseWidth.Text = "0.000015";
            this.textBoxPulseWidth.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBoxChartCount_KeyPressDouble);
            // 
            // textBoxLoopCount
            // 
            this.textBoxLoopCount.Location = new System.Drawing.Point(91, 24);
            this.textBoxLoopCount.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxLoopCount.Name = "textBoxLoopCount";
            this.textBoxLoopCount.Size = new System.Drawing.Size(37, 24);
            this.textBoxLoopCount.TabIndex = 2;
            this.textBoxLoopCount.Text = "0";
            this.textBoxLoopCount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxChartCount_KeyPress);
            // 
            // groupBoxChartCount
            // 
            this.groupBoxChartCount.Controls.Add(this.textBoxChartCount);
            this.groupBoxChartCount.Controls.Add(this.radioButtonAuto);
            this.groupBoxChartCount.Controls.Add(this.radioButtonManual);
            this.groupBoxChartCount.Location = new System.Drawing.Point(260, 172);
            this.groupBoxChartCount.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxChartCount.Name = "groupBoxChartCount";
            this.groupBoxChartCount.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBoxChartCount.Size = new System.Drawing.Size(256, 60);
            this.groupBoxChartCount.TabIndex = 25;
            this.groupBoxChartCount.TabStop = false;
            this.groupBoxChartCount.Text = "Chart Count Per Slide";
            // 
            // textBoxChartCount
            // 
            this.textBoxChartCount.Enabled = false;
            this.textBoxChartCount.Location = new System.Drawing.Point(198, 23);
            this.textBoxChartCount.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxChartCount.Name = "textBoxChartCount";
            this.textBoxChartCount.Size = new System.Drawing.Size(37, 24);
            this.textBoxChartCount.TabIndex = 2;
            this.textBoxChartCount.Text = "1";
            this.textBoxChartCount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxChartCount_KeyPress);
            // 
            // radioButtonAuto
            // 
            this.radioButtonAuto.AutoSize = true;
            this.radioButtonAuto.Checked = true;
            this.radioButtonAuto.Location = new System.Drawing.Point(26, 24);
            this.radioButtonAuto.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonAuto.Name = "radioButtonAuto";
            this.radioButtonAuto.Size = new System.Drawing.Size(59, 22);
            this.radioButtonAuto.TabIndex = 1;
            this.radioButtonAuto.TabStop = true;
            this.radioButtonAuto.Text = "Auto";
            this.radioButtonAuto.UseVisualStyleBackColor = true;
            // 
            // radioButtonManual
            // 
            this.radioButtonManual.AutoSize = true;
            this.radioButtonManual.Location = new System.Drawing.Point(103, 24);
            this.radioButtonManual.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.radioButtonManual.Name = "radioButtonManual";
            this.radioButtonManual.Size = new System.Drawing.Size(77, 22);
            this.radioButtonManual.TabIndex = 0;
            this.radioButtonManual.TabStop = true;
            this.radioButtonManual.Text = "Manual";
            this.radioButtonManual.UseVisualStyleBackColor = true;
            this.radioButtonManual.CheckedChanged += new System.EventHandler(this.radioButtonManual_CheckedChanged);
            // 
            // Button2
            // 
            this.Button2.Location = new System.Drawing.Point(458, 308);
            this.Button2.Margin = new System.Windows.Forms.Padding(0);
            this.Button2.Name = "Button2";
            this.Button2.Size = new System.Drawing.Size(321, 48);
            this.Button2.TabIndex = 24;
            this.Button2.Text = "Run";
            this.Button2.Click += new System.EventHandler(this.button_run2_Click);
            // 
            // FileOpen_OutputPath2
            // 
            this.FileOpen_OutputPath2.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_OutputPath2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_OutputPath2.LebalText = "OutputPath";
            this.FileOpen_OutputPath2.Location = new System.Drawing.Point(4, 108);
            this.FileOpen_OutputPath2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_OutputPath2.Name = "FileOpen_OutputPath2";
            this.FileOpen_OutputPath2.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_OutputPath2.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_OutputPath2.TabIndex = 23;
            this.FileOpen_OutputPath2.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_Output2_Click);
            // 
            // FileOpen_ProfilePath2
            // 
            this.FileOpen_ProfilePath2.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_ProfilePath2.Enabled = false;
            this.FileOpen_ProfilePath2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_ProfilePath2.LebalText = "Profile Path2";
            this.FileOpen_ProfilePath2.Location = new System.Drawing.Point(4, 56);
            this.FileOpen_ProfilePath2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_ProfilePath2.Name = "FileOpen_ProfilePath2";
            this.FileOpen_ProfilePath2.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_ProfilePath2.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_ProfilePath2.TabIndex = 26;
            this.FileOpen_ProfilePath2.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_profile_Click);
            // 
            // FileOpen_ProfilePath1
            // 
            this.FileOpen_ProfilePath1.Dock = System.Windows.Forms.DockStyle.Top;
            this.FileOpen_ProfilePath1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.FileOpen_ProfilePath1.LebalText = "Profile Path1";
            this.FileOpen_ProfilePath1.Location = new System.Drawing.Point(4, 4);
            this.FileOpen_ProfilePath1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.FileOpen_ProfilePath1.Name = "FileOpen_ProfilePath1";
            this.FileOpen_ProfilePath1.Padding = new System.Windows.Forms.Padding(13, 12, 13, 12);
            this.FileOpen_ProfilePath1.Size = new System.Drawing.Size(798, 52);
            this.FileOpen_ProfilePath1.TabIndex = 22;
            this.FileOpen_ProfilePath1.ButtonTextBoxButtonClick += new System.EventHandler(this.buttonLoad_profile_Click);
            // 
            // ProfileToolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(874, 790);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProfileToolForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Profile Tool";
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox.ResumeLayout(false);
            this.groupBox.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBoxChartGroupby.ResumeLayout(false);
            this.groupBoxChartGroupby.PerformLayout();
            this.groupBoxFilter.ResumeLayout(false);
            this.groupBoxFilter.PerformLayout();
            this.groupBoxChartCount.ResumeLayout(false);
            this.groupBoxChartCount.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private Button Button1;
        private System.Windows.Forms.GroupBox groupBox;
        public System.Windows.Forms.RadioButton radioButtonByFlow;
        public System.Windows.Forms.RadioButton radioButtonByInstance;
        public System.Windows.Forms.ComboBox ComboBox_ChanMap;
        private System.Windows.Forms.Label label_ChanMap;
        public MyFileOpen FileOpen_TestProgram;
        public MyFileOpen FileOpen_ExecutionProfile;

        private Button Button2;
        public MyFileOpen FileOpen_ProfilePath1;
        private System.Windows.Forms.GroupBox groupBoxChartGroupby;
        public System.Windows.Forms.RadioButton radioButtonIndividual;
        public System.Windows.Forms.RadioButton radioButtonMerge;
        private System.Windows.Forms.GroupBox groupBoxChartCount;
        public System.Windows.Forms.RadioButton radioButtonAuto;
        public System.Windows.Forms.RadioButton radioButtonManual;
        public System.Windows.Forms.TextBox textBoxChartCount;
        private System.Windows.Forms.GroupBox groupBoxFilter;
        public System.Windows.Forms.TextBox textBoxLoopCount;
        private System.Windows.Forms.Label labelLoopCount;
        private System.Windows.Forms.Label labelPulseWidth;
        public System.Windows.Forms.TextBox textBoxPulseWidth;
        public System.Windows.Forms.CheckBox checkBoxOnlyLast;
        private System.Windows.Forms.Label labelSec;
        private System.Windows.Forms.Label labelStdev;
        public System.Windows.Forms.TextBox textBoxStdev;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.CheckBox checkBoxCurrent;
        public System.Windows.Forms.CheckBox checkBoxVoltage;
        public MyFileOpen FileOpen_OutputPath2;
        public System.Windows.Forms.CheckBox checkBox_Power;
        public MyFileOpen FileOpen_ProfilePath2;
        public MyFileOpen FileOpen_CorePowerPins;
        public System.Windows.Forms.CheckBox checkBox_PowerPinOnly;
        public MyFileOpen FileOpen_OutputPath1;
        public System.Windows.Forms.CheckBox checkBox_MultiPins;
        public System.Windows.Forms.CheckBox checkBox_Legend;
        public CheckBox checkBox_Excluding_By_Job;
        private StatusStrip myStatus;
    }
}

