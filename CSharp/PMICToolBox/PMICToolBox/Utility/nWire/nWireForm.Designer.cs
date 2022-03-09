using PmicAutomation.Utility.nWire.Component;

namespace PmicAutomation.Utility.nWire
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.openFileDialog_Input = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog_Output = new System.Windows.Forms.FolderBrowserDialog();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.tabPage_Frame1 = new System.Windows.Forms.TabPage();
            this.groupBox_FrameName_Frame1 = new System.Windows.Forms.GroupBox();
            this.textBox_FrameName_Frame1 = new System.Windows.Forms.TextBox();
            this.groupBox_PatternFile_Frame1 = new System.Windows.Forms.GroupBox();
            this.textBox_PatternFile_Frame1 = new System.Windows.Forms.TextBox();
            this.button_SelectPatternFile_Frame1 = new System.Windows.Forms.Button();
            this.groupBox_FieldInfo_Frame1 = new System.Windows.Forms.GroupBox();
            this.dataGridView_FieldInfo_Frame1 = new System.Windows.Forms.DataGridView();
            this.Field_FieldName_Frame1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_PinName_Frame1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_Bits_Frame1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StartVector_Frame1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StopVector_Frame1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip_FieldInfo = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ItemAddField = new System.Windows.Forms.ToolStripMenuItem();
            this.ItemDeleteField = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox_TimeSetName = new System.Windows.Forms.GroupBox();
            this.textBox_TimeSetName = new System.Windows.Forms.TextBox();
            this.bt_Generate = new System.Windows.Forms.Button();
            this.tabControl_Frames = new System.Windows.Forms.TabControl();
            this.contextMenuStrip_tabControl = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ItemAddFrame = new System.Windows.Forms.ToolStripMenuItem();
            this.ItemDeleteFrame = new System.Windows.Forms.ToolStripMenuItem();
            this.tabPage_Frame2 = new System.Windows.Forms.TabPage();
            this.groupBox_FrameName_Frame2 = new System.Windows.Forms.GroupBox();
            this.textBox_FrameName_Frame2 = new System.Windows.Forms.TextBox();
            this.groupBox_PatternFile_Frame2 = new System.Windows.Forms.GroupBox();
            this.textBox_PatternFile_Frame2 = new System.Windows.Forms.TextBox();
            this.button_SelectPatternFile_Frame2 = new System.Windows.Forms.Button();
            this.groupBox_FieldInfo_Frame2 = new System.Windows.Forms.GroupBox();
            this.dataGridView_FieldInfo_Frame2 = new System.Windows.Forms.DataGridView();
            this.Field_FieldName_Frame2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_PinName_Frame2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_Bits_Frame2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StartVector_Frame2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StopVector_Frame2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage_Frame3 = new System.Windows.Forms.TabPage();
            this.groupBox_FrameName_Frame3 = new System.Windows.Forms.GroupBox();
            this.textBox_FrameName_Frame3 = new System.Windows.Forms.TextBox();
            this.groupBox_PatternFile_Frame3 = new System.Windows.Forms.GroupBox();
            this.textBox_PatternFile_Frame3 = new System.Windows.Forms.TextBox();
            this.button_SelectPatternFile_Frame3 = new System.Windows.Forms.Button();
            this.groupBox_FieldInfo_Frame3 = new System.Windows.Forms.GroupBox();
            this.dataGridView_FieldInfo_Frame3 = new System.Windows.Forms.DataGridView();
            this.Field_FieldName_Frame3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_PinName_Frame3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_Bits_Frame3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StartVector_Frame3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StopVector_Frame3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage_Frame4 = new System.Windows.Forms.TabPage();
            this.groupBox_FrameName_Frame4 = new System.Windows.Forms.GroupBox();
            this.textBox_FrameName_Frame4 = new System.Windows.Forms.TextBox();
            this.groupBox_PatternFile_Frame4 = new System.Windows.Forms.GroupBox();
            this.textBox_PatternFile_Frame4 = new System.Windows.Forms.TextBox();
            this.button_SelectPatternFile_Frame4 = new System.Windows.Forms.Button();
            this.groupBox_FieldInfo_Frame4 = new System.Windows.Forms.GroupBox();
            this.dataGridView_FieldInfo_Frame4 = new System.Windows.Forms.DataGridView();
            this.Field_FieldName_Frame4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_PinName_Frame4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_Bits_Frame4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StartVector_Frame4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StopVector_Frame4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabPage_Frame5 = new System.Windows.Forms.TabPage();
            this.groupBox_FrameName_Frame5 = new System.Windows.Forms.GroupBox();
            this.textBox_FrameName_Frame5 = new System.Windows.Forms.TextBox();
            this.groupBox_PatternFile_Frame5 = new System.Windows.Forms.GroupBox();
            this.textBox_PatternFile_Frame5 = new System.Windows.Forms.TextBox();
            this.button_SelectPatternFile_Frame5 = new System.Windows.Forms.Button();
            this.groupBox_FieldInfo_Frame5 = new System.Windows.Forms.GroupBox();
            this.dataGridView_FieldInfo_Frame5 = new System.Windows.Forms.DataGridView();
            this.Field_FieldName_Frame5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_PinName_Frame5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_Bits_Frame5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StartVector_Frame5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Field_StopVector_Frame5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox_PortPinMappingInfo = new System.Windows.Forms.GroupBox();
            this.comboBox_Protocol = new System.Windows.Forms.ComboBox();
            this.label_Protocol = new System.Windows.Forms.Label();
            this.dataGridView_PortPinMappingInfo = new System.Windows.Forms.DataGridView();
            this.Protocol_PortName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Protocol_PinName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox_OutputDir = new System.Windows.Forms.GroupBox();
            this.textBox_OutputDir = new System.Windows.Forms.TextBox();
            this.button_SelectOutputDir = new System.Windows.Forms.Button();
            this.progressBarEx_Process = new PmicAutomation.Utility.nWire.Component.ProgressBarEx();
            this.tabPage_Frame1.SuspendLayout();
            this.groupBox_FrameName_Frame1.SuspendLayout();
            this.groupBox_PatternFile_Frame1.SuspendLayout();
            this.groupBox_FieldInfo_Frame1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame1)).BeginInit();
            this.contextMenuStrip_FieldInfo.SuspendLayout();
            this.groupBox_TimeSetName.SuspendLayout();
            this.tabControl_Frames.SuspendLayout();
            this.contextMenuStrip_tabControl.SuspendLayout();
            this.tabPage_Frame2.SuspendLayout();
            this.groupBox_FrameName_Frame2.SuspendLayout();
            this.groupBox_PatternFile_Frame2.SuspendLayout();
            this.groupBox_FieldInfo_Frame2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame2)).BeginInit();
            this.tabPage_Frame3.SuspendLayout();
            this.groupBox_FrameName_Frame3.SuspendLayout();
            this.groupBox_PatternFile_Frame3.SuspendLayout();
            this.groupBox_FieldInfo_Frame3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame3)).BeginInit();
            this.tabPage_Frame4.SuspendLayout();
            this.groupBox_FrameName_Frame4.SuspendLayout();
            this.groupBox_PatternFile_Frame4.SuspendLayout();
            this.groupBox_FieldInfo_Frame4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame4)).BeginInit();
            this.tabPage_Frame5.SuspendLayout();
            this.groupBox_FrameName_Frame5.SuspendLayout();
            this.groupBox_PatternFile_Frame5.SuspendLayout();
            this.groupBox_FieldInfo_Frame5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame5)).BeginInit();
            this.groupBox_PortPinMappingInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_PortPinMappingInfo)).BeginInit();
            this.groupBox_OutputDir.SuspendLayout();
            this.SuspendLayout();
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // tabPage_Frame1
            // 
            this.tabPage_Frame1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage_Frame1.Controls.Add(this.groupBox_FrameName_Frame1);
            this.tabPage_Frame1.Controls.Add(this.groupBox_PatternFile_Frame1);
            this.tabPage_Frame1.Controls.Add(this.groupBox_FieldInfo_Frame1);
            this.tabPage_Frame1.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Frame1.Name = "tabPage_Frame1";
            this.tabPage_Frame1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Frame1.Size = new System.Drawing.Size(745, 295);
            this.tabPage_Frame1.TabIndex = 1;
            this.tabPage_Frame1.Text = "Frame1";
            // 
            // groupBox_FrameName_Frame1
            // 
            this.groupBox_FrameName_Frame1.Controls.Add(this.textBox_FrameName_Frame1);
            this.groupBox_FrameName_Frame1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FrameName_Frame1.Location = new System.Drawing.Point(6, 64);
            this.groupBox_FrameName_Frame1.Name = "groupBox_FrameName_Frame1";
            this.groupBox_FrameName_Frame1.Size = new System.Drawing.Size(732, 52);
            this.groupBox_FrameName_Frame1.TabIndex = 45;
            this.groupBox_FrameName_Frame1.TabStop = false;
            this.groupBox_FrameName_Frame1.Text = "Frame Name";
            // 
            // textBox_FrameName_Frame1
            // 
            this.textBox_FrameName_Frame1.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_FrameName_Frame1.Location = new System.Drawing.Point(6, 19);
            this.textBox_FrameName_Frame1.Name = "textBox_FrameName_Frame1";
            this.textBox_FrameName_Frame1.Size = new System.Drawing.Size(713, 20);
            this.textBox_FrameName_Frame1.TabIndex = 48;
            // 
            // groupBox_PatternFile_Frame1
            // 
            this.groupBox_PatternFile_Frame1.Controls.Add(this.textBox_PatternFile_Frame1);
            this.groupBox_PatternFile_Frame1.Controls.Add(this.button_SelectPatternFile_Frame1);
            this.groupBox_PatternFile_Frame1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PatternFile_Frame1.Location = new System.Drawing.Point(6, 9);
            this.groupBox_PatternFile_Frame1.Name = "groupBox_PatternFile_Frame1";
            this.groupBox_PatternFile_Frame1.Size = new System.Drawing.Size(732, 52);
            this.groupBox_PatternFile_Frame1.TabIndex = 44;
            this.groupBox_PatternFile_Frame1.TabStop = false;
            this.groupBox_PatternFile_Frame1.Text = "Pattern File";
            // 
            // textBox_PatternFile_Frame1
            // 
            this.textBox_PatternFile_Frame1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_PatternFile_Frame1.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_PatternFile_Frame1.Location = new System.Drawing.Point(6, 20);
            this.textBox_PatternFile_Frame1.Name = "textBox_PatternFile_Frame1";
            this.textBox_PatternFile_Frame1.Size = new System.Drawing.Size(620, 20);
            this.textBox_PatternFile_Frame1.TabIndex = 41;
            // 
            // button_SelectPatternFile_Frame1
            // 
            this.button_SelectPatternFile_Frame1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectPatternFile_Frame1.Location = new System.Drawing.Point(632, 20);
            this.button_SelectPatternFile_Frame1.Name = "button_SelectPatternFile_Frame1";
            this.button_SelectPatternFile_Frame1.Size = new System.Drawing.Size(87, 23);
            this.button_SelectPatternFile_Frame1.TabIndex = 2;
            this.button_SelectPatternFile_Frame1.Text = "Select";
            this.button_SelectPatternFile_Frame1.UseVisualStyleBackColor = true;
            this.button_SelectPatternFile_Frame1.Click += new System.EventHandler(this.bt_SelectPatternFile_Click);
            // 
            // groupBox_FieldInfo_Frame1
            // 
            this.groupBox_FieldInfo_Frame1.Controls.Add(this.dataGridView_FieldInfo_Frame1);
            this.groupBox_FieldInfo_Frame1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FieldInfo_Frame1.Location = new System.Drawing.Point(6, 118);
            this.groupBox_FieldInfo_Frame1.Name = "groupBox_FieldInfo_Frame1";
            this.groupBox_FieldInfo_Frame1.Size = new System.Drawing.Size(732, 167);
            this.groupBox_FieldInfo_Frame1.TabIndex = 45;
            this.groupBox_FieldInfo_Frame1.TabStop = false;
            this.groupBox_FieldInfo_Frame1.Text = "Field Info";
            // 
            // dataGridView_FieldInfo_Frame1
            // 
            this.dataGridView_FieldInfo_Frame1.AllowUserToAddRows = false;
            this.dataGridView_FieldInfo_Frame1.AllowUserToResizeColumns = false;
            this.dataGridView_FieldInfo_Frame1.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_FieldInfo_Frame1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_FieldInfo_Frame1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Field_FieldName_Frame1,
            this.Field_PinName_Frame1,
            this.Field_Bits_Frame1,
            this.Field_StartVector_Frame1,
            this.Field_StopVector_Frame1});
            this.dataGridView_FieldInfo_Frame1.ContextMenuStrip = this.contextMenuStrip_FieldInfo;
            this.dataGridView_FieldInfo_Frame1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_FieldInfo_Frame1.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_FieldInfo_Frame1.Location = new System.Drawing.Point(6, 17);
            this.dataGridView_FieldInfo_Frame1.MultiSelect = false;
            this.dataGridView_FieldInfo_Frame1.Name = "dataGridView_FieldInfo_Frame1";
            this.dataGridView_FieldInfo_Frame1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_FieldInfo_Frame1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_FieldInfo_Frame1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_FieldInfo_Frame1.Size = new System.Drawing.Size(713, 136);
            this.dataGridView_FieldInfo_Frame1.TabIndex = 48;
            // 
            // Field_FieldName_Frame1
            // 
            this.Field_FieldName_Frame1.Frozen = true;
            this.Field_FieldName_Frame1.HeaderText = "Field Name";
            this.Field_FieldName_Frame1.MinimumWidth = 200;
            this.Field_FieldName_Frame1.Name = "Field_FieldName_Frame1";
            this.Field_FieldName_Frame1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_FieldName_Frame1.Width = 200;
            // 
            // Field_PinName_Frame1
            // 
            this.Field_PinName_Frame1.Frozen = true;
            this.Field_PinName_Frame1.HeaderText = "Pin Name";
            this.Field_PinName_Frame1.MinimumWidth = 190;
            this.Field_PinName_Frame1.Name = "Field_PinName_Frame1";
            this.Field_PinName_Frame1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_PinName_Frame1.Width = 190;
            // 
            // Field_Bits_Frame1
            // 
            dataGridViewCellStyle2.Format = "N0";
            this.Field_Bits_Frame1.DefaultCellStyle = dataGridViewCellStyle2;
            this.Field_Bits_Frame1.Frozen = true;
            this.Field_Bits_Frame1.HeaderText = "Bits";
            this.Field_Bits_Frame1.MaxInputLength = 5;
            this.Field_Bits_Frame1.Name = "Field_Bits_Frame1";
            this.Field_Bits_Frame1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_Bits_Frame1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_Bits_Frame1.Width = 60;
            // 
            // Field_StartVector_Frame1
            // 
            dataGridViewCellStyle3.Format = "N0";
            dataGridViewCellStyle3.NullValue = null;
            this.Field_StartVector_Frame1.DefaultCellStyle = dataGridViewCellStyle3;
            this.Field_StartVector_Frame1.Frozen = true;
            this.Field_StartVector_Frame1.HeaderText = "Start Vector";
            this.Field_StartVector_Frame1.MaxInputLength = 5;
            this.Field_StartVector_Frame1.Name = "Field_StartVector_Frame1";
            this.Field_StartVector_Frame1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StartVector_Frame1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StartVector_Frame1.Width = 110;
            // 
            // Field_StopVector_Frame1
            // 
            dataGridViewCellStyle4.Format = "N0";
            dataGridViewCellStyle4.NullValue = null;
            this.Field_StopVector_Frame1.DefaultCellStyle = dataGridViewCellStyle4;
            this.Field_StopVector_Frame1.Frozen = true;
            this.Field_StopVector_Frame1.HeaderText = "Stop Vector";
            this.Field_StopVector_Frame1.MaxInputLength = 5;
            this.Field_StopVector_Frame1.Name = "Field_StopVector_Frame1";
            this.Field_StopVector_Frame1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StopVector_Frame1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StopVector_Frame1.Width = 110;
            // 
            // contextMenuStrip_FieldInfo
            // 
            this.contextMenuStrip_FieldInfo.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip_FieldInfo.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ItemAddField,
            this.ItemDeleteField});
            this.contextMenuStrip_FieldInfo.Name = "contextMenuStrip1";
            this.contextMenuStrip_FieldInfo.Size = new System.Drawing.Size(142, 48);
            this.contextMenuStrip_FieldInfo.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.contextMenuStrip_FieldInfo_ItemClicked);
            // 
            // ItemAddField
            // 
            this.ItemAddField.Name = "ItemAddField";
            this.ItemAddField.Size = new System.Drawing.Size(141, 22);
            this.ItemAddField.Text = "Add Field";
            // 
            // ItemDeleteField
            // 
            this.ItemDeleteField.Name = "ItemDeleteField";
            this.ItemDeleteField.Size = new System.Drawing.Size(141, 22);
            this.ItemDeleteField.Text = "Delete Field";
            // 
            // groupBox_TimeSetName
            // 
            this.groupBox_TimeSetName.Controls.Add(this.textBox_TimeSetName);
            this.groupBox_TimeSetName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_TimeSetName.Location = new System.Drawing.Point(468, 335);
            this.groupBox_TimeSetName.Name = "groupBox_TimeSetName";
            this.groupBox_TimeSetName.Size = new System.Drawing.Size(293, 51);
            this.groupBox_TimeSetName.TabIndex = 46;
            this.groupBox_TimeSetName.TabStop = false;
            this.groupBox_TimeSetName.Text = "TimeSet Name in Pattern";
            // 
            // textBox_TimeSetName
            // 
            this.textBox_TimeSetName.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_TimeSetName.Location = new System.Drawing.Point(6, 19);
            this.textBox_TimeSetName.Name = "textBox_TimeSetName";
            this.textBox_TimeSetName.Size = new System.Drawing.Size(281, 20);
            this.textBox_TimeSetName.TabIndex = 47;
            // 
            // bt_Generate
            // 
            this.bt_Generate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.bt_Generate.Location = new System.Drawing.Point(468, 396);
            this.bt_Generate.Name = "bt_Generate";
            this.bt_Generate.Size = new System.Drawing.Size(293, 126);
            this.bt_Generate.TabIndex = 43;
            this.bt_Generate.Text = "Generate";
            this.bt_Generate.UseVisualStyleBackColor = true;
            this.bt_Generate.Click += new System.EventHandler(this.bt_Generate_Click);
            // 
            // tabControl_Frames
            // 
            this.tabControl_Frames.ContextMenuStrip = this.contextMenuStrip_tabControl;
            this.tabControl_Frames.Controls.Add(this.tabPage_Frame1);
            this.tabControl_Frames.Controls.Add(this.tabPage_Frame2);
            this.tabControl_Frames.Controls.Add(this.tabPage_Frame3);
            this.tabControl_Frames.Controls.Add(this.tabPage_Frame4);
            this.tabControl_Frames.Controls.Add(this.tabPage_Frame5);
            this.tabControl_Frames.Location = new System.Drawing.Point(12, 12);
            this.tabControl_Frames.Name = "tabControl_Frames";
            this.tabControl_Frames.SelectedIndex = 0;
            this.tabControl_Frames.Size = new System.Drawing.Size(753, 321);
            this.tabControl_Frames.TabIndex = 43;
            // 
            // contextMenuStrip_tabControl
            // 
            this.contextMenuStrip_tabControl.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip_tabControl.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ItemAddFrame,
            this.ItemDeleteFrame});
            this.contextMenuStrip_tabControl.Name = "contextMenuStrip_tabControl";
            this.contextMenuStrip_tabControl.Size = new System.Drawing.Size(150, 48);
            this.contextMenuStrip_tabControl.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.contextMenuStrip_tabControl_ItemClicked);
            // 
            // ItemAddFrame
            // 
            this.ItemAddFrame.Name = "ItemAddFrame";
            this.ItemAddFrame.Size = new System.Drawing.Size(149, 22);
            this.ItemAddFrame.Text = "Add Frame";
            // 
            // ItemDeleteFrame
            // 
            this.ItemDeleteFrame.Name = "ItemDeleteFrame";
            this.ItemDeleteFrame.Size = new System.Drawing.Size(149, 22);
            this.ItemDeleteFrame.Text = "Delete Frame";
            // 
            // tabPage_Frame2
            // 
            this.tabPage_Frame2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage_Frame2.Controls.Add(this.groupBox_FrameName_Frame2);
            this.tabPage_Frame2.Controls.Add(this.groupBox_PatternFile_Frame2);
            this.tabPage_Frame2.Controls.Add(this.groupBox_FieldInfo_Frame2);
            this.tabPage_Frame2.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Frame2.Name = "tabPage_Frame2";
            this.tabPage_Frame2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Frame2.Size = new System.Drawing.Size(745, 295);
            this.tabPage_Frame2.TabIndex = 2;
            this.tabPage_Frame2.Text = "Frame2";
            // 
            // groupBox_FrameName_Frame2
            // 
            this.groupBox_FrameName_Frame2.Controls.Add(this.textBox_FrameName_Frame2);
            this.groupBox_FrameName_Frame2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FrameName_Frame2.Location = new System.Drawing.Point(6, 64);
            this.groupBox_FrameName_Frame2.Name = "groupBox_FrameName_Frame2";
            this.groupBox_FrameName_Frame2.Size = new System.Drawing.Size(732, 52);
            this.groupBox_FrameName_Frame2.TabIndex = 47;
            this.groupBox_FrameName_Frame2.TabStop = false;
            this.groupBox_FrameName_Frame2.Text = "Frame Name";
            // 
            // textBox_FrameName_Frame2
            // 
            this.textBox_FrameName_Frame2.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_FrameName_Frame2.Location = new System.Drawing.Point(6, 19);
            this.textBox_FrameName_Frame2.Name = "textBox_FrameName_Frame2";
            this.textBox_FrameName_Frame2.Size = new System.Drawing.Size(713, 20);
            this.textBox_FrameName_Frame2.TabIndex = 48;
            // 
            // groupBox_PatternFile_Frame2
            // 
            this.groupBox_PatternFile_Frame2.Controls.Add(this.textBox_PatternFile_Frame2);
            this.groupBox_PatternFile_Frame2.Controls.Add(this.button_SelectPatternFile_Frame2);
            this.groupBox_PatternFile_Frame2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PatternFile_Frame2.Location = new System.Drawing.Point(6, 9);
            this.groupBox_PatternFile_Frame2.Name = "groupBox_PatternFile_Frame2";
            this.groupBox_PatternFile_Frame2.Size = new System.Drawing.Size(732, 52);
            this.groupBox_PatternFile_Frame2.TabIndex = 46;
            this.groupBox_PatternFile_Frame2.TabStop = false;
            this.groupBox_PatternFile_Frame2.Text = "Pattern File";
            // 
            // textBox_PatternFile_Frame2
            // 
            this.textBox_PatternFile_Frame2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_PatternFile_Frame2.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_PatternFile_Frame2.Location = new System.Drawing.Point(6, 20);
            this.textBox_PatternFile_Frame2.Name = "textBox_PatternFile_Frame2";
            this.textBox_PatternFile_Frame2.Size = new System.Drawing.Size(620, 20);
            this.textBox_PatternFile_Frame2.TabIndex = 41;
            // 
            // button_SelectPatternFile_Frame2
            // 
            this.button_SelectPatternFile_Frame2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectPatternFile_Frame2.Location = new System.Drawing.Point(632, 20);
            this.button_SelectPatternFile_Frame2.Name = "button_SelectPatternFile_Frame2";
            this.button_SelectPatternFile_Frame2.Size = new System.Drawing.Size(87, 23);
            this.button_SelectPatternFile_Frame2.TabIndex = 2;
            this.button_SelectPatternFile_Frame2.Text = "Select";
            this.button_SelectPatternFile_Frame2.UseVisualStyleBackColor = true;
            this.button_SelectPatternFile_Frame2.Click += new System.EventHandler(this.bt_SelectPatternFile_Click);
            // 
            // groupBox_FieldInfo_Frame2
            // 
            this.groupBox_FieldInfo_Frame2.Controls.Add(this.dataGridView_FieldInfo_Frame2);
            this.groupBox_FieldInfo_Frame2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FieldInfo_Frame2.Location = new System.Drawing.Point(6, 118);
            this.groupBox_FieldInfo_Frame2.Name = "groupBox_FieldInfo_Frame2";
            this.groupBox_FieldInfo_Frame2.Size = new System.Drawing.Size(732, 167);
            this.groupBox_FieldInfo_Frame2.TabIndex = 48;
            this.groupBox_FieldInfo_Frame2.TabStop = false;
            this.groupBox_FieldInfo_Frame2.Text = "Field Info";
            // 
            // dataGridView_FieldInfo_Frame2
            // 
            this.dataGridView_FieldInfo_Frame2.AllowUserToAddRows = false;
            this.dataGridView_FieldInfo_Frame2.AllowUserToResizeColumns = false;
            this.dataGridView_FieldInfo_Frame2.AllowUserToResizeRows = false;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_FieldInfo_Frame2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridView_FieldInfo_Frame2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Field_FieldName_Frame2,
            this.Field_PinName_Frame2,
            this.Field_Bits_Frame2,
            this.Field_StartVector_Frame2,
            this.Field_StopVector_Frame2});
            this.dataGridView_FieldInfo_Frame2.ContextMenuStrip = this.contextMenuStrip_FieldInfo;
            this.dataGridView_FieldInfo_Frame2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_FieldInfo_Frame2.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_FieldInfo_Frame2.Location = new System.Drawing.Point(6, 17);
            this.dataGridView_FieldInfo_Frame2.MultiSelect = false;
            this.dataGridView_FieldInfo_Frame2.Name = "dataGridView_FieldInfo_Frame2";
            this.dataGridView_FieldInfo_Frame2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame2.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_FieldInfo_Frame2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_FieldInfo_Frame2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_FieldInfo_Frame2.Size = new System.Drawing.Size(713, 136);
            this.dataGridView_FieldInfo_Frame2.TabIndex = 48;
            // 
            // Field_FieldName_Frame2
            // 
            this.Field_FieldName_Frame2.Frozen = true;
            this.Field_FieldName_Frame2.HeaderText = "Field Name";
            this.Field_FieldName_Frame2.MinimumWidth = 200;
            this.Field_FieldName_Frame2.Name = "Field_FieldName_Frame2";
            this.Field_FieldName_Frame2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_FieldName_Frame2.Width = 200;
            // 
            // Field_PinName_Frame2
            // 
            this.Field_PinName_Frame2.Frozen = true;
            this.Field_PinName_Frame2.HeaderText = "Pin Name";
            this.Field_PinName_Frame2.MinimumWidth = 190;
            this.Field_PinName_Frame2.Name = "Field_PinName_Frame2";
            this.Field_PinName_Frame2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_PinName_Frame2.Width = 190;
            // 
            // Field_Bits_Frame2
            // 
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            this.Field_Bits_Frame2.DefaultCellStyle = dataGridViewCellStyle6;
            this.Field_Bits_Frame2.Frozen = true;
            this.Field_Bits_Frame2.HeaderText = "Bits";
            this.Field_Bits_Frame2.MaxInputLength = 5;
            this.Field_Bits_Frame2.Name = "Field_Bits_Frame2";
            this.Field_Bits_Frame2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_Bits_Frame2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_Bits_Frame2.Width = 60;
            // 
            // Field_StartVector_Frame2
            // 
            dataGridViewCellStyle7.Format = "N0";
            dataGridViewCellStyle7.NullValue = null;
            this.Field_StartVector_Frame2.DefaultCellStyle = dataGridViewCellStyle7;
            this.Field_StartVector_Frame2.Frozen = true;
            this.Field_StartVector_Frame2.HeaderText = "Start Vector";
            this.Field_StartVector_Frame2.MaxInputLength = 5;
            this.Field_StartVector_Frame2.Name = "Field_StartVector_Frame2";
            this.Field_StartVector_Frame2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StartVector_Frame2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StartVector_Frame2.Width = 110;
            // 
            // Field_StopVector_Frame2
            // 
            dataGridViewCellStyle8.Format = "N0";
            dataGridViewCellStyle8.NullValue = null;
            this.Field_StopVector_Frame2.DefaultCellStyle = dataGridViewCellStyle8;
            this.Field_StopVector_Frame2.Frozen = true;
            this.Field_StopVector_Frame2.HeaderText = "Stop Vector";
            this.Field_StopVector_Frame2.MaxInputLength = 5;
            this.Field_StopVector_Frame2.Name = "Field_StopVector_Frame2";
            this.Field_StopVector_Frame2.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StopVector_Frame2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StopVector_Frame2.Width = 110;
            // 
            // tabPage_Frame3
            // 
            this.tabPage_Frame3.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage_Frame3.Controls.Add(this.groupBox_FrameName_Frame3);
            this.tabPage_Frame3.Controls.Add(this.groupBox_PatternFile_Frame3);
            this.tabPage_Frame3.Controls.Add(this.groupBox_FieldInfo_Frame3);
            this.tabPage_Frame3.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Frame3.Name = "tabPage_Frame3";
            this.tabPage_Frame3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Frame3.Size = new System.Drawing.Size(745, 295);
            this.tabPage_Frame3.TabIndex = 3;
            this.tabPage_Frame3.Text = "Frame3";
            // 
            // groupBox_FrameName_Frame3
            // 
            this.groupBox_FrameName_Frame3.Controls.Add(this.textBox_FrameName_Frame3);
            this.groupBox_FrameName_Frame3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FrameName_Frame3.Location = new System.Drawing.Point(6, 64);
            this.groupBox_FrameName_Frame3.Name = "groupBox_FrameName_Frame3";
            this.groupBox_FrameName_Frame3.Size = new System.Drawing.Size(732, 52);
            this.groupBox_FrameName_Frame3.TabIndex = 47;
            this.groupBox_FrameName_Frame3.TabStop = false;
            this.groupBox_FrameName_Frame3.Text = "Frame Name";
            // 
            // textBox_FrameName_Frame3
            // 
            this.textBox_FrameName_Frame3.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_FrameName_Frame3.Location = new System.Drawing.Point(6, 19);
            this.textBox_FrameName_Frame3.Name = "textBox_FrameName_Frame3";
            this.textBox_FrameName_Frame3.Size = new System.Drawing.Size(713, 20);
            this.textBox_FrameName_Frame3.TabIndex = 48;
            // 
            // groupBox_PatternFile_Frame3
            // 
            this.groupBox_PatternFile_Frame3.Controls.Add(this.textBox_PatternFile_Frame3);
            this.groupBox_PatternFile_Frame3.Controls.Add(this.button_SelectPatternFile_Frame3);
            this.groupBox_PatternFile_Frame3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PatternFile_Frame3.Location = new System.Drawing.Point(6, 9);
            this.groupBox_PatternFile_Frame3.Name = "groupBox_PatternFile_Frame3";
            this.groupBox_PatternFile_Frame3.Size = new System.Drawing.Size(732, 52);
            this.groupBox_PatternFile_Frame3.TabIndex = 46;
            this.groupBox_PatternFile_Frame3.TabStop = false;
            this.groupBox_PatternFile_Frame3.Text = "Pattern File";
            // 
            // textBox_PatternFile_Frame3
            // 
            this.textBox_PatternFile_Frame3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_PatternFile_Frame3.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_PatternFile_Frame3.Location = new System.Drawing.Point(6, 20);
            this.textBox_PatternFile_Frame3.Name = "textBox_PatternFile_Frame3";
            this.textBox_PatternFile_Frame3.Size = new System.Drawing.Size(620, 20);
            this.textBox_PatternFile_Frame3.TabIndex = 41;
            // 
            // button_SelectPatternFile_Frame3
            // 
            this.button_SelectPatternFile_Frame3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectPatternFile_Frame3.Location = new System.Drawing.Point(632, 20);
            this.button_SelectPatternFile_Frame3.Name = "button_SelectPatternFile_Frame3";
            this.button_SelectPatternFile_Frame3.Size = new System.Drawing.Size(87, 23);
            this.button_SelectPatternFile_Frame3.TabIndex = 2;
            this.button_SelectPatternFile_Frame3.Text = "Select";
            this.button_SelectPatternFile_Frame3.UseVisualStyleBackColor = true;
            this.button_SelectPatternFile_Frame3.Click += new System.EventHandler(this.bt_SelectPatternFile_Click);
            // 
            // groupBox_FieldInfo_Frame3
            // 
            this.groupBox_FieldInfo_Frame3.Controls.Add(this.dataGridView_FieldInfo_Frame3);
            this.groupBox_FieldInfo_Frame3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FieldInfo_Frame3.Location = new System.Drawing.Point(6, 118);
            this.groupBox_FieldInfo_Frame3.Name = "groupBox_FieldInfo_Frame3";
            this.groupBox_FieldInfo_Frame3.Size = new System.Drawing.Size(732, 167);
            this.groupBox_FieldInfo_Frame3.TabIndex = 48;
            this.groupBox_FieldInfo_Frame3.TabStop = false;
            this.groupBox_FieldInfo_Frame3.Text = "Field Info";
            // 
            // dataGridView_FieldInfo_Frame3
            // 
            this.dataGridView_FieldInfo_Frame3.AllowUserToAddRows = false;
            this.dataGridView_FieldInfo_Frame3.AllowUserToResizeColumns = false;
            this.dataGridView_FieldInfo_Frame3.AllowUserToResizeRows = false;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_FieldInfo_Frame3.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.dataGridView_FieldInfo_Frame3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Field_FieldName_Frame3,
            this.Field_PinName_Frame3,
            this.Field_Bits_Frame3,
            this.Field_StartVector_Frame3,
            this.Field_StopVector_Frame3});
            this.dataGridView_FieldInfo_Frame3.ContextMenuStrip = this.contextMenuStrip_FieldInfo;
            this.dataGridView_FieldInfo_Frame3.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_FieldInfo_Frame3.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_FieldInfo_Frame3.Location = new System.Drawing.Point(6, 17);
            this.dataGridView_FieldInfo_Frame3.MultiSelect = false;
            this.dataGridView_FieldInfo_Frame3.Name = "dataGridView_FieldInfo_Frame3";
            this.dataGridView_FieldInfo_Frame3.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame3.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_FieldInfo_Frame3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_FieldInfo_Frame3.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_FieldInfo_Frame3.Size = new System.Drawing.Size(713, 136);
            this.dataGridView_FieldInfo_Frame3.TabIndex = 48;
            // 
            // Field_FieldName_Frame3
            // 
            this.Field_FieldName_Frame3.Frozen = true;
            this.Field_FieldName_Frame3.HeaderText = "Field Name";
            this.Field_FieldName_Frame3.MinimumWidth = 200;
            this.Field_FieldName_Frame3.Name = "Field_FieldName_Frame3";
            this.Field_FieldName_Frame3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_FieldName_Frame3.Width = 200;
            // 
            // Field_PinName_Frame3
            // 
            this.Field_PinName_Frame3.Frozen = true;
            this.Field_PinName_Frame3.HeaderText = "Pin Name";
            this.Field_PinName_Frame3.MinimumWidth = 190;
            this.Field_PinName_Frame3.Name = "Field_PinName_Frame3";
            this.Field_PinName_Frame3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_PinName_Frame3.Width = 190;
            // 
            // Field_Bits_Frame3
            // 
            dataGridViewCellStyle10.Format = "N0";
            dataGridViewCellStyle10.NullValue = null;
            this.Field_Bits_Frame3.DefaultCellStyle = dataGridViewCellStyle10;
            this.Field_Bits_Frame3.Frozen = true;
            this.Field_Bits_Frame3.HeaderText = "Bits";
            this.Field_Bits_Frame3.MaxInputLength = 5;
            this.Field_Bits_Frame3.Name = "Field_Bits_Frame3";
            this.Field_Bits_Frame3.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_Bits_Frame3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_Bits_Frame3.Width = 60;
            // 
            // Field_StartVector_Frame3
            // 
            dataGridViewCellStyle11.Format = "N0";
            dataGridViewCellStyle11.NullValue = null;
            this.Field_StartVector_Frame3.DefaultCellStyle = dataGridViewCellStyle11;
            this.Field_StartVector_Frame3.Frozen = true;
            this.Field_StartVector_Frame3.HeaderText = "Start Vector";
            this.Field_StartVector_Frame3.MaxInputLength = 5;
            this.Field_StartVector_Frame3.Name = "Field_StartVector_Frame3";
            this.Field_StartVector_Frame3.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StartVector_Frame3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StartVector_Frame3.Width = 110;
            // 
            // Field_StopVector_Frame3
            // 
            dataGridViewCellStyle12.Format = "N0";
            dataGridViewCellStyle12.NullValue = null;
            this.Field_StopVector_Frame3.DefaultCellStyle = dataGridViewCellStyle12;
            this.Field_StopVector_Frame3.Frozen = true;
            this.Field_StopVector_Frame3.HeaderText = "Stop Vector";
            this.Field_StopVector_Frame3.MaxInputLength = 5;
            this.Field_StopVector_Frame3.Name = "Field_StopVector_Frame3";
            this.Field_StopVector_Frame3.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StopVector_Frame3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StopVector_Frame3.Width = 110;
            // 
            // tabPage_Frame4
            // 
            this.tabPage_Frame4.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage_Frame4.Controls.Add(this.groupBox_FrameName_Frame4);
            this.tabPage_Frame4.Controls.Add(this.groupBox_PatternFile_Frame4);
            this.tabPage_Frame4.Controls.Add(this.groupBox_FieldInfo_Frame4);
            this.tabPage_Frame4.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Frame4.Name = "tabPage_Frame4";
            this.tabPage_Frame4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Frame4.Size = new System.Drawing.Size(745, 295);
            this.tabPage_Frame4.TabIndex = 4;
            this.tabPage_Frame4.Text = "Frame4";
            // 
            // groupBox_FrameName_Frame4
            // 
            this.groupBox_FrameName_Frame4.Controls.Add(this.textBox_FrameName_Frame4);
            this.groupBox_FrameName_Frame4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FrameName_Frame4.Location = new System.Drawing.Point(6, 64);
            this.groupBox_FrameName_Frame4.Name = "groupBox_FrameName_Frame4";
            this.groupBox_FrameName_Frame4.Size = new System.Drawing.Size(732, 52);
            this.groupBox_FrameName_Frame4.TabIndex = 47;
            this.groupBox_FrameName_Frame4.TabStop = false;
            this.groupBox_FrameName_Frame4.Text = "Frame Name";
            // 
            // textBox_FrameName_Frame4
            // 
            this.textBox_FrameName_Frame4.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_FrameName_Frame4.Location = new System.Drawing.Point(6, 19);
            this.textBox_FrameName_Frame4.Name = "textBox_FrameName_Frame4";
            this.textBox_FrameName_Frame4.Size = new System.Drawing.Size(713, 20);
            this.textBox_FrameName_Frame4.TabIndex = 48;
            // 
            // groupBox_PatternFile_Frame4
            // 
            this.groupBox_PatternFile_Frame4.Controls.Add(this.textBox_PatternFile_Frame4);
            this.groupBox_PatternFile_Frame4.Controls.Add(this.button_SelectPatternFile_Frame4);
            this.groupBox_PatternFile_Frame4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PatternFile_Frame4.Location = new System.Drawing.Point(6, 9);
            this.groupBox_PatternFile_Frame4.Name = "groupBox_PatternFile_Frame4";
            this.groupBox_PatternFile_Frame4.Size = new System.Drawing.Size(732, 52);
            this.groupBox_PatternFile_Frame4.TabIndex = 46;
            this.groupBox_PatternFile_Frame4.TabStop = false;
            this.groupBox_PatternFile_Frame4.Text = "Pattern File";
            // 
            // textBox_PatternFile_Frame4
            // 
            this.textBox_PatternFile_Frame4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_PatternFile_Frame4.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_PatternFile_Frame4.Location = new System.Drawing.Point(6, 20);
            this.textBox_PatternFile_Frame4.Name = "textBox_PatternFile_Frame4";
            this.textBox_PatternFile_Frame4.Size = new System.Drawing.Size(620, 20);
            this.textBox_PatternFile_Frame4.TabIndex = 41;
            // 
            // button_SelectPatternFile_Frame4
            // 
            this.button_SelectPatternFile_Frame4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectPatternFile_Frame4.Location = new System.Drawing.Point(632, 20);
            this.button_SelectPatternFile_Frame4.Name = "button_SelectPatternFile_Frame4";
            this.button_SelectPatternFile_Frame4.Size = new System.Drawing.Size(87, 23);
            this.button_SelectPatternFile_Frame4.TabIndex = 2;
            this.button_SelectPatternFile_Frame4.Text = "Select";
            this.button_SelectPatternFile_Frame4.UseVisualStyleBackColor = true;
            this.button_SelectPatternFile_Frame4.Click += new System.EventHandler(this.bt_SelectPatternFile_Click);
            // 
            // groupBox_FieldInfo_Frame4
            // 
            this.groupBox_FieldInfo_Frame4.Controls.Add(this.dataGridView_FieldInfo_Frame4);
            this.groupBox_FieldInfo_Frame4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FieldInfo_Frame4.Location = new System.Drawing.Point(6, 118);
            this.groupBox_FieldInfo_Frame4.Name = "groupBox_FieldInfo_Frame4";
            this.groupBox_FieldInfo_Frame4.Size = new System.Drawing.Size(732, 167);
            this.groupBox_FieldInfo_Frame4.TabIndex = 48;
            this.groupBox_FieldInfo_Frame4.TabStop = false;
            this.groupBox_FieldInfo_Frame4.Text = "Field Info";
            // 
            // dataGridView_FieldInfo_Frame4
            // 
            this.dataGridView_FieldInfo_Frame4.AllowUserToAddRows = false;
            this.dataGridView_FieldInfo_Frame4.AllowUserToResizeColumns = false;
            this.dataGridView_FieldInfo_Frame4.AllowUserToResizeRows = false;
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_FieldInfo_Frame4.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle13;
            this.dataGridView_FieldInfo_Frame4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Field_FieldName_Frame4,
            this.Field_PinName_Frame4,
            this.Field_Bits_Frame4,
            this.Field_StartVector_Frame4,
            this.Field_StopVector_Frame4});
            this.dataGridView_FieldInfo_Frame4.ContextMenuStrip = this.contextMenuStrip_FieldInfo;
            this.dataGridView_FieldInfo_Frame4.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_FieldInfo_Frame4.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_FieldInfo_Frame4.Location = new System.Drawing.Point(6, 17);
            this.dataGridView_FieldInfo_Frame4.MultiSelect = false;
            this.dataGridView_FieldInfo_Frame4.Name = "dataGridView_FieldInfo_Frame4";
            this.dataGridView_FieldInfo_Frame4.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame4.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_FieldInfo_Frame4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_FieldInfo_Frame4.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_FieldInfo_Frame4.Size = new System.Drawing.Size(713, 136);
            this.dataGridView_FieldInfo_Frame4.TabIndex = 48;
            // 
            // Field_FieldName_Frame4
            // 
            this.Field_FieldName_Frame4.Frozen = true;
            this.Field_FieldName_Frame4.HeaderText = "Field Name";
            this.Field_FieldName_Frame4.MinimumWidth = 200;
            this.Field_FieldName_Frame4.Name = "Field_FieldName_Frame4";
            this.Field_FieldName_Frame4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_FieldName_Frame4.Width = 200;
            // 
            // Field_PinName_Frame4
            // 
            this.Field_PinName_Frame4.Frozen = true;
            this.Field_PinName_Frame4.HeaderText = "Pin Name";
            this.Field_PinName_Frame4.MinimumWidth = 190;
            this.Field_PinName_Frame4.Name = "Field_PinName_Frame4";
            this.Field_PinName_Frame4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_PinName_Frame4.Width = 190;
            // 
            // Field_Bits_Frame4
            // 
            dataGridViewCellStyle14.Format = "N0";
            dataGridViewCellStyle14.NullValue = null;
            this.Field_Bits_Frame4.DefaultCellStyle = dataGridViewCellStyle14;
            this.Field_Bits_Frame4.Frozen = true;
            this.Field_Bits_Frame4.HeaderText = "Bits";
            this.Field_Bits_Frame4.MaxInputLength = 5;
            this.Field_Bits_Frame4.Name = "Field_Bits_Frame4";
            this.Field_Bits_Frame4.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_Bits_Frame4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_Bits_Frame4.Width = 60;
            // 
            // Field_StartVector_Frame4
            // 
            dataGridViewCellStyle15.Format = "N0";
            dataGridViewCellStyle15.NullValue = null;
            this.Field_StartVector_Frame4.DefaultCellStyle = dataGridViewCellStyle15;
            this.Field_StartVector_Frame4.Frozen = true;
            this.Field_StartVector_Frame4.HeaderText = "Start Vector";
            this.Field_StartVector_Frame4.MaxInputLength = 5;
            this.Field_StartVector_Frame4.Name = "Field_StartVector_Frame4";
            this.Field_StartVector_Frame4.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StartVector_Frame4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StartVector_Frame4.Width = 110;
            // 
            // Field_StopVector_Frame4
            // 
            dataGridViewCellStyle16.Format = "N0";
            dataGridViewCellStyle16.NullValue = null;
            this.Field_StopVector_Frame4.DefaultCellStyle = dataGridViewCellStyle16;
            this.Field_StopVector_Frame4.Frozen = true;
            this.Field_StopVector_Frame4.HeaderText = "Stop Vector";
            this.Field_StopVector_Frame4.MaxInputLength = 5;
            this.Field_StopVector_Frame4.Name = "Field_StopVector_Frame4";
            this.Field_StopVector_Frame4.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StopVector_Frame4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StopVector_Frame4.Width = 110;
            // 
            // tabPage_Frame5
            // 
            this.tabPage_Frame5.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage_Frame5.Controls.Add(this.groupBox_FrameName_Frame5);
            this.tabPage_Frame5.Controls.Add(this.groupBox_PatternFile_Frame5);
            this.tabPage_Frame5.Controls.Add(this.groupBox_FieldInfo_Frame5);
            this.tabPage_Frame5.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Frame5.Name = "tabPage_Frame5";
            this.tabPage_Frame5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Frame5.Size = new System.Drawing.Size(745, 295);
            this.tabPage_Frame5.TabIndex = 5;
            this.tabPage_Frame5.Text = "Frame5";
            // 
            // groupBox_FrameName_Frame5
            // 
            this.groupBox_FrameName_Frame5.Controls.Add(this.textBox_FrameName_Frame5);
            this.groupBox_FrameName_Frame5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FrameName_Frame5.Location = new System.Drawing.Point(6, 64);
            this.groupBox_FrameName_Frame5.Name = "groupBox_FrameName_Frame5";
            this.groupBox_FrameName_Frame5.Size = new System.Drawing.Size(732, 52);
            this.groupBox_FrameName_Frame5.TabIndex = 47;
            this.groupBox_FrameName_Frame5.TabStop = false;
            this.groupBox_FrameName_Frame5.Text = "Frame Name";
            // 
            // textBox_FrameName_Frame5
            // 
            this.textBox_FrameName_Frame5.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_FrameName_Frame5.Location = new System.Drawing.Point(6, 19);
            this.textBox_FrameName_Frame5.Name = "textBox_FrameName_Frame5";
            this.textBox_FrameName_Frame5.Size = new System.Drawing.Size(713, 20);
            this.textBox_FrameName_Frame5.TabIndex = 48;
            // 
            // groupBox_PatternFile_Frame5
            // 
            this.groupBox_PatternFile_Frame5.Controls.Add(this.textBox_PatternFile_Frame5);
            this.groupBox_PatternFile_Frame5.Controls.Add(this.button_SelectPatternFile_Frame5);
            this.groupBox_PatternFile_Frame5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PatternFile_Frame5.Location = new System.Drawing.Point(6, 9);
            this.groupBox_PatternFile_Frame5.Name = "groupBox_PatternFile_Frame5";
            this.groupBox_PatternFile_Frame5.Size = new System.Drawing.Size(732, 52);
            this.groupBox_PatternFile_Frame5.TabIndex = 46;
            this.groupBox_PatternFile_Frame5.TabStop = false;
            this.groupBox_PatternFile_Frame5.Text = "Pattern File";
            // 
            // textBox_PatternFile_Frame5
            // 
            this.textBox_PatternFile_Frame5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_PatternFile_Frame5.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_PatternFile_Frame5.Location = new System.Drawing.Point(6, 20);
            this.textBox_PatternFile_Frame5.Name = "textBox_PatternFile_Frame5";
            this.textBox_PatternFile_Frame5.Size = new System.Drawing.Size(620, 20);
            this.textBox_PatternFile_Frame5.TabIndex = 41;
            // 
            // button_SelectPatternFile_Frame5
            // 
            this.button_SelectPatternFile_Frame5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectPatternFile_Frame5.Location = new System.Drawing.Point(632, 20);
            this.button_SelectPatternFile_Frame5.Name = "button_SelectPatternFile_Frame5";
            this.button_SelectPatternFile_Frame5.Size = new System.Drawing.Size(87, 23);
            this.button_SelectPatternFile_Frame5.TabIndex = 2;
            this.button_SelectPatternFile_Frame5.Text = "Select";
            this.button_SelectPatternFile_Frame5.UseVisualStyleBackColor = true;
            this.button_SelectPatternFile_Frame5.Click += new System.EventHandler(this.bt_SelectPatternFile_Click);
            // 
            // groupBox_FieldInfo_Frame5
            // 
            this.groupBox_FieldInfo_Frame5.Controls.Add(this.dataGridView_FieldInfo_Frame5);
            this.groupBox_FieldInfo_Frame5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_FieldInfo_Frame5.Location = new System.Drawing.Point(6, 118);
            this.groupBox_FieldInfo_Frame5.Name = "groupBox_FieldInfo_Frame5";
            this.groupBox_FieldInfo_Frame5.Size = new System.Drawing.Size(732, 167);
            this.groupBox_FieldInfo_Frame5.TabIndex = 48;
            this.groupBox_FieldInfo_Frame5.TabStop = false;
            this.groupBox_FieldInfo_Frame5.Text = "Field Info";
            // 
            // dataGridView_FieldInfo_Frame5
            // 
            this.dataGridView_FieldInfo_Frame5.AllowUserToAddRows = false;
            this.dataGridView_FieldInfo_Frame5.AllowUserToResizeColumns = false;
            this.dataGridView_FieldInfo_Frame5.AllowUserToResizeRows = false;
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle17.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle17.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle17.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_FieldInfo_Frame5.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle17;
            this.dataGridView_FieldInfo_Frame5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame5.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Field_FieldName_Frame5,
            this.Field_PinName_Frame5,
            this.Field_Bits_Frame5,
            this.Field_StartVector_Frame5,
            this.Field_StopVector_Frame5});
            this.dataGridView_FieldInfo_Frame5.ContextMenuStrip = this.contextMenuStrip_FieldInfo;
            this.dataGridView_FieldInfo_Frame5.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_FieldInfo_Frame5.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_FieldInfo_Frame5.Location = new System.Drawing.Point(6, 17);
            this.dataGridView_FieldInfo_Frame5.MultiSelect = false;
            this.dataGridView_FieldInfo_Frame5.Name = "dataGridView_FieldInfo_Frame5";
            this.dataGridView_FieldInfo_Frame5.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_FieldInfo_Frame5.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_FieldInfo_Frame5.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_FieldInfo_Frame5.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_FieldInfo_Frame5.Size = new System.Drawing.Size(713, 136);
            this.dataGridView_FieldInfo_Frame5.TabIndex = 48;
            // 
            // Field_FieldName_Frame5
            // 
            this.Field_FieldName_Frame5.Frozen = true;
            this.Field_FieldName_Frame5.HeaderText = "Field Name";
            this.Field_FieldName_Frame5.MinimumWidth = 200;
            this.Field_FieldName_Frame5.Name = "Field_FieldName_Frame5";
            this.Field_FieldName_Frame5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_FieldName_Frame5.Width = 200;
            // 
            // Field_PinName_Frame5
            // 
            this.Field_PinName_Frame5.Frozen = true;
            this.Field_PinName_Frame5.HeaderText = "Pin Name";
            this.Field_PinName_Frame5.MinimumWidth = 190;
            this.Field_PinName_Frame5.Name = "Field_PinName_Frame5";
            this.Field_PinName_Frame5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_PinName_Frame5.Width = 190;
            // 
            // Field_Bits_Frame5
            // 
            dataGridViewCellStyle18.Format = "N0";
            dataGridViewCellStyle18.NullValue = null;
            this.Field_Bits_Frame5.DefaultCellStyle = dataGridViewCellStyle18;
            this.Field_Bits_Frame5.Frozen = true;
            this.Field_Bits_Frame5.HeaderText = "Bits";
            this.Field_Bits_Frame5.MaxInputLength = 5;
            this.Field_Bits_Frame5.Name = "Field_Bits_Frame5";
            this.Field_Bits_Frame5.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_Bits_Frame5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_Bits_Frame5.Width = 60;
            // 
            // Field_StartVector_Frame5
            // 
            dataGridViewCellStyle19.Format = "N0";
            dataGridViewCellStyle19.NullValue = null;
            this.Field_StartVector_Frame5.DefaultCellStyle = dataGridViewCellStyle19;
            this.Field_StartVector_Frame5.Frozen = true;
            this.Field_StartVector_Frame5.HeaderText = "Start Vector";
            this.Field_StartVector_Frame5.MaxInputLength = 5;
            this.Field_StartVector_Frame5.Name = "Field_StartVector_Frame5";
            this.Field_StartVector_Frame5.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StartVector_Frame5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StartVector_Frame5.Width = 110;
            // 
            // Field_StopVector_Frame5
            // 
            dataGridViewCellStyle20.Format = "N0";
            dataGridViewCellStyle20.NullValue = null;
            this.Field_StopVector_Frame5.DefaultCellStyle = dataGridViewCellStyle20;
            this.Field_StopVector_Frame5.Frozen = true;
            this.Field_StopVector_Frame5.HeaderText = "Stop Vector";
            this.Field_StopVector_Frame5.MaxInputLength = 5;
            this.Field_StopVector_Frame5.Name = "Field_StopVector_Frame5";
            this.Field_StopVector_Frame5.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.Field_StopVector_Frame5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Field_StopVector_Frame5.Width = 110;
            // 
            // groupBox_PortPinMappingInfo
            // 
            this.groupBox_PortPinMappingInfo.Controls.Add(this.comboBox_Protocol);
            this.groupBox_PortPinMappingInfo.Controls.Add(this.label_Protocol);
            this.groupBox_PortPinMappingInfo.Controls.Add(this.dataGridView_PortPinMappingInfo);
            this.groupBox_PortPinMappingInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_PortPinMappingInfo.Location = new System.Drawing.Point(12, 335);
            this.groupBox_PortPinMappingInfo.Name = "groupBox_PortPinMappingInfo";
            this.groupBox_PortPinMappingInfo.Size = new System.Drawing.Size(450, 197);
            this.groupBox_PortPinMappingInfo.TabIndex = 49;
            this.groupBox_PortPinMappingInfo.TabStop = false;
            this.groupBox_PortPinMappingInfo.Text = "Port -> Pin Mapping Info";
            // 
            // comboBox_Protocol
            // 
            this.comboBox_Protocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Protocol.Location = new System.Drawing.Point(80, 20);
            this.comboBox_Protocol.Name = "comboBox_Protocol";
            this.comboBox_Protocol.Size = new System.Drawing.Size(185, 21);
            this.comboBox_Protocol.TabIndex = 49;
            this.comboBox_Protocol.SelectedIndexChanged += new System.EventHandler(this.comboBox_Protocol_SelectedIndexChanged);
            // 
            // label_Protocol
            // 
            this.label_Protocol.AutoSize = true;
            this.label_Protocol.Location = new System.Drawing.Point(6, 23);
            this.label_Protocol.Name = "label_Protocol";
            this.label_Protocol.Size = new System.Drawing.Size(54, 13);
            this.label_Protocol.TabIndex = 46;
            this.label_Protocol.Text = "Protocol";
            // 
            // dataGridView_PortPinMappingInfo
            // 
            this.dataGridView_PortPinMappingInfo.AllowUserToAddRows = false;
            this.dataGridView_PortPinMappingInfo.AllowUserToResizeColumns = false;
            this.dataGridView_PortPinMappingInfo.AllowUserToResizeRows = false;
            dataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_PortPinMappingInfo.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle21;
            this.dataGridView_PortPinMappingInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView_PortPinMappingInfo.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Protocol_PortName,
            this.Type,
            this.Protocol_PinName});
            this.dataGridView_PortPinMappingInfo.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView_PortPinMappingInfo.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.dataGridView_PortPinMappingInfo.Location = new System.Drawing.Point(7, 51);
            this.dataGridView_PortPinMappingInfo.MultiSelect = false;
            this.dataGridView_PortPinMappingInfo.Name = "dataGridView_PortPinMappingInfo";
            this.dataGridView_PortPinMappingInfo.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_PortPinMappingInfo.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_PortPinMappingInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView_PortPinMappingInfo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_PortPinMappingInfo.Size = new System.Drawing.Size(434, 136);
            this.dataGridView_PortPinMappingInfo.TabIndex = 48;
            // 
            // Protocol_PortName
            // 
            this.Protocol_PortName.Frozen = true;
            this.Protocol_PortName.HeaderText = "Port Name";
            this.Protocol_PortName.MaxInputLength = 100;
            this.Protocol_PortName.MinimumWidth = 150;
            this.Protocol_PortName.Name = "Protocol_PortName";
            this.Protocol_PortName.ReadOnly = true;
            this.Protocol_PortName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Protocol_PortName.Width = 150;
            // 
            // Type
            // 
            this.Type.HeaderText = "Type";
            this.Type.Name = "Type";
            this.Type.Visible = false;
            // 
            // Protocol_PinName
            // 
            this.Protocol_PinName.HeaderText = "Pin Name";
            this.Protocol_PinName.MinimumWidth = 240;
            this.Protocol_PinName.Name = "Protocol_PinName";
            this.Protocol_PinName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Protocol_PinName.Width = 240;
            // 
            // groupBox_OutputDir
            // 
            this.groupBox_OutputDir.Controls.Add(this.textBox_OutputDir);
            this.groupBox_OutputDir.Controls.Add(this.button_SelectOutputDir);
            this.groupBox_OutputDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox_OutputDir.Location = new System.Drawing.Point(12, 538);
            this.groupBox_OutputDir.Name = "groupBox_OutputDir";
            this.groupBox_OutputDir.Size = new System.Drawing.Size(749, 52);
            this.groupBox_OutputDir.TabIndex = 50;
            this.groupBox_OutputDir.TabStop = false;
            this.groupBox_OutputDir.Text = "Output Dir";
            // 
            // textBox_OutputDir
            // 
            this.textBox_OutputDir.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.textBox_OutputDir.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.textBox_OutputDir.Location = new System.Drawing.Point(6, 20);
            this.textBox_OutputDir.Name = "textBox_OutputDir";
            this.textBox_OutputDir.Size = new System.Drawing.Size(637, 20);
            this.textBox_OutputDir.TabIndex = 41;
            // 
            // button_SelectOutputDir
            // 
            this.button_SelectOutputDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button_SelectOutputDir.Location = new System.Drawing.Point(649, 20);
            this.button_SelectOutputDir.Name = "button_SelectOutputDir";
            this.button_SelectOutputDir.Size = new System.Drawing.Size(87, 23);
            this.button_SelectOutputDir.TabIndex = 2;
            this.button_SelectOutputDir.Text = "Select";
            this.button_SelectOutputDir.UseVisualStyleBackColor = true;
            this.button_SelectOutputDir.Click += new System.EventHandler(this.button_SelectOutputDir_Click);
            // 
            // progressBarEx_Process
            // 
            this.progressBarEx_Process.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBarEx_Process.Location = new System.Drawing.Point(0, 597);
            this.progressBarEx_Process.Name = "progressBarEx_Process";
            this.progressBarEx_Process.Size = new System.Drawing.Size(774, 23);
            this.progressBarEx_Process.Stage = null;
            this.progressBarEx_Process.Step = 1;
            this.progressBarEx_Process.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBarEx_Process.TabIndex = 40;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 620);
            this.Controls.Add(this.groupBox_OutputDir);
            this.Controls.Add(this.groupBox_PortPinMappingInfo);
            this.Controls.Add(this.groupBox_TimeSetName);
            this.Controls.Add(this.bt_Generate);
            this.Controls.Add(this.tabControl_Frames);
            this.Controls.Add(this.progressBarEx_Process);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "nWireDefinitionGen";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tabPage_Frame1.ResumeLayout(false);
            this.groupBox_FrameName_Frame1.ResumeLayout(false);
            this.groupBox_FrameName_Frame1.PerformLayout();
            this.groupBox_PatternFile_Frame1.ResumeLayout(false);
            this.groupBox_PatternFile_Frame1.PerformLayout();
            this.groupBox_FieldInfo_Frame1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame1)).EndInit();
            this.contextMenuStrip_FieldInfo.ResumeLayout(false);
            this.groupBox_TimeSetName.ResumeLayout(false);
            this.groupBox_TimeSetName.PerformLayout();
            this.tabControl_Frames.ResumeLayout(false);
            this.contextMenuStrip_tabControl.ResumeLayout(false);
            this.tabPage_Frame2.ResumeLayout(false);
            this.groupBox_FrameName_Frame2.ResumeLayout(false);
            this.groupBox_FrameName_Frame2.PerformLayout();
            this.groupBox_PatternFile_Frame2.ResumeLayout(false);
            this.groupBox_PatternFile_Frame2.PerformLayout();
            this.groupBox_FieldInfo_Frame2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame2)).EndInit();
            this.tabPage_Frame3.ResumeLayout(false);
            this.groupBox_FrameName_Frame3.ResumeLayout(false);
            this.groupBox_FrameName_Frame3.PerformLayout();
            this.groupBox_PatternFile_Frame3.ResumeLayout(false);
            this.groupBox_PatternFile_Frame3.PerformLayout();
            this.groupBox_FieldInfo_Frame3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame3)).EndInit();
            this.tabPage_Frame4.ResumeLayout(false);
            this.groupBox_FrameName_Frame4.ResumeLayout(false);
            this.groupBox_FrameName_Frame4.PerformLayout();
            this.groupBox_PatternFile_Frame4.ResumeLayout(false);
            this.groupBox_PatternFile_Frame4.PerformLayout();
            this.groupBox_FieldInfo_Frame4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame4)).EndInit();
            this.tabPage_Frame5.ResumeLayout(false);
            this.groupBox_FrameName_Frame5.ResumeLayout(false);
            this.groupBox_FrameName_Frame5.PerformLayout();
            this.groupBox_PatternFile_Frame5.ResumeLayout(false);
            this.groupBox_PatternFile_Frame5.PerformLayout();
            this.groupBox_FieldInfo_Frame5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_FieldInfo_Frame5)).EndInit();
            this.groupBox_PortPinMappingInfo.ResumeLayout(false);
            this.groupBox_PortPinMappingInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_PortPinMappingInfo)).EndInit();
            this.groupBox_OutputDir.ResumeLayout(false);
            this.groupBox_OutputDir.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog_Input;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog_Output;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private ProgressBarEx progressBarEx_Process;
        private System.Windows.Forms.TabPage tabPage_Frame1;
        private System.Windows.Forms.GroupBox groupBox_PatternFile_Frame1;
        private System.Windows.Forms.TextBox textBox_PatternFile_Frame1;
        private System.Windows.Forms.Button button_SelectPatternFile_Frame1;
        private System.Windows.Forms.Button bt_Generate;
        private System.Windows.Forms.GroupBox groupBox_FieldInfo_Frame1;
        private System.Windows.Forms.TabControl tabControl_Frames;
        private System.Windows.Forms.GroupBox groupBox_TimeSetName;
        private System.Windows.Forms.TextBox textBox_TimeSetName;
        private System.Windows.Forms.DataGridView dataGridView_FieldInfo_Frame1;
        private System.Windows.Forms.GroupBox groupBox_FrameName_Frame1;
        private System.Windows.Forms.TextBox textBox_FrameName_Frame1;
        private System.Windows.Forms.GroupBox groupBox_PortPinMappingInfo;
        private System.Windows.Forms.DataGridView dataGridView_PortPinMappingInfo;
        private System.Windows.Forms.Label label_Protocol;
        private System.Windows.Forms.GroupBox groupBox_OutputDir;
        private System.Windows.Forms.TextBox textBox_OutputDir;
        private System.Windows.Forms.Button button_SelectOutputDir;
        private System.Windows.Forms.TabPage tabPage_Frame2;
        private System.Windows.Forms.GroupBox groupBox_FrameName_Frame2;
        private System.Windows.Forms.TextBox textBox_FrameName_Frame2;
        private System.Windows.Forms.GroupBox groupBox_PatternFile_Frame2;
        private System.Windows.Forms.TextBox textBox_PatternFile_Frame2;
        private System.Windows.Forms.Button button_SelectPatternFile_Frame2;
        private System.Windows.Forms.GroupBox groupBox_FieldInfo_Frame2;
        private System.Windows.Forms.DataGridView dataGridView_FieldInfo_Frame2;
        private System.Windows.Forms.TabPage tabPage_Frame3;
        private System.Windows.Forms.TabPage tabPage_Frame4;
        private System.Windows.Forms.TabPage tabPage_Frame5;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip_tabControl;
        private System.Windows.Forms.ToolStripMenuItem ItemAddFrame;
        private System.Windows.Forms.ToolStripMenuItem ItemDeleteFrame;
        private System.Windows.Forms.GroupBox groupBox_FrameName_Frame3;
        private System.Windows.Forms.TextBox textBox_FrameName_Frame3;
        private System.Windows.Forms.GroupBox groupBox_PatternFile_Frame3;
        private System.Windows.Forms.TextBox textBox_PatternFile_Frame3;
        private System.Windows.Forms.Button button_SelectPatternFile_Frame3;
        private System.Windows.Forms.GroupBox groupBox_FieldInfo_Frame3;
        private System.Windows.Forms.DataGridView dataGridView_FieldInfo_Frame3;
        private System.Windows.Forms.GroupBox groupBox_FrameName_Frame4;
        private System.Windows.Forms.TextBox textBox_FrameName_Frame4;
        private System.Windows.Forms.GroupBox groupBox_PatternFile_Frame4;
        private System.Windows.Forms.TextBox textBox_PatternFile_Frame4;
        private System.Windows.Forms.Button button_SelectPatternFile_Frame4;
        private System.Windows.Forms.GroupBox groupBox_FieldInfo_Frame4;
        private System.Windows.Forms.DataGridView dataGridView_FieldInfo_Frame4;
        private System.Windows.Forms.GroupBox groupBox_FrameName_Frame5;
        private System.Windows.Forms.TextBox textBox_FrameName_Frame5;
        private System.Windows.Forms.GroupBox groupBox_PatternFile_Frame5;
        private System.Windows.Forms.TextBox textBox_PatternFile_Frame5;
        private System.Windows.Forms.Button button_SelectPatternFile_Frame5;
        private System.Windows.Forms.GroupBox groupBox_FieldInfo_Frame5;
        private System.Windows.Forms.DataGridView dataGridView_FieldInfo_Frame5;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip_FieldInfo;
        private System.Windows.Forms.ToolStripMenuItem ItemAddField;
        private System.Windows.Forms.ToolStripMenuItem ItemDeleteField;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_FieldName_Frame1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_PinName_Frame1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_Bits_Frame1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StartVector_Frame1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StopVector_Frame1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_FieldName_Frame2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_PinName_Frame2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_Bits_Frame2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StartVector_Frame2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StopVector_Frame2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_FieldName_Frame3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_PinName_Frame3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_Bits_Frame3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StartVector_Frame3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StopVector_Frame3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_FieldName_Frame4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_PinName_Frame4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_Bits_Frame4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StartVector_Frame4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StopVector_Frame4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_FieldName_Frame5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_PinName_Frame5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_Bits_Frame5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StartVector_Frame5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Field_StopVector_Frame5;
        private System.Windows.Forms.ComboBox comboBox_Protocol;
        private System.Windows.Forms.DataGridViewTextBoxColumn Protocol_PortName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Protocol_PinName;
    }
}

