using FWFrame;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace PmicAutomation.Utility.nWire
{
    public partial class MainForm : Form
    {
        private GUIInfo guiInfo = null;
        private Dictionary<string, List<Tuple<string, string>>> protocalInfo = null;

        private int tabPageCount = 1;
        private List<TabPage> tabPages = new List<TabPage>();
        private List<DataGridView> dataGridViewFieldInfos = new List<DataGridView>();
        private List<TextBox> textBoxPatternFiles = new List<TextBox>();
        private List<Button> textBoxSelectPatternFiles = new List<Button>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // tabPages
            tabPages.Add(tabPage_Frame1);
            tabPages.Add(tabPage_Frame2);
            tabPages.Add(tabPage_Frame3);
            tabPages.Add(tabPage_Frame4);
            tabPages.Add(tabPage_Frame5);

            // dataGridViewFieldInfos
            dataGridViewFieldInfos.Add(dataGridView_FieldInfo_Frame1);
            dataGridViewFieldInfos.Add(dataGridView_FieldInfo_Frame2);
            dataGridViewFieldInfos.Add(dataGridView_FieldInfo_Frame3);
            dataGridViewFieldInfos.Add(dataGridView_FieldInfo_Frame4);
            dataGridViewFieldInfos.Add(dataGridView_FieldInfo_Frame5);

            // textBoxPatternFiles
            textBoxPatternFiles.Add(textBox_PatternFile_Frame1);
            textBoxPatternFiles.Add(textBox_PatternFile_Frame2);
            textBoxPatternFiles.Add(textBox_PatternFile_Frame3);
            textBoxPatternFiles.Add(textBox_PatternFile_Frame4);
            textBoxPatternFiles.Add(textBox_PatternFile_Frame5);

            // textBoxSelectPatternFiles
            textBoxSelectPatternFiles.Add(button_SelectPatternFile_Frame1);
            textBoxSelectPatternFiles.Add(button_SelectPatternFile_Frame2);
            textBoxSelectPatternFiles.Add(button_SelectPatternFile_Frame3);
            textBoxSelectPatternFiles.Add(button_SelectPatternFile_Frame4);
            textBoxSelectPatternFiles.Add(button_SelectPatternFile_Frame5);

            // Form Text
            Text = Text + " " + Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // Frames
            for (int i = tabPageCount; i < tabPages.Count; i++)
            {
                tabControl_Frames.TabPages.Remove(tabPages[i]);
            }

            // Retrieve Protocol Info
            RetrieveProtocolInfo();
        }

        private void RetrieveProtocolInfo()
        {
            // Collect info
            guiInfo = GetGuiInfoTemplate();
            guiInfo.Command = "RETRIEVE_PROTOCOL_INFO";

            // Start Process Logic
            backgroundWorker.RunWorkerAsync(guiInfo);
        }

        private void bt_SelectPatternFile_Click(object sender, EventArgs e)
        {
            // Show dialog for file select
            openFileDialog_Input.Filter = "PatternFile |*.PAT";
            openFileDialog_Input.Title = "Select Pattern File";
            openFileDialog_Input.FileName = string.Empty;
            if (openFileDialog_Input.ShowDialog() == DialogResult.OK)
            {
                Button currentButton = (Button)sender;
                int index = textBoxSelectPatternFiles.IndexOf(currentButton);

                textBoxPatternFiles[index].Text = openFileDialog_Input.FileName;
            }
        }

        private T GetChildControl<T>(Control container, string baseName) where T : Control
        {
            T result = null;

            int maxIndex = tabPages.Count;

            for (int i = 0; i < tabPages.Count; i++)
            {
                string completeName = string.Concat(baseName, i + 1);
                Control[] controls = container.Controls.Find(completeName, true);
                if (controls.Length > 0)
                {
                    result = (T)controls[0];
                    break;
                }
            }

            if (result != null)
            {
                return result;
            }
            else
            {
                throw new FWFrameException(string.Format("Can not find Control with name like [{0}] which should be contained by [{1}]", baseName, container.Name));
            }
        }

        private string GetStringValue(DataGridViewRow row, int index)
        {
            string result = string.Empty;

            object value = row.Cells[index].Value;
            if (value != null)
            {
                result = value.ToString().Trim();
            }

            return result;
        }

        private void button_SelectOutputDir_Click(object sender, EventArgs e)
        {
            // Show dialog for folder select
            folderBrowserDialog_Output.Description = "Select Output Dir";
            if (folderBrowserDialog_Output.ShowDialog() == DialogResult.OK)
            {
                textBox_OutputDir.Text = folderBrowserDialog_Output.SelectedPath;
            }
        }

        private void bt_Generate_Click(object sender, EventArgs e)
        {
            // Collect info
            guiInfo = GetGuiInfoTemplate();
            guiInfo.Command = "GENERATE_PROTOCOL_DEFINITION";
            guiInfo.AddParameter("reportStatus", (Action<int, string>)ChangeReportProgress);

            //////////////////////////////////
            //
            // Check & Collect Common Info
            //
            //////////////////////////////////

            // 检查：Protocol的信息是否完备
            Dictionary<string, Tuple<string, string>> pinMappingInfo = new Dictionary<string, Tuple<string, string>>();
            List<string> protocalPinNames = new List<string>();
            for (int i = 0; i < dataGridView_PortPinMappingInfo.Rows.Count; i++)
            {
                DataGridViewRow row = dataGridView_PortPinMappingInfo.Rows[i];

                // Pin Name
                string protocalPinName = GetStringValue(row, 2);
                if (string.IsNullOrWhiteSpace(protocalPinName))
                {
                    MessageBox.Show(string.Format("Port -> Pin Mapping Info's Pin Name of Row[{0}] is not assigned", i + 1), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (protocalPinNames.Contains(protocalPinName))
                {
                    MessageBox.Show(string.Format("Assigned Pin Name [{0}] for Port -> Pin Mapping Info is duplicate", protocalPinName), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    protocalPinNames.Add(protocalPinName);
                }

                pinMappingInfo[GetStringValue(row, 0)] = new Tuple<string, string>(protocalPinName, GetStringValue(row, 1));
            }
            guiInfo.AddParameter("pinMappingInfo", pinMappingInfo);
            guiInfo.AddParameter("protocalName", comboBox_Protocol.SelectedItem.ToString());

            // 检查：OutputDir是否指定
            // 检查：OutputDir是否存在
            string outputDir = textBox_OutputDir.Text;
            if (string.IsNullOrWhiteSpace(outputDir))
            {
                MessageBox.Show(string.Format("Output Dir is not assigned"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!Directory.Exists(outputDir))
            {
                MessageBox.Show(string.Format("Can not find assigned Output Dir"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            guiInfo.AddParameter("outputDir", outputDir);

            // 检查：TimeSetName是否指定
            string timeSetName = textBox_TimeSetName.Text;
            if (string.IsNullOrWhiteSpace(timeSetName))
            {
                MessageBox.Show(string.Format("TimeSet Name is not assigned"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            guiInfo.AddParameter("timeSetName", timeSetName);


            //////////////////////////////////
            //
            // Check & Collect Frames' Info
            //
            //////////////////////////////////

            List<string> patternFiles = new List<string>();
            List<string> frameNames = new List<string>();
            Dictionary<string, List<List<string>>> fieldInfoForAllFrames = new Dictionary<string, List<List<string>>>();
            TabControl.TabPageCollection tabPages = tabControl_Frames.TabPages;
            foreach (TabPage tabPage in tabPages)
            {
                // 检查：PatternFile是否指定
                // 检查：PatternFile是否存在
                // 检查：PatternFile是否为Pat文件
                // 检查：PatternFile是否不重复
                string patternFile = GetChildControl<TextBox>(tabPage, @"textBox_PatternFile_Frame").Text;
                if (string.IsNullOrWhiteSpace(patternFile))
                {
                    MessageBox.Show(string.Format("Pattern File for [{0}] is not assigned", patternFile, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!File.Exists(patternFile))
                {
                    MessageBox.Show(string.Format("Can not find assigned Pattern File [{0}] for Frame [{1}]", patternFile, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (!Path.GetExtension(patternFile).Equals(".pat", StringComparison.CurrentCultureIgnoreCase))
                {
                    MessageBox.Show(string.Format("Assigned Pattern File [{0}] for Frame [{1}] is not a Pattern File", patternFile, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (patternFiles.Contains(patternFile))
                {
                    MessageBox.Show(string.Format("Assigned Pattern File [{0}] for {1} is duplicate with other Frames", patternFile, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    patternFiles.Add(patternFile);
                }

                // 检查：FrameName是否不为空
                // 检查：FrameName是否不重复
                string frameName = GetChildControl<TextBox>(tabPage, @"textBox_FrameName_Frame").Text;
                if (string.IsNullOrWhiteSpace(frameName))
                {
                    MessageBox.Show(string.Format("Frame Name for {0} is not assigned", tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (frameNames.Contains(frameName))
                {
                    MessageBox.Show(string.Format("Assigned Frame Name [{0}] for {1} is duplicate with other Frames", frameName, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    frameNames.Add(frameName);
                }

                // 检查：Field的信息是否完备
                // 检查：Field的FieldName是否不重复
                // 检查：Field的PinName的值是否存在于Protocol.PinName集合中
                List<List<string>> fieldInfoForSingleFrame = new List<List<string>>();
                DataGridView dataGridViewFieldInfo = GetChildControl<DataGridView>(tabPage, @"dataGridView_FieldInfo_Frame");
                List<string> fieldNames = new List<string>();
                for (int i = 0; i < dataGridViewFieldInfo.Rows.Count; i++)
                {
                    List<string> fieldInfo = new List<string>();

                    DataGridViewRow row = dataGridViewFieldInfo.Rows[i];

                    // Field Name
                    string fieldName = GetStringValue(row, 0);
                    if (string.IsNullOrWhiteSpace(fieldName))
                    {
                        MessageBox.Show(string.Format("Field Info's Field Name of Row[{0}] for {1} is not assigned", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (fieldNames.Contains(fieldName))
                    {
                        MessageBox.Show(string.Format("Assigned Field Name [{0}] for {1} is duplicate", fieldName, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        fieldNames.Add(fieldName);
                    }
                    fieldInfo.Add(fieldName);

                    // Pin Name
                    string pinName = GetStringValue(row, 1);
                    if (string.IsNullOrWhiteSpace(pinName))
                    {
                        MessageBox.Show(string.Format("Field Info's Pin Name of Row[{0}] for {1} is not assigned", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (!protocalPinNames.Any(x => x.Equals(pinName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        MessageBox.Show(string.Format("Field Info's Pin Name of Row[{0}] for {1} is not contained by Port -> Pin Mapping Info", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    fieldInfo.Add(pinName);

                    // Bits
                    string bits = GetStringValue(row, 2);
                    int intBit;
                    if (string.IsNullOrWhiteSpace(bits))
                    {
                        MessageBox.Show(string.Format("Field Info's Bits of Row[{0}] for {1} is not assigned", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (!int.TryParse(bits, out intBit))
                    {
                        MessageBox.Show(string.Format("Field Info's Bits of Row[{0}] for {1} must be Integer", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (intBit < 1)
                    {
                        MessageBox.Show(string.Format("Field Info's Bits of Row[{0}] for {1} must be not less than {2}", i + 1, tabPage.Text, 1), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    fieldInfo.Add(bits);

                    // Start Vector
                    string startVector = GetStringValue(row, 3);
                    int intStartVector;
                    if (string.IsNullOrWhiteSpace(startVector))
                    {
                        MessageBox.Show(string.Format("Field Info's Start Vector of Row[{0}] for {1} is not assigned", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (!int.TryParse(startVector, out intStartVector))
                    {
                        MessageBox.Show(string.Format("Field Info's Start Vector of Row[{0}] for {1} must be Integer", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (intStartVector < 0)
                    {
                        MessageBox.Show(string.Format("Field Info's Start Vector of Row[{0}] for {1} must be not less than {2}", i + 1, tabPage.Text, 0), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    fieldInfo.Add(startVector);

                    // Stop Vector
                    string stopVector = GetStringValue(row, 4);
                    int intStopVector;
                    if (string.IsNullOrWhiteSpace(stopVector))
                    {
                        MessageBox.Show(string.Format("Field Info's Stop Vector of Row[{0}] for {1} is not assigned", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (!int.TryParse(stopVector, out intStopVector))
                    {
                        MessageBox.Show(string.Format("Field Info's Stop Vector of Row[{0}] for {1} must be Integer", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (intStopVector < intStartVector)
                    {
                        MessageBox.Show(string.Format("Field Info's Stop Vector of Row[{0}] for {1} must be not less than Start Vector", i + 1, tabPage.Text), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    fieldInfo.Add(stopVector);

                    fieldInfoForSingleFrame.Add(fieldInfo);
                }

                fieldInfoForAllFrames[frameName] = fieldInfoForSingleFrame;
            }
            guiInfo.AddParameter("patternFiles", patternFiles);
            guiInfo.AddParameter("frameNames", frameNames);
            guiInfo.AddParameter("fieldInfoForAllFrames", fieldInfoForAllFrames);

            // Start Process Logic
            backgroundWorker.RunWorkerAsync(guiInfo);
        }

        private void contextMenuStrip_tabControl_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string clickedItemName = e.ClickedItem.Name;

            switch (clickedItemName)
            {
                case "ItemAddFrame":
                    AddOneFrameTabPage();
                    break;
                case "ItemDeleteFrame":
                    RemoveOneFramTabPage();
                    break;
                default:
                    break;
            }
        }

        private void AddOneFrameTabPage()
        {
            // Max 5
            if (tabPageCount >= 5)
            {
                MessageBox.Show("At most 5 Frames are supported!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            tabControl_Frames.TabPages.Add(tabPages[tabPageCount]);
            tabControl_Frames.SelectedTab = tabPages[tabPageCount];
            tabPageCount++;
        }

        private void RemoveOneFramTabPage()
        {
            // Min 1
            if (tabPageCount <= 1)
            {
                MessageBox.Show("At least 1 Frame is needed!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            tabPageCount--;
            tabControl_Frames.TabPages.Remove(tabPages[tabPageCount]);
            tabControl_Frames.SelectedTab = tabPages[tabPageCount - 1];
        }

        private void contextMenuStrip_FieldInfo_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            contextMenuStrip_FieldInfo.Hide();

            int selectedFrameIndex = tabControl_Frames.SelectedIndex;

            DataGridView dataGridView_FieldInfo = dataGridViewFieldInfos[selectedFrameIndex];

            string clickedItemName = e.ClickedItem.Name;

            switch (clickedItemName)
            {
                case "ItemAddField":
                    dataGridView_FieldInfo.Rows.Add();
                    break;
                case "ItemDeleteField":
                    if (dataGridView_FieldInfo.Rows.Count > 1)
                    {
                        dataGridView_FieldInfo.Rows.Remove(dataGridView_FieldInfo.SelectedRows[0]);
                    }
                    else
                    {
                        dataGridView_FieldInfo.Rows.Clear();
                    }
                    break;
                default:
                    break;
            }
        }

        private void comboBox_Protocol_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;

            dataGridView_PortPinMappingInfo.Rows.Clear();
            foreach (var port in protocalInfo[cb.SelectedItem.ToString()])
            {
                dataGridView_PortPinMappingInfo.Rows.Add(port.Item1, port.Item2);
            }
        }

        private GUIInfo GetGuiInfoTemplate()
        {
            GUIInfo guiInfo = new GUIInfo();

            guiInfo.ChipType = "nWireDefinition";
            guiInfo.WorkFolder = @".";
            guiInfo.AssemblyFile = "nWireDefinition.dll";

            return guiInfo;
        }

        #region BackgroudWorker

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            SetFormStatus(false);

            Dispatcher dp = new Dispatcher(guiInfo);
            string dispatherConfigureFilePath = Path.Combine(guiInfo.WorkFolder, "Configure", "DispatherConfigure.xml");
            string configFilePath = Path.Combine(guiInfo.WorkFolder, "Configure", "ProtocolConfigure.xml");
            dp.Dispatch(GetResources(dispatherConfigureFilePath), GetResources(configFilePath));
        }

        public StreamReader GetResources(string path)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                string name = Path.GetFileName(path);
                if (resourceName.Contains(name))
                {
                    var resource = assembly.GetManifestResourceStream(resourceName);
                    return new StreamReader(resource);
                }
            }
            return null;
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SetProgressBarStatus(e.ProgressPercentage, (string)e.UserState);
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Wait for progressBar refresh complete
            Thread.Sleep(TimeSpan.FromSeconds(1));

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                progressBarEx_Process.Value = 0;
            }
            else
            {
                switch (guiInfo.Command)
                {
                    case "RETRIEVE_PROTOCOL_INFO":
                        protocalInfo = guiInfo.GetParameter<Dictionary<string, List<Tuple<string, string>>>>("protocalInfo");
                        foreach (string protocalName in protocalInfo.Keys)
                        {
                            comboBox_Protocol.Items.Add(protocalName);
                        }
                        comboBox_Protocol.SelectedIndex = 0;
                        break;
                    case "GENERATE_PROTOCOL_DEFINITION":
                        string outputDir = Path.Combine(guiInfo.GetParameter<string>("outputDir"), "Temp");
                        Directory.Delete(outputDir, true);
                        break;
                    default:
                        break;
                }
            }

            SetFormStatus(true);
        }

        private void SetFormStatus(bool enabled)
        {
            Invoke(new Action(() => Enabled = enabled));
        }

        /// <summary>
        /// Change progress report
        /// </summary>
        /// <param name="phase"></param>
        /// <param name="message"></param>
        private void ChangeReportProgress(int percentage, string message)
        {
            backgroundWorker.ReportProgress(percentage, message);
        }

        /// <summary>
        /// Set ProgressBar's Status
        /// </summary>
        /// <param name="percentage"></param>
        /// <param name="message"></param>
        private void SetProgressBarStatus(int percentage, string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                progressBarEx_Process.Stage = message;
            }
            progressBarEx_Process.Value = percentage;
        }

        #endregion
    }
}
