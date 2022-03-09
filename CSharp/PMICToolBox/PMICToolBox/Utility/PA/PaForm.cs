using PmicAutomation.MyControls;
using PmicAutomation.Utility.PA.Input;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using IgxlData.IgxlBase;
using System.Reflection;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace PmicAutomation.Utility.PA
{
    public partial class PaParser : MyForm
    {
        private UflexConfig _uflexConfig;
        public string IGXL_Version;

        public PaParser()
        {
            InitializeComponent();

            HelpButtonClicked += PaParser_HelpButtonClicked;

            comboBox_Device.SelectedIndex = 0;

            LoadIgxlVersion();
            IGXL_Version = comboBox_igxlversion.SelectedItem.ToString();

            UpdateTesterConfig();
        }

        private void Btn_PA_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.PaFile, true) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_TesterConfig_Click(object sender, EventArgs e)
        {
            if (FileDialog(sender, FileFilter.XmlFile) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_OutputPath_Click(object sender, EventArgs e)
        {
            if (PathDialog(sender, true) == null)
            {
                return;
            }

            CheckStatus();
        }

        private void Btn_DGSReference_Click(object sender, EventArgs e)
        {
            if(FileDialog(sender,FileFilter.Excel)==null)
            {
                return;
            }
        }

        private void ComboBox_Device_TextChanged(object sender, EventArgs e)
        {            
            CheckStatus();
        }

        private void Btn_Run_Click(object sender, EventArgs e)
        {
            try
            {
                richTextBox.Clear();

                new PaMain(this, _uflexConfig).WorkFlow();
            }
            catch (Exception exception)
            {
                AppendText("The exception was found !!!", Color.Red);
                AppendText(exception.ToString(), Color.Red);
            }
        }

        private void CheckStatus()
        {
            richTextBox.Clear();
            UpdateTesterConfig();

            if (string.IsNullOrEmpty(FileOpen_OutputPath.ButtonTextBox.Text))
            {
                FileOpen_OutputPath.ButtonTextBox.Text = DefaultPath;
            }

            if (string.IsNullOrEmpty(comboBox_Device.Text))
            {
                return;
            }

            if (string.IsNullOrEmpty(FileOpen_PA.ButtonTextBox.Text))
            {
                return;
            }

            if (!Directory.Exists(FileOpen_OutputPath.ButtonTextBox.Text))
            {
                return;
            }

            if (!File.Exists(FileOpen_TesterConfig.ButtonTextBox.Text) ||
                string.IsNullOrEmpty(FileOpen_TesterConfig.ButtonTextBox.Text))
            {
                return;
            }

            Btn_RunDownload.Run.Enabled = true;
        }

        private void UpdateTesterConfig()
        {
            try
            {
                string file = DefaultPath + "\\TesterConfig_" + comboBox_Device.Text + ".xml";

                if (!File.Exists(file))
                {
                    file = DefaultPath + "\\TesterConfig_Default.xml";
                }

                if (!File.Exists(file))
                {
                    file = Directory.GetCurrentDirectory() + "\\Config\\Tester\\" + "TesterConfig_" + comboBox_Device.Text +
                           ".xml";
                }

                if (!File.Exists(file))
                {
                    file = Directory.GetCurrentDirectory() + "\\Config\\Tester\\" + "TesterConfig_Default.xml";
                }

                if (!File.Exists(file))
                {
                    file = Directory.GetCurrentDirectory() + "\\TesterConfig_" + comboBox_Device.Text + ".xml";
                }

                if (!File.Exists(file))
                {
                    file = Directory.GetCurrentDirectory() + "\\TesterConfig_Default.xml";
                }

                if (File.Exists(file))
                {
                    FileOpen_TesterConfig.ButtonTextBox.Text = file;
                    _uflexConfig = UflexConfigReader.GetXml(FileOpen_TesterConfig.ButtonTextBox.Text);
                    textBox_HexVs.Text = _uflexConfig.HexVS;
                }
                else
                {
                    FileOpen_TesterConfig.ButtonTextBox.Text = "";
                }
            }catch(Exception ex)
            {
                AppendText("Update test config error!!", Color.Red);
                AppendText(ex.ToString(), Color.Red);
            }
        }

        public void AppendText(string text, Color color)
        {
            richTextBox.SelectionColor = color;
            richTextBox.AppendText(text + Environment.NewLine);
            richTextBox.ScrollToCaret();
            richTextBox.Refresh();
        }

        private void Btn_Download_Click(object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".Template.").Show();
        }

        private void PaParser_HelpButtonClicked(Object sender, EventArgs e)
        {
            int cnt = GetType().ToString().Split('.').Length;
            var resourceName = string.Join(".", GetType().ToString().Split('.').ToList().GetRange(0, cnt - 1));
            new MyDownloadForm().Download(resourceName + ".UserManual.").Show();
        }

        private void LoadIgxlVersion()
        {
            var assembly = Assembly.GetAssembly(typeof(IgxlData.IgxlSheets.IgxlSheet));
            var resourceNames = assembly.GetManifestResourceNames();
            List<string> igxlVersion = new List<string>();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.Contains(".IGXLSheetsVersion."))
                {
                    var xs = new XmlSerializer(typeof(IGXLVersion));
                    var igxlConfig = (IGXLVersion)xs.Deserialize(assembly.GetManifestResourceStream(resourceName));
                    igxlVersion.Add(igxlConfig.igxlVersion);
                }
            }

            comboBox_igxlversion.Items.AddRange(igxlVersion.ToArray());
            string igxlroot = Environment.GetEnvironmentVariable("IGXLROOT");
            if(string.IsNullOrEmpty(igxlroot))
            {
                SelectDefaultIgxlVersion();
                //MessageBox.Show("No installed IGXL found, output IGXL sheet version will be " + comboBox_igxlversion.SelectedItem, "No IGXL found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                string path = Path.Combine(igxlroot, "bin", "Version.Txt");
                if(File.Exists(path))
                {
                    string version = GetInstalledIGXLVersion(path);
                    string[] items = version.Split('.');
                    foreach(object v in comboBox_igxlversion.Items)
                    {
                        string[] vItems = v.ToString().Split('.');
                        if(items[0]==vItems[0]&&items[1]==vItems[1])
                        {
                            comboBox_igxlversion.SelectedItem = v;
                            return;
                        }
                    }

                    SelectDefaultIgxlVersion();
                    //MessageBox.Show("Current installed IGXL " + version + " is not supported, output IGXL sheet version will be " + comboBox_igxlversion.SelectedItem, "No supported IGXL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    SelectDefaultIgxlVersion();
                    //MessageBox.Show("Unrecognized IGXL version, output IGXL sheet version will be " + comboBox_igxlversion.SelectedItem, "Unrecognized IGXL version", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void SelectDefaultIgxlVersion()
        {
            foreach(object version in comboBox_igxlversion.Items)
            {
                if(version.ToString().StartsWith("9.00"))
                {
                    comboBox_igxlversion.SelectedItem = version;
                    return;
                }
            }

            comboBox_igxlversion.SelectedIndex = 0;
        }

        private string GetInstalledIGXLVersion(string path)
        {
            StreamReader reader = new StreamReader(path);
            while(reader.Peek()!=-1)
            {
                string line = reader.ReadLine();
                if(line.Trim().StartsWith("Version:"))
                {
                    string[] items = line.Split(':');
                    return items[1].Trim();
                }
            }

            return string.Empty;
        }

    }
}