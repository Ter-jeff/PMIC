using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace PmicAutogen.Config.ProjectConfig
{
    public partial class ProjectConfigSetting
    {
        public ProjectConfigSetting(string projectName = "")
        {
            ProjectName = projectName;
            InitializeComponent();
            _controlList = InitializeControls();
            CreateControls(_controlList);
            DeviceType = GetDeviceType();
            RemoveTabByDeviceName(DeviceType);
        }

        private void Changed(object sender, SelectionChangedEventArgs e)
        {
            var item = (ComboBox) sender;
            if (item != null)
                RemoveTabByDeviceName(item.SelectedValue.ToString());
        }

        private void RemoveTabByDeviceName(string deviceType)
        {
            if (!string.IsNullOrEmpty(deviceType))
                if (_projectConfigSetting.Exists(x =>
                    x.GroupName.Equals("Device", StringComparison.OrdinalIgnoreCase) &&
                    x.Name.Equals("Type", StringComparison.OrdinalIgnoreCase)))
                {
                    var row = _projectConfigSetting.Find(x =>
                        x.GroupName.Equals("Device", StringComparison.OrdinalIgnoreCase) &&
                        x.Name.Equals("Type", StringComparison.OrdinalIgnoreCase));
                    var list = row.OptionValue.Where(x => !string.IsNullOrEmpty(x)).ToList();
                    list.Remove(deviceType);
                    foreach (var tabControl in GridMain.Children)
                    {
                        var control = (TabControl) tabControl;
                        foreach (var item in control.Items)
                        {
                            var tab = (TabItem) item;
                            tab.IsEnabled = true;
                            foreach (var name in list)
                                if (tab.Name.Equals(name))
                                    tab.IsEnabled = false;
                        }
                    }
                }
        }

        private string GetDeviceType()
        {
            if (_projectConfig.Exists(x =>
                x.GroupName.Equals("Device", StringComparison.OrdinalIgnoreCase) &&
                x.Name.Equals("Type", StringComparison.OrdinalIgnoreCase)))
                return _projectConfig.Find(x =>
                    x.GroupName.Equals("Device", StringComparison.OrdinalIgnoreCase) &&
                    x.Name.Equals("Type", StringComparison.OrdinalIgnoreCase)).Value;
            return "";
        }

        private List<ProjectConfigControlGroup> InitializeControls()
        {
            _projectConfigSetting = ProjectConfigSingleton.Instance().GetProjectConfigSetting();
            _projectConfig = ProjectConfigSingleton.Instance().GetProjectConfigRow();

            var projectConfigControlGroups = new List<ProjectConfigControlGroup>();
            var groupName = "";
            ProjectConfigControlGroup oneGroup = null;
            for (var i = 0; i < _projectConfigSetting.Count; i++)
            {
                if (groupName != string.Format(_projectConfigSetting[i].GroupName + _projectConfigSetting[i].TabGroup))
                {
                    oneGroup = new ProjectConfigControlGroup();
                    var tGroupBox = new GroupBox();
                    oneGroup.Device = _projectConfigSetting[i].TabGroup;
                    tGroupBox.Header = _projectConfigSetting[i].GroupName;
                    oneGroup.GroupBox = tGroupBox;
                    projectConfigControlGroups.Add(oneGroup);
                    groupName = string.Format(_projectConfigSetting[i].GroupName + _projectConfigSetting[i].TabGroup);
                }

                var tControlItem = GetProjectConfigControlItem(i);
                if (oneGroup != null) oneGroup.ProjectConfigControls.Add(tControlItem);
            }

            return projectConfigControlGroups;
        }

        private ProjectConfigControl GetProjectConfigControlItem(int i)
        {
            var tControlItem = new ProjectConfigControl();
            var tLabel = new Label
                {Content = _projectConfigSetting[i].Name, ToolTip = _projectConfigSetting[i].Description};
            tControlItem.Name = tLabel;

            switch (_projectConfigSetting[i].Type)
            {
                case "Text":
                    var textBox = new TextBox();
                    textBox.MaxLength = _projectConfigSetting[i].MaxLength;
                    if (_projectConfigSetting[i].Name == "ProjectName" && !string.IsNullOrEmpty(ProjectName))
                        textBox.Text = ProjectName;
                    else
                        textBox.Text =
                            _projectConfig.Exists(x =>
                                x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                    StringComparison.OrdinalIgnoreCase) && x.Name.Equals(_projectConfigSetting[i].Name,
                                    StringComparison.OrdinalIgnoreCase))
                                ? _projectConfig.Find(x =>
                                        x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                            StringComparison.OrdinalIgnoreCase) &&
                                        x.Name.Equals(_projectConfigSetting[i].Name,
                                            StringComparison.OrdinalIgnoreCase))
                                    .Value
                                : _projectConfigSetting[i].Default;
                    tControlItem.Value = textBox;
                    break;
                case "List":
                    var comboBox = new ComboBox();
                    for (var j = 0; j < _projectConfigSetting[i].OptionValue.Count; j++)
                        if (!string.IsNullOrEmpty(_projectConfigSetting[i].OptionValue[j]))
                            comboBox.Items.Add(_projectConfigSetting[i].OptionValue[j]);
                    comboBox.SelectedItem =
                        _projectConfig.Exists(x =>
                            x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                StringComparison.OrdinalIgnoreCase) && x.Name.Equals(_projectConfigSetting[i].Name,
                                StringComparison.OrdinalIgnoreCase))
                            ? _projectConfig.Find(x =>
                                x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                    StringComparison.OrdinalIgnoreCase) && x.Name.Equals(_projectConfigSetting[i].Name,
                                    StringComparison.OrdinalIgnoreCase)).Value
                            : _projectConfigSetting[i].Default;
                    tControlItem.Value = comboBox;
                    break;
                case "CheckBox":
                    var checkBox = new CheckBox();
                    checkBox.IsChecked =
                        _projectConfig.Exists(x =>
                            x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                StringComparison.OrdinalIgnoreCase) && x.Name.Equals(_projectConfigSetting[i].Name,
                                StringComparison.OrdinalIgnoreCase))
                            ? _projectConfig
                                .Find(x =>
                                    x.GroupName.Equals(_projectConfigSetting[i].GroupName,
                                        StringComparison.OrdinalIgnoreCase) &&
                                    x.Name.Equals(_projectConfigSetting[i].Name, StringComparison.OrdinalIgnoreCase))
                                .Value.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase)
                            : _projectConfigSetting[i].Default
                                .Equals("TRUE", StringComparison.CurrentCultureIgnoreCase);
                    tControlItem.Value = checkBox;
                    break;
            }

            return tControlItem;
        }

        private void CreateControls(List<ProjectConfigControlGroup> pControls)
        {
            var maxHeight = 0;
            foreach (var group in pControls.GroupBy(x => x.Device))
            {
                var value = SetFormSize(group.ToList());
                if (value > maxHeight)
                    maxHeight = value;
            }

            Width = LabelWidth + ItemWidthGap + TextBoxWidth + 70;
            Height = maxHeight + 90;
            SrVmain.MaxHeight = maxHeight + 90;
            SrVmain.VerticalScrollBarVisibility =
                maxHeight == MaxFormSize ? ScrollBarVisibility.Auto : ScrollBarVisibility.Hidden;
            var tabControl = new TabControl();
            tabControl.Height = maxHeight + 90;
            foreach (var group in pControls.GroupBy(x => x.Device))
            {
                var device = group.First().Device;
                var tabItem = new TabItem();
                tabItem.Header = device;
                tabItem.Name = device;
                tabItem.Content = GetStackPanel(group.ToList(), maxHeight);
                tabControl.Items.Add(tabItem);
            }

            GridMain.Children.Add(tabControl);
        }

        private StackPanel GetStackPanel(List<ProjectConfigControlGroup> lControls, double maxHeight)
        {
            var stackPanel = new StackPanel();
            stackPanel.Height = maxHeight + 100;
            for (var i = 0; i < lControls.Count; i++)
            {
                Grid.SetRow(lControls[i].GroupBox, i);
                stackPanel.Children.Add(lControls[i].GroupBox);

                var tGrid = new Grid();
                lControls[i].GroupBox.Content = tGrid;
                lControls[i].GroupBox.HorizontalAlignment = HorizontalAlignment.Left;
                lControls[i].GroupBox.Margin = new Thickness(10, 0, 0, 0);
                AddGridRow(tGrid, lControls[i].ProjectConfigControls.Count, RowHeight);

                for (var j = 0; j < lControls[i].ProjectConfigControls.Count; j++)
                {
                    var tLabel = lControls[i].ProjectConfigControls[j].Name;
                    tLabel.Width = LabelWidth;
                    tLabel.HorizontalAlignment = HorizontalAlignment.Left;
                    tLabel.VerticalAlignment = VerticalAlignment.Center;
                    Grid.SetRow(tLabel, j);
                    tGrid.Children.Add(tLabel);

                    if (lControls[i].ProjectConfigControls[j].Value is TextBox)
                    {
                        var tTextBox = (TextBox) lControls[i].ProjectConfigControls[j].Value;
                        tTextBox.Width = TextBoxWidth;
                        tTextBox.HorizontalAlignment = HorizontalAlignment.Left;
                        tTextBox.VerticalAlignment = VerticalAlignment.Center;
                        tTextBox.Margin = new Thickness(LabelWidth + ItemWidthGap, 0, 0, 0);
                        Grid.SetRow(tTextBox, j);
                        tGrid.Children.Add(tTextBox);
                    }
                    else if (lControls[i].ProjectConfigControls[j].Value is ComboBox)
                    {
                        var tComboBox = (ComboBox) lControls[i].ProjectConfigControls[j].Value;
                        tComboBox.Width = ComboBoxWidth;
                        tComboBox.Margin = new Thickness(LabelWidth + ItemWidthGap, 0, 0, 0);
                        tComboBox.HorizontalAlignment = HorizontalAlignment.Left;
                        tComboBox.VerticalAlignment = VerticalAlignment.Center;
                        if (lControls[i].GroupBox.Header.ToString()
                            .Equals("Device", StringComparison.OrdinalIgnoreCase) && tLabel.Content.ToString()
                            .Equals("Type", StringComparison.OrdinalIgnoreCase))
                            tComboBox.SelectionChanged += Changed;
                        Grid.SetRow(tComboBox, j);
                        tGrid.Children.Add(tComboBox);
                    }
                    else if (lControls[i].ProjectConfigControls[j].Value is CheckBox)
                    {
                        var tCheckBox = (CheckBox) lControls[i].ProjectConfigControls[j].Value;
                        tCheckBox.Width = ComboBoxWidth;
                        tCheckBox.Margin = new Thickness(LabelWidth + ItemWidthGap, 0, 0, 0);
                        tCheckBox.HorizontalAlignment = HorizontalAlignment.Left;
                        tCheckBox.VerticalAlignment = VerticalAlignment.Center;
                        tCheckBox.IsChecked = tCheckBox.IsChecked;
                        Grid.SetRow(tCheckBox, j);
                        tGrid.Children.Add(tCheckBox);
                    }
                }
            }

            return stackPanel;
        }

        private void AddGridRow(Grid pGrid, int pRowCount, int pRowHeight)
        {
            for (var i = 0; i < pRowCount; i++)
            {
                var tRd = new RowDefinition();
                tRd.Height = new GridLength(pRowHeight);
                pGrid.RowDefinitions.Add(tRd);
            }
        }

        private int SetFormSize(List<ProjectConfigControlGroup> pControls)
        {
            var totalHeight = 34;
            for (var i = 0; i < pControls.Count; i++)
            {
                totalHeight += 2 * ItemWidthGap;
                totalHeight += pControls[i].ProjectConfigControls.Count * RowHeight;
            }

            if (totalHeight > MaxFormSize) totalHeight = MaxFormSize;
            return totalHeight;
        }

        public void SaveConfig()
        {
            var lProjectName = ((TextBox) _controlList[0].ProjectConfigControls[0].Value).Text;
            _projectConfig.Clear();

            for (var i = 1; i < _controlList.Count; i++)
            {
                var tGroupName = _controlList[i].GroupBox.Header.ToString();
                for (var j = 0; j < _controlList[i].ProjectConfigControls.Count; j++)
                {
                    var tName = _controlList[i].ProjectConfigControls[j].Name.Content.ToString();
                    var tValue = "";
                    if (_controlList[i].ProjectConfigControls[j].Value is TextBox)
                        tValue = ((TextBox) _controlList[i].ProjectConfigControls[j].Value).Text;
                    else if (_controlList[i].ProjectConfigControls[j].Value is ComboBox)
                        tValue = ((ComboBox) _controlList[i].ProjectConfigControls[j].Value).Text;
                    else if (_controlList[i].ProjectConfigControls[j].Value is CheckBox)
                        tValue = ((CheckBox) _controlList[i].ProjectConfigControls[j].Value).IsChecked.ToString();
                    var tConfigRow = new ProjectConfigRow();
                    tConfigRow.GroupName = tGroupName;
                    tConfigRow.Name = tName;
                    tConfigRow.Value = tValue;
                    _projectConfig.Add(tConfigRow);
                }
            }

            ProjectName = lProjectName;
            var deviceGroup = _controlList.FirstOrDefault(p => p.GroupBox.Header.ToString().Equals("Device"));
            if (deviceGroup != null)
            {
                var deviceTypeObj =
                    deviceGroup.ProjectConfigControls.FirstOrDefault(p => p.Name.Content.ToString().Equals("Type"));
                if (deviceTypeObj != null)
                    DeviceType = ((ComboBox) deviceTypeObj.Value).Text;
            }

            DialogResult = true;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            SaveConfig();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Do you want to exit?", "Question", MessageBoxButton.YesNo, MessageBoxImage.Question) ==
                MessageBoxResult.Yes)
                //this.Close();
                DialogResult = false;
        }

        #region Const

        private const int MaxFormSize = 868;
        private const int ItemWidthGap = 10;
        private const int RowHeight = 28;
        private const int LabelWidth = 200;
        private const int TextBoxWidth = 100;
        private const int ComboBoxWidth = 100;

        #endregion

        #region field

        private List<ProjectConfigRow> _projectConfig;
        private List<ProjectConfigSettingRow> _projectConfigSetting;
        private readonly List<ProjectConfigControlGroup> _controlList;
        public string ProjectName;
        public string DeviceType;

        #endregion
    }

    internal class ProjectConfigControlGroup
    {
        public string Device;
        public GroupBox GroupBox;
        public List<ProjectConfigControl> ProjectConfigControls = new List<ProjectConfigControl>();
    }

    internal class ProjectConfigControl
    {
        public Label Name;
        public Control Value;
    }
}