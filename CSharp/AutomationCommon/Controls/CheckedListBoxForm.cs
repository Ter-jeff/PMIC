using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace AutomationCommon.Controls
{
    public class CheckedListBoxForm : Form
    {
        private CheckedListBox _checkedListBox1;
        private List<string> _checkedEnables;

        public CheckedListBoxForm()
        {
            InitializeComponent();
        }

        public void BindingCheckedListBox(List<string> enables, List<string> checkedList)
        {
            foreach (var enable in enables)
            {
                if (checkedList == null || checkedList.Exists(x => x.Equals(enable, StringComparison.CurrentCultureIgnoreCase)))
                    _checkedListBox1.Items.Add(enable, true);
                else
                    _checkedListBox1.Items.Add(enable, false);
            }

            AdjustSize(_checkedListBox1);
        }

        private void AdjustSize(CheckedListBox checkedListBox)
        {
            int h = checkedListBox.ItemHeight * checkedListBox.Items.Count;
            checkedListBox.Height = h + checkedListBox.Height - checkedListBox.ClientSize.Height;
            Height = checkedListBox.Height + 70;
        }

        private void InitializeComponent()
        {
            _checkedListBox1 = new CheckedListBox();
            SuspendLayout();
            // 
            // checkedListBox1
            // 
            _checkedListBox1.FormattingEnabled = true;
            _checkedListBox1.Location = new Point(12, 12);
            _checkedListBox1.Name = "_checkedListBox1";
            _checkedListBox1.Size = new Size(260, 89);
            _checkedListBox1.TabIndex = 0;
            // 
            // CheckedListBoxForm
            // 
            ClientSize = new Size(284, 262);
            Controls.Add(_checkedListBox1);
            Name = "CheckedListBoxForm";
            FormClosed += CheckedListBoxForm_FormClosed;
            ResumeLayout(false);

        }

        public List<string> GetCheckedList()
        {

            return _checkedEnables;
        }

        private void CheckedListBoxForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _checkedEnables = new List<string>();
            foreach (var item in _checkedListBox1.CheckedItems)
            {
                _checkedEnables.Add(item.ToString());
            }
        }

    }

}