using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PmicAutogen.Inputs.TestPlan.Reader;

namespace PmicAutogen
{
    public partial class VDDRefForm : Form
    {
        private Dictionary<string, string> _PinList = new Dictionary<string, string>();
        private Dictionary<string, VddLevelsRow> _VDDPinInfoList = new Dictionary<string, VddLevelsRow>();

        public Dictionary<string, VddLevelsRow> RefVddPins { get; set; }

        public VDDRefForm(Dictionary<string, string> pinList, Dictionary<string, VddLevelsRow> vDDPinList)
        {
            InitializeComponent();
            _PinList = pinList;
            _VDDPinInfoList = vDDPinList;

            DataGridViewComboBoxColumn comboBoxColumn = (DataGridViewComboBoxColumn)dataGridView1.Columns[2];
            comboBoxColumn.Items.Add("No Reference");
            comboBoxColumn.Items.AddRange(_VDDPinInfoList.Keys.ToArray());

            for (int i = 0; i < _PinList.Count; i++)
            {
                var pinInfo = _PinList.ElementAt(i);
                string[] row = new string[] { pinInfo.Key, pinInfo.Value };
                dataGridView1.Rows.Add(row);
                DataGridViewComboBoxCell dataGridViewComboBoxCell = (DataGridViewComboBoxCell)dataGridView1.Rows[i].Cells[2];
                dataGridViewComboBoxCell.Value = "No Reference";
            }
            dataGridView1.CellValueChanged += new DataGridViewCellEventHandler(dataGridView1_CellValueChanged);
        }

        public void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "ReferencePin")
            {
                DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[2];
                var pin = cell.Value.ToString();
                if (pin == "No Reference")
                    return;

                string pinVolt = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                VddLevelsRow vddLevelRow = _VDDPinInfoList[pin];

                if (!CheckMapping(pinVolt, vddLevelRow))
                {
                    string msg = string.Format("Voltage Not Match!\n  IO Pin {0} : {1}\n Vdd Pin {2} : {3}.",
                                dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString(),
                                pinVolt,
                                pin, vddLevelRow.Nv);
                    MessageBox.Show(msg, "Warnning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private bool CheckMapping(string pinVolt, VddLevelsRow vddLevelRow)
        {
            if (vddLevelRow.Nv != vddLevelRow.Lv ||
                vddLevelRow.Nv != vddLevelRow.Hv ||
                (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
            {
                return true;
            }

            if (vddLevelRow.Nv != pinVolt)
            {
                return false;
            }
            return true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            RefVddPins = new Dictionary<string, VddLevelsRow>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var combox = row.Cells[2];
                if (combox.Value.ToString() != "No Reference")
                {
                    var ioPin = row.Cells[0].Value;
                    VddLevelsRow vddLevelRow = _VDDPinInfoList[combox.Value.ToString()];

                    if (vddLevelRow.Nv != vddLevelRow.Lv ||
                        vddLevelRow.Nv != vddLevelRow.Hv ||
                        (vddLevelRow.ULv != "" && vddLevelRow.Nv != vddLevelRow.ULv) ||
                        (vddLevelRow.UHv != "" && vddLevelRow.Nv != vddLevelRow.UHv))
                    {
                        RefVddPins.Add(ioPin.ToString(), vddLevelRow);
                    }
                }
            }
            DialogResult = DialogResult.OK;
        }
    }
}
