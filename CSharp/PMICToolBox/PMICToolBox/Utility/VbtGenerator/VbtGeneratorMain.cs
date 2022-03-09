using PmicAutomation.MyControls;
using PmicAutomation.Utility.VbtGenerator.Function;
using PmicAutomation.Utility.VbtGenerator.Input;
using Library.Function;
using Library.Function.ErrorReport;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PmicAutomation.Utility.VbtGenerator
{
    public class VbtGeneratorMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _template;
        private readonly string _table;
        private readonly string _basFile;
        private readonly string _outputPath;

        public VbtGeneratorMain(string template, string table, string basFile, string outputPath)
        {
            _template = template;
            _table = table;
            _basFile = basFile;
            _outputPath = outputPath;
        }

        public VbtGeneratorMain(VbtGeneratorFrom vbtGenerator)
        {
            _appendText = vbtGenerator.AppendText;
            _template = vbtGenerator.FileOpen_Template.ButtonTextBox.Text;
            _table = vbtGenerator.FileOpen_Table.ButtonTextBox.Text;
            _basFile = vbtGenerator.FileOpen_BasFile.ButtonTextBox.Text;
            _outputPath = vbtGenerator.FileOpen_OutputPath.ButtonTextBox.Text;
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            GenVbt();

            GenTable();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private void GenTable()
        {
            if (!string.IsNullOrEmpty(_basFile))
            {
                BasParser basParser = new BasParser(_template, _appendText);
                string file = Path.Combine(_outputPath, "Mapping.xlsx");
                if (File.Exists(file))
                {
                    File.Delete(file);
                }

                using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    foreach (string bas in _basFile.Split(','))
                    {
                        _appendText.Invoke("Starting to generate mapping table ...", Color.Black);
                        List<Comment> vbt = basParser.GetComment(bas);
                        ExcelWorksheet sheet = workBook.AddSheet(Path.GetFileNameWithoutExtension(bas));
                        basParser.GenTable(sheet, vbt);
                        package.Save();
                    }
                }
            }
        }

        private void GenVbt()
        {
            List<string> templates = _template.Split(',').ToList();
            bool isPMIC_IDS_TP = false; // indicate that input excel containing PMIC_IDS sheet

            string fileName = Path.GetFileName(Path.ChangeExtension(_table, "bas"));
            if (fileName != null)
            {
                string outputFile = Path.Combine(_outputPath, fileName);
                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }

                if (!string.IsNullOrEmpty(_table))
                {
                    using (ExcelPackage inputExcel = new ExcelPackage(new FileInfo(_table)))
                    {
                        if (inputExcel.Workbook.Worksheets.ToList().Exists(p => p.Name.Equals("PMIC_IDS")))
                            isPMIC_IDS_TP = true;

                        List<string> lines = new List<string>();
                        foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                        {
                            if (isPMIC_IDS_TP && !sheet.Name.Equals("PMIC_IDS"))
                                continue;

                            string file = templates.First();
                            foreach (string template in templates)
                            {
                                if (template != null &&
                                    Path.GetFileNameWithoutExtension(template).Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    file = template;
                                    break;
                                }
                            }

                            BasParser basParser = new BasParser(file, _appendText);
                            _appendText.Invoke("Starting to read mapping table " + sheet.Name + " ...", Color.Black);
                            TableFactory tableFactory = TableFactory.CreateFactory(isPMIC_IDS_TP);
                            List<TableSheet> tables = tableFactory.ReadSheet(sheet);
                            _appendText.Invoke("Starting to generate Bas " + sheet.Name + " ...", Color.Black);
                            lines.AddRange(basParser.GenBas(tables));
                            lines.Add("");
                        }

                        BasFileManage.GenBasFile(outputFile, lines);
                    }
                }
            }
        }
    }
}