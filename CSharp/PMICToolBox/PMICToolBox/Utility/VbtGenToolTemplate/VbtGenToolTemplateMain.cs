using PmicAutomation.MyControls;
using PmicAutomation.Utility.VbtGenToolTemplate.Base;
using PmicAutomation.Utility.VbtGenToolTemplate.Input;
using Library.Function.ErrorReport;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.VbtGenToolTemplate
{
    public class VbtGenToolTemplateMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _tcm;
        private readonly string _outputPath;

        private readonly List<TcmSheet> _testPlanSheets = new List<TcmSheet>();

        public VbtGenToolTemplateMain(string tcm, string outputPath)
        {
            _tcm = tcm;
            _outputPath = outputPath;
        }

        public VbtGenToolTemplateMain(VbtGenToolGenerator vbtGenToolGenerator)
        {
            _appendText = vbtGenToolGenerator.AppendText;
            _tcm = vbtGenToolGenerator.FileOpen_TCM.ButtonTextBox.Text;
            _outputPath = vbtGenToolGenerator.FileOpen_OutputPath.ButtonTextBox.Text;
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            ReadFiles();

            GenFiles();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private void ReadFiles()
        {
            using (ExcelPackage inputExcel =
                new ExcelPackage(new FileInfo(_tcm)))
            {
                foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                {
                    _appendText.Invoke("Starting to read TCM " + sheet.Name + " ...", Color.Black);
                    TcmReader testPlanReader = new TcmReader();
                    TcmSheet testPlanSheet = testPlanReader.ReadSheet(sheet);
                    _testPlanSheets.Add(testPlanSheet);
                }
            }
        }

        private void GenFiles()
        {
            foreach (TcmSheet testPlanSheet in _testPlanSheets)
            {
                if (testPlanSheet != null)
                {
                    VbtTestPlanSheet test = new VbtTestPlanSheet { Name = testPlanSheet.Name };
                    test.Rows.AddRange(testPlanSheet.ConvertTestRows());
                    test.GenTxt(Path.Combine(_outputPath,
                        testPlanSheet.Name + "_Test.txt"));

                    VbtTestPlanSheet post = new VbtTestPlanSheet { Name = testPlanSheet.Name };
                    post.Rows.AddRange(testPlanSheet.ConvertPostRows());
                    post.GenTxt(Path.Combine(_outputPath,
                        testPlanSheet.Name + "_Post.txt"));

                    VbtTestPlanSheet other = new VbtTestPlanSheet { Name = testPlanSheet.Name };
                    other.Rows.AddRange(testPlanSheet.ConvertOtherRows());
                    other.GenTxt(Path.Combine(_outputPath,
                        testPlanSheet.Name + ".txt"));
                }
            }
        }
    }
}