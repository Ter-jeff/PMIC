using IgxlData.IgxlManager;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PmicAutogen.GenerateIgxl;
using PmicAutogen.InputPackages;

namespace PmicAutogen.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {

            var TestPlan = @"C:\01.Jeffli\AutoGen_Team\trunk\Demo\Input\ProjectName_A0_TestPlan_20220226.xlsx";
            var app = new Application();
            var _workbook = app.Workbooks.Open(TestPlan);
            var _inputPackageAutomation = new InputPackage();

            var pmicGenerator = new PmicGenerator();
            pmicGenerator.Run(_workbook, _inputPackageAutomation);

            var exportMain = new IgxlManagerMain();

            //var igxlItems = GetAllIgxlItems();
            //exportMain.GenIgxlProgram(igxlItems, LocalSpecs.TarDir, LocalSpecs.CurrentProject, TestProgram.IgxlWorkBk,
            //    Response.Report, LocalSpecs.TargetIgxlVersion);

            //if (EpplusErrorManager.GetErrorCount() > 0)
            //    GenErrorReport();
        }
    }
}
