using PmicAutomation.Utility.AHBEnum;
using PmicAutomation.Utility.PA;
using PmicAutomation.Utility.Relay;
using PmicAutomation.Utility.VbtGenerator;
using PmicAutomation.Utility.VbtGenToolTemplate;
using Library.Function;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace UnitTest
{
    [TestClass]
    public class UnitTest
    {
        private static string CurrentDirectory = Directory.GetCurrentDirectory().Replace(@"\bin\Debug", "");
        private readonly string _projectPath = CurrentDirectory + "\\Utility\\";
        private readonly string _winMergeFile = CurrentDirectory + "\\WinMerge\\WinMergeU.exe";

        [TestMethod]
        public void AhbEnum()
        {
            string abhRegister = Path.Combine(_projectPath, "AHBEnum\\Input\\TestPlan.xlsx");
            string outputPath = Path.Combine(_projectPath, "AHBEnum\\Output");
            string expectedPath = Path.Combine(_projectPath, "AHBEnum\\Expected");
            string batFilePath = Path.Combine(_projectPath, "AHBEnum");
            string report = Path.Combine(_projectPath, "AHBEnum\\AhbNum.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            string maxBitWidth = "8";
            bool regNameAndFieldName = true;
            bool regNameOnly = false;

            AhbMain ahbMain = new AhbMain(abhRegister, outputPath, maxBitWidth, regNameAndFieldName, regNameOnly);
            ahbMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void Pa()
        {
            List<string> files = new List<string>()
            {
                Path.Combine(_projectPath, "PA\\Input\\Suzuka_CP_PA.csv")
            };
            string uflexConfigPath = Path.Combine(_projectPath,
                "PA\\Input\\TesterConfig_PMIC.xml");
            string outputPath = Path.Combine(_projectPath, "PA\\Output");
            List<string> hexVs = new List<string>();
            string device = "PMIC";
            string expectedPath = Path.Combine(_projectPath, "PA\\Expected");
            string batFilePath = Path.Combine(_projectPath, "PA");
            string report = Path.Combine(_projectPath, "PA\\PA.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            PaMain paMain = new PaMain(files, uflexConfigPath, outputPath, hexVs, device);
            paMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void Relay()
        {
            string comPin = Path.Combine(_projectPath,
                "Relay\\Input\\Component_Pin_Report.xlsx");
            string relayConfig = Path.Combine(_projectPath,
                "Relay\\Input\\RelayConfig.xlsm");
            string outputPath = Path.Combine(_projectPath, "Relay\\Output");
            string expectedPath = Path.Combine(_projectPath, "Relay\\Expected");
            string batFilePath = Path.Combine(_projectPath, "Relay");
            string report = Path.Combine(_projectPath, "Relay\\Relay.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            RelayMain relayMain = new RelayMain(comPin, relayConfig, outputPath);
            relayMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void VbtGeneratorDcEnum()
        {
            string template =
                Path.Combine(_projectPath,
                    "VbtGenerator\\DCEnum\\Input\\SP_Conti_Pins_Cond.tmp")
                + "," + Path.Combine(_projectPath,
                    "VbtGenerator\\DCEnum\\Input\\SP_Leak_Pins_Cond.tmp");
            string table = Path.Combine(_projectPath,
                "VbtGenerator\\DCEnum\\Input\\DCEnum.xlsx");
            string basFile = "";
            string outputPath = Path.Combine(_projectPath, "VbtGenerator\\DCEnum\\Output");
            string expectedPath = Path.Combine(_projectPath, "VbtGenerator\\DCEnum\\Expected");
            string batFilePath = Path.Combine(_projectPath, "VbtGenerator\\DCEnum");
            string report = Path.Combine(_projectPath, "VbtGenerator\\DCEnum\\DCEnum.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            VbtGeneratorMain vbtGeneratorMain = new VbtGeneratorMain(template, table, basFile, outputPath);
            vbtGeneratorMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void VbtGeneratorIdsEnum()
        {
            string template =
                Path.Combine(_projectPath,
                    "VbtGenerator\\IDSEnum\\Input\\ActiveSetting.tmp")
                + "," + Path.Combine(_projectPath,
                    "VbtGenerator\\IDSEnum\\Input\\OffSetting.tmp");
            string table = Path.Combine(_projectPath,
                "VbtGenerator\\IDSEnum\\Input\\SP_IDS_Pins_Cond.xlsx");
            string basFile = "";
            string outputPath = Path.Combine(_projectPath, "VbtGenerator\\IDSEnum\\Output");
            string expectedPath = Path.Combine(_projectPath, "VbtGenerator\\IDSEnum\\Expected");
            string batFilePath = Path.Combine(_projectPath, "VbtGenerator\\IDSEnum");
            string report = Path.Combine(_projectPath, "VbtGenerator\\IDSEnum\\IDSEnum.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            VbtGeneratorMain vbtGeneratorMain = new VbtGeneratorMain(template, table, basFile, outputPath);
            vbtGeneratorMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void VbtGeneratorPowerUp()
        {
            string template = Path.Combine(_projectPath,
                "VbtGenerator\\PowerUp\\Input\\PowerUp.tmp");
            string table = Path.Combine(_projectPath,
                "VbtGenerator\\PowerUp\\Input\\VDD_Level_From_TestPlan.xlsx");
            string basFile = "";
            string outputPath = Path.Combine(_projectPath, "VbtGenerator\\PowerUp\\Output");
            string expectedPath = Path.Combine(_projectPath, "VbtGenerator\\PowerUp\\Expected");
            string batFilePath = Path.Combine(_projectPath, "VbtGenerator\\PowerUp");
            string report = Path.Combine(_projectPath, "VbtGenerator\\PowerUp\\PowerUp.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            VbtGeneratorMain vbtGeneratorMain = new VbtGeneratorMain(template, table, basFile, outputPath);
            vbtGeneratorMain.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void VbtGenToolTemplate()
        {
            string tcm = Path.Combine(_projectPath,
                "VbtGenToolTemplate\\Input\\Sylvester_TCM.xlsx");
            string outputPath = Path.Combine(_projectPath, "VbtGenToolTemplate\\Output");
            string expectedPath = Path.Combine(_projectPath, "VbtGenToolTemplate\\Expected");
            string batFilePath = Path.Combine(_projectPath, "VbtGenToolTemplate");
            string report = Path.Combine(_projectPath, "VbtGenToolTemplate\\VbtGenToolTemplate.xlsx");
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            VbtGenToolTemplateMain vbtGenToolTemplate = new VbtGenToolTemplateMain(tcm, outputPath);
            vbtGenToolTemplate.WorkFlow();
            FileComparision fileComparision =
                new FileComparision(expectedPath, outputPath, report, batFilePath, _winMergeFile);
            if (fileComparision.IsFolderCompareFail())
            {
                Assert.Fail();
            }
        }
    }
}