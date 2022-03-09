using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Library.Common;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using Library.Input;
using Library.Output;
using Library.DataStruct;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Diagnostics;

namespace Library
{
    public class MainLogic
    {
        #region Field
        private static MainLogic _instance = null;
        public List<DiffResultLogRow> logDiffResultlst = new List<DiffResultLogRow>();
        public List<InstanceCompareResult> newrequirelist = new List<InstanceCompareResult>();
        #endregion

        #region Property
 
        #endregion

        #region Constructor
        private MainLogic()
        {

        }
        #endregion

        #region Static Function

        public static MainLogic Instance()
        {
            if (_instance == null)
            {
                _instance = new MainLogic();
            }
            return _instance;
        }
 
        #endregion

        #region Member Function

        public void MainFlow()
        {
            try
            {
                CommonData.GetInstance().Init();
                string currentTimeStr = DateTime.Now.ToString("yyyymmddhhmmss");
                CommonData.GetInstance().worker.ReportProgress(5);
                //compare datalog
                List<InstanceCompareResult> compareResult = CompareDatalog(CommonData.GetInstance().BaseTxtDatalogPath,
                    CommonData.GetInstance().CompareTxtDatalogPath);
                //generate output report
                CommonData.GetInstance().UISateInfo = "Write Comparison report...";
                CommonData.GetInstance().worker.ReportProgress(90);
                string reportFIlePath = CommonData.GetInstance().OutputPath + "\\DiffReport_" + currentTimeStr + ".xlsx";
                new ReportWriter().GenerateDiffReport(compareResult, logDiffResultlst, reportFIlePath);

                CommonData.GetInstance().UISateInfo = "Completed!(100%)";
                CommonData.GetInstance().worker.ReportProgress(100);
                MessageBox.Show("Compare complete!", "Message", MessageBoxButtons.OK);

            }
            catch(Exception e){
                MessageBox.Show("Meet Error: " + e.Message);
                CommonData.GetInstance().UISateInfo = "Failed!";
                CommonData.GetInstance().worker.ReportProgress(0);
            }
        }

        // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator add start
        public void MainFlowForCommandLine()
        {
            try
            {
                string currentTimeStr = DateTime.Now.ToString("yyyymmddhhmmss");
                //compare datalog
                List<InstanceCompareResult> compareResult = CompareDatalog(CommonData.GetInstance().BaseTxtDatalogPath,
                    CommonData.GetInstance().CompareTxtDatalogPath, true);
                //generate output report
                string reportFIlePath = CommonData.GetInstance().OutputPath + "\\DiffReport_" + currentTimeStr + ".xlsx";
                new ReportWriter().GenerateDiffReport(compareResult, logDiffResultlst, reportFIlePath);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator add end

        // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg start
        //public List<InstanceCompareResult> CompareDatalog(string baseDatalogFile, string compareDatalog)
        public List<InstanceCompareResult> CompareDatalog(string baseDatalogFile, string compareDatalog, bool isCommandLine = false)
        // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg end
        {
            DatalogReader datalogReader = new DatalogReader();
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg start
            //CommonData.GetInstance().UISateInfo = "Read Base Datalog...";
            //CommonData.GetInstance().worker.ReportProgress(10);
            if (!isCommandLine)
            {
                CommonData.GetInstance().UISateInfo = "Read Base Datalog...";
                CommonData.GetInstance().worker.ReportProgress(10);
            }
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg end
            string copyRealDatalogFile = CommonData.GetInstance().OutputPath + "\\" + Path.GetFileName(baseDatalogFile).Replace(".txt", "_AutoRunCopy.txt");
            string copyRefDatalogFile = CommonData.GetInstance().OutputPath + "\\" + Path.GetFileName(compareDatalog).Replace(".txt", "_AutoRunCopy.txt");
            File.Copy(baseDatalogFile, copyRealDatalogFile, true);
            File.Copy(compareDatalog, copyRefDatalogFile, true);
            List<Device> baseDevicelist = datalogReader.Read(copyRealDatalogFile);
            List<TestInstanceLogBase> ForTestNameLoglstbase = new List<TestInstanceLogBase>(datalogReader.ForTestNameLoglst);
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg start
            //CommonData.GetInstance().UISateInfo = "Read Compare Datalog...";
            //CommonData.GetInstance().worker.ReportProgress(40);
            if (!isCommandLine)
            {
                CommonData.GetInstance().UISateInfo = "Read Compare Datalog...";
                CommonData.GetInstance().worker.ReportProgress(40);
            }
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg end
            List<Device> compareDevicelist = datalogReader.Read(copyRefDatalogFile);
            List<TestInstanceLogBase> ForTestNameLoglstcompare = new List<TestInstanceLogBase>(datalogReader.ForTestNameLoglst);
            File.Delete(copyRealDatalogFile);
            File.Delete(copyRefDatalogFile);
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg start
            //CommonData.GetInstance().UISateInfo = "Compare Datalogs...";
            //CommonData.GetInstance().worker.ReportProgress(70);
            if (!isCommandLine)
            {
                CommonData.GetInstance().UISateInfo = "Compare Datalogs...";
                CommonData.GetInstance().worker.ReportProgress(70);
            }
            // 2022-02-08  Steven Chen    #306	          Create command-line for datalog comparator chg end
            List<InstanceCompareResult> compareResult = new List<InstanceCompareResult>();
            
            DiffResultLogRow logResultRow = null;
            foreach (TestInstanceLogBase test in ForTestNameLoglstbase)
            {                
                TestInstanceLogBase compareTN = ForTestNameLoglstcompare.Find(s => s.KeyWord.Equals(test.KeyWord, StringComparison.OrdinalIgnoreCase) 
                && s.DuplicateIndex == test.DuplicateIndex);
                
                if (compareTN == null)
                {
                    logResultRow = test.ConvertToReportRow(test.Row.ToString(), "");
                    logResultRow.Result = DiffResultType.OnlyInBaseDatalog;
                    logResultRow.BasedInst = test.InstanceName;
                    logDiffResultlst.Add(logResultRow);
                    continue;
                }
                //Compare real datalog item and reference datalog item
                bool logEquals = test.Compare(compareTN, out logResultRow);
                if (logEquals == false)
                {
                    //result = false;
                    logDiffResultlst.Add(logResultRow);
                }

            }
            foreach (TestInstanceLogBase test1 in ForTestNameLoglstcompare)
            {
                TestInstanceLogBase baseTN =
                    ForTestNameLoglstbase.Find(
                    s => s.KeyWord.Equals(test1.KeyWord, StringComparison.OrdinalIgnoreCase) && s.DuplicateIndex == test1.DuplicateIndex);

                if (baseTN == null)
                {
                    //result = false;
                    DiffResultLogRow diffResultRow = test1.ConvertToReportRow("", test1.Row.ToString());
                    diffResultRow.Result = DiffResultType.OnlyInCompareDatalog;
                    diffResultRow.ComparedInst = test1.InstanceName;
                    logDiffResultlst.Add(diffResultRow);
                    continue;
                }

            }



            foreach (Device baseDevice in baseDevicelist)
            {
                Device compareDevice = compareDevicelist.Find(s => s.DeviceNumber.Equals(baseDevice.DeviceNumber));
                if (compareDevice != null)
                {
                    List<InstanceCompareResult> deviceCompareResult = baseDevice.Compare(compareDevice);
                    if (deviceCompareResult != null && deviceCompareResult.Count > 0)
                    {
                        compareResult.AddRange(deviceCompareResult);
                    }
                }
                else
                {
                    foreach (TestInstanceItem baseInstance in baseDevice.TestInstanceItemlist)
                    {
                        InstanceCompareResult result = new InstanceCompareResult(baseInstance.DeviceNumber, baseInstance.InstanceName, DiffResultType.OnlyInBaseDatalog, null);
                        compareResult.Add(result);
                    }
                }
            }
            foreach (Device compareDevice in compareDevicelist.FindAll(s => !baseDevicelist.Exists(m => m.DeviceNumber.Equals(s.DeviceNumber))).ToList())
            {
                foreach (TestInstanceItem compareInstance in compareDevice.TestInstanceItemlist)
                {
                    InstanceCompareResult result = new InstanceCompareResult(compareInstance.DeviceNumber, compareInstance.InstanceName, DiffResultType.OnlyInCompareDatalog, null);
                    compareResult.Add(result);
                }
            }
            return compareResult;
        }
        
        #endregion



    }
}
