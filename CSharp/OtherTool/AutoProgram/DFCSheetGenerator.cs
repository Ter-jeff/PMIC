using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using NewPostProcessorLib.Controller;
using NewPostProcessorLib.IGLinkProcessor.DataStructure;
using NewPostProcessorLib.LocalSpec;
using NewPostProcessorLib.Utility.VbtModuleManager;
using NewPostProcessorLib.Utility.UtilityFunctions;

namespace NewPostProcessorLib.Bussiness
{
    public class DFCSheetGenerator
    {


        /* Member function */
        public static void Generate()
        {
            GeneralFunc.WriteMessage("Generating DFC sheet... ");

            var outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);

           
                // create blank instance sheet for each char plan sheet
                //var czInstSheet = new InstanceSheet("TestInst_CZ_" + planSheet.SheetName);
                
                // export cz inst sheet
            var czFileName = Path.Combine(outputFolder, "DFC_List.txt");
                WriteDFC(czFileName);
            
        }

        public static void WriteDFC(string fileName)
        {
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                return;
            using (var sw = new StreamWriter(fileName, false))
            {
                //Test Instance	DFC Info
                sw.WriteLine("Test Instance");
                for (int rowindex = 2; rowindex <= LocalSpecs.DFCSheet.Dimension.Rows; rowindex++)
                {
                    var instance = LocalSpecs.DFCSheet.Cells[rowindex, 1].Text;                    
                    sw.WriteLine(instance);
                }

            }
        }
         
    }
}
