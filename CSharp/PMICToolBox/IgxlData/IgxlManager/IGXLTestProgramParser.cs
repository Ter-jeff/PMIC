using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IgxlData.IgxlManager
{
    public class IGXLTestProgramParser
    {
        public static IgxlProgram getIGXLTestProgram(string p_strFileName)
        {
            var exportfolder = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "Teradyne", "PMICToolBox", "ExportTmp", "exportProg");
            TestProgramUtility.ExportWorkBookCmd(p_strFileName, exportfolder);
            IgxlProgram l_Rtn = new IgxlProgram(p_strFileName);
            l_Rtn.LoadIgxlProgramAsync(exportfolder);

            return l_Rtn;
        }
    }
}
