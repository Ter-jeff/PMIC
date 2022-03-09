using PmicAutomation.MyControls;
using PmicAutomation.Utility.Relay.Base;
using PmicAutomation.Utility.Relay.Input;
using Library.Function;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.Relay.Output
{
    public class GenADG1414File
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _outputPath;
        private const int MaxCnt = 8;

        public GenADG1414File(MyForm.RichTextBoxAppend appendText, string outputPath)
        {
            _appendText = appendText;
            _outputPath = outputPath;
        }

        public void GenADG1414Matrix(List<ADG1414Group> ADG1414Group)
        {
            var lines = GenADG1414MatrixLines(ADG1414Group);
            File.WriteAllLines(Path.Combine(_outputPath, "ADG1414_Matrix.txt"), lines);
        }

        private List<string> GenADG1414MatrixLines(List<ADG1414Group> adg141PinRows)
        {
            const string DCol = "		IN		";
            const string BitsLine = "				128	64	32	16	8	4	2	1";
            const string BlankLine = "											";
            const string EndLine = "											End";

            List<string> lines = new List<string>
            {
                "ADG_Matrix											",
                "			ADG IN	D8	D7	D6	D5	D4	D3	D2	D1",
            };

            string line;
            for (int i = 0; i < adg141PinRows.Count; i++)
            {
                ADG1414Group adg1414Group = adg141PinRows[i];
                int rowNum = i + 1;
                line = DCol;
                for (int num = MaxCnt - 1; num >= 0; num--)
                {
                    line += adg1414Group.DNames[num];
                    if (num != 0) line += "\t";
                }
                lines.Add(line);
                lines.Add(BitsLine);

                for (int num = 0; num < MaxCnt; num++)
                {
                    line = "";
                    int Snum = num + 1;
                    if (num == 0)
                    {
                        line = rowNum.ToString() + "\t" + adg1414Group.DesignName + "\t";
                    }
                    else
                    {
                        line = "\t\t";
                    }

                    line += "S" + Snum.ToString() + "\t";
                    line += adg1414Group.SNames[num]; 

                    for (int num2 = MaxCnt - 1; num2 >= 0; num2--)
                    {
                        if (num2 == num && adg1414Group.SNames[num] != "" && adg1414Group.DNames[num] != "")
                        {
                            line += "\t" + "v";
                        }
                        else
                        {
                            line += "\t";
                        }
                    }
                    lines.Add(line);
                }
                lines.Add(BlankLine);
            }
            lines.Add(EndLine);
            return lines;
        }

        public void GenADG1414CONTROL(List<ADG1414Group> adg141PinRows)
        {
            var lines = GenADG1414CONTROLLines(adg141PinRows);
            File.WriteAllLines(Path.Combine(_outputPath, "ADG1414_CONTROL.txt"), lines);
        }

        private List<string> GenADG1414CONTROLLines(List<ADG1414Group> adg141PinRows)
        {

            const string FuncName = "Public Function ADG1414_CONTROL(";
            const string FuncParaBlank = "                                ";
            // 'U3901 As Long, _
            const string SNameParaPat = "{0} As Long, _";
            // g_ADG1414ArgList = "U3901,U3902,U3903,U3907,U3904,U3905"
            const string AdgListPat = "g_ADG1414ArgList = \"{0}\"";
            // g_ADG1414Data(0) = U3901
            const string AdgListItemPat = "g_ADG1414Data({0}) = {1}";
            // Call SPI_BYTE_WRITE1(U3905,"ADG1414_PINS")
            const string CallFirstPat = "Call SPI_BYTE_WRITE1({0}, \"ADG1414_PINS\")";
            const string CallSecondPat = "Call SPI_BYTE_WRITE2({0}, \"ADG1414_PINS\")";


            List<string> lines = new List<string>();

            for (int i = adg141PinRows.Count - 1; i >= 0; i--)
            {
                ADG1414Group adg1414Group = adg141PinRows[i];
                if (i == adg141PinRows.Count - 1)
                {
                    lines.Add(FuncName + string.Format(SNameParaPat, adg1414Group.DesignName));
                }
                else
                {
                    lines.Add(FuncParaBlank + string.Format(SNameParaPat, adg1414Group.DesignName));
                }
            }

            lines.AddRange(new List<string>{
                "                                Optional ADGOverWrite As Boolean = False, _",
                "                                Optional ADGResetN As Boolean = False) As Long",
                "",
                "On Error GoTo ErrHandler",
                "Dim sCurrentFuncName As String:: sCurrentFuncName = \"ADG1414_CONTROL\"",
                "TheHdw.DIB.PowerOn = True",
                ""
            });

            lines.Add(string.Format(AdgListPat, string.Join(",", adg141PinRows.Select(r => r.DesignName))));

            for (int i = 0; i < adg141PinRows.Count; i++)
            {
                ADG1414Group adg1414Group = adg141PinRows[i];
                lines.Add(string.Format(AdgListItemPat, i, adg1414Group.DesignName));
            }

            lines.AddRange(new List<string>{
                "",
                "Call SetFRCPath(\"ADG1414_SCLK\", True)",
                "Call SetFRCPath(\"ADG1414_RESET\", True)",
                "",
                "With TheHdw.Digital.Pins(\"ADG1414_PINS\")",
                "    .StartState = chStartOff",
                "    .InitState = chInitoff",
                "End With",
                "",
                "TheHdw.Protocol.Ports(\"ADG1414_PINS\").Enabled = True",
                "TheHdw.Protocol.Ports(\"ADG1414_PINS\").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_First",
                "TheHdw.Protocol.Ports(\"ADG1414_PINS\").NWire.HRAM.Setup.WaitForEvent = False",
                "TheHdw.Protocol.ModuleRecordingEnabled = True",
                ""
            });

            for (int i = 0; i < adg141PinRows.Count; i++)
            {
                ADG1414Group adg1414Group = adg141PinRows[adg141PinRows.Count - i - 1];
                if (i == 0)
                {
                    lines.Add(string.Format(CallFirstPat, adg1414Group.DesignName));
                }
                else
                {
                    lines.Add(string.Format(CallSecondPat, adg1414Group.DesignName));
                }
            }

            lines.AddRange(new List<string>{
                "",
                "TheHdw.Protocol.Ports(\"ADG1414_PINS\").Halt",
                "TheHdw.Protocol.Ports(\"ADG1414_PINS\").Enabled = False",
                "",
                "Exit Function",
                "ErrHandler:",
                "    TheExec.AddOutput \"<Error> \" + sCurrentFuncName + \":: please Check it out.\"",
                "    TheExec.Datalog.WriteComment \"<Error> \" + sCurrentFuncName + \":: please check it out.\"",
                "    If AbortTest Then Exit Function Else Resume Next",
                "End Function"
            });

            return lines;
        }

        public void GenSINExtract(List<ADG1414Group> adg141PinRows)
        {
            var lines = GenSINExtractLines(adg141PinRows);
            File.WriteAllLines(Path.Combine(_outputPath, "SIN_Extract.txt"), lines);
        }

        private List<string> GenSINExtractLines(List<ADG1414Group> adg141PinRows)
        {
            // '___U3901
            const string DesignNamePat = "'___{0}";
            //	Case "U3901_S1":SIN_Extract ="BUCK0_LX4_UP1600_S1"
            const string SNamePat = "	Case \"{0}_S{1}\":SIN_Extract =\"{2}\"";

            List<string> lines = new List<string>
            {
                "Public Function SIN_Extract(Trace As String) As String",
                "Dim SplitArray() As String",
                "Dim Result As String",
                "Dim i As Double",
                "",
                "Dim FunctionName As String:: FunctionName = \"SIN_Extract\"",
                "On Error GoTo ErrHandler",
                "",
                "If InStr(Trace, \",\") > 0 Then",
                "	SplitArray() = Split(Trace, \",\")",
                "Else",
                "	ReDim SplitArray(0)",
                "	SplitArray(0) = Trace",
                "End If",
                "",
                "For i = 0 To UBound(SplitArray())",
                "	Trace = SplitArray(i)",
                "	Select Case Trace",
            };

            for (int i = 0; i < adg141PinRows.Count; i++)
            {
                ADG1414Group adg1414Group = adg141PinRows[i];
                // '___U3901
                lines.Add(string.Format(DesignNamePat, adg1414Group.DesignName));

                for (int num = 0; num < MaxCnt; num++)
                {
                    if (adg1414Group.SNames[num] != "")
                        //	Case "U3901_S1":SIN_Extract ="BUCK0_LX4_UP1600_S1"
                        lines.Add(string.Format(SNamePat, adg1414Group.DesignName, num + 1, adg1414Group.SNames[num]));
                }
            }
            lines.AddRange(new List<string>{
                "	End Select",
                "	Result = SIN_Extract & \",\" & Result",
                "Next",
                "",
                "SIN_Extract = Left(Result, Len(Result) - 1)",
                "",
                "Exit Function",
                "",
                "ErrHandler:",
                "	TheExec.Datalog.WriteComment \"<Error> \" + funcName + \":: please check it out.\"",
                "	If AbortTest Then Exit Function Else Resume Next",
                "",
                "End Function"
            });
            return lines;
        }
    }
}