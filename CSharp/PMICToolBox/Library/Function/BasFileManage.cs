using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Library.Function
{
    public class BasFile
    {
        public BasFile()
        {
            Procedures = new List<Procedure>();
        }

        public string Name { get; set; }
        public List<Procedure> Procedures { get; set; }
    }

    public class Procedure
    {
        public int End;
        public string Name;
        public int Start;
        public string Type;
        public bool ExistOnErrorGoTo;
        public bool ExistOnErrorResumeNext;
        public bool ModifiedErrorName;
        public bool ExistErrHandler;
        public bool ExistFuncName;
        public bool StopKeyword;
        public bool MsgBoxKeyword;

        public Procedure copyObject()
        {
            string data = JsonConvert.SerializeObject(this);
            Procedure copy = JsonConvert.DeserializeObject<Procedure>(data);
            return copy;
        }
    }

    public static class BasFileManage
    {
        public static void GenBasFile(string outputFile, List<string> lines)
        {
            List<Procedure> procedures = GetProcedureOfLine(lines);
            string module = Path.GetFileNameWithoutExtension(outputFile);
            if (procedures.Count > 999)
            {
                decimal loop = Math.Ceiling((decimal)procedures.Count / 999);
                for (int i = 0; i < loop; i++)
                {
                    List<string> newLines = new List<string> { "Attribute VB_Name = \"" + module + (i + 1) + "\"" };
                    int last = i == loop - 1 ? procedures.Last().End : procedures[999 * (i + 1) - 1].End;
                    newLines.AddRange(lines.GetRange(procedures[999 * i].Start, last - procedures[999 * i].Start + 1));
                    string file = Path.GetDirectoryName(outputFile) + "\\" +
                                  Path.GetFileNameWithoutExtension(outputFile) + (i + 1) +
                                  Path.GetExtension(outputFile);
                    File.WriteAllLines(file, newLines);
                }
            }
            else
            {
                lines.Insert(0, "Attribute VB_Name = \"" + module + "\"");
                File.WriteAllLines(outputFile, lines);
            }
        }

        public static List<Procedure> GetProcedureOfLine(string file)
        {
            return GetProcedureOfLineForErrHandler(File.ReadAllLines(file).ToList());
        }

        public static List<Procedure> GetProcedureOfLine(List<string> lines)
        {
            List<Procedure> procedures = new List<Procedure>();
            for (int i = 0; i < lines.Count; i++)
            {
                string line = RemoveComment(lines[i]);
                string type;
                string subName = GetFunctionName(line, out type);
                if (string.IsNullOrEmpty(type))
                {
                    continue;
                }

                for (int j = i + 1; j < lines.Count; j++)
                {
                    line = RemoveComment(lines[j]);
                    if (line.IndexOf("END FUNCTION", StringComparison.OrdinalIgnoreCase) > -1 ||
                        line.IndexOf("END SUB", StringComparison.OrdinalIgnoreCase) > -1 ||
                        line.IndexOf("END ENUM", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        procedures.Add(new Procedure { Name = subName, Start = i, End = j, Type = type });
                        break;
                    }
                }
            }

            return procedures;
        }

        public static List<Procedure> GetProcedureOfLineForErrHandler(List<string> lines)
        {
            var multiLine = false;
            var procedures = new List<Procedure>();
            for (int i = 0; i < lines.Count; i++)
            {
                var startIndex = i;
                string line = RemoveComment(lines[i]);
                string type;
                string subName = GetFunctionName(line, out type);
                if (string.IsNullOrEmpty(type) || type == "Enum")
                {
                    continue;
                }
                if (line.EndsWith(" _"))
                    multiLine = true;

                var procedure = new Procedure { Name = subName, Start = startIndex, Type = type };

                for (int j = i + 1; j < lines.Count; j++)
                {
                    line = RemoveComment(lines[j]);
                    if (multiLine)
                    {
                        if (Regex.IsMatch(line.TrimEnd(), @"\)") && !line.EndsWith(" _"))
                        {
                            procedure.Start = j;
                            multiLine = false;
                        }
                    }

                    if (line.IndexOf("on error",StringComparison.OrdinalIgnoreCase)>-1)
                    {
                        procedure.ExistOnErrorGoTo = true;
                    }

                    if (line.IndexOf("dim sCurrentFuncName as string", StringComparison.OrdinalIgnoreCase)>-1)
                    {
                        procedure.ExistFuncName = true;
                    }

                    if (line.IndexOf("handler:",StringComparison.OrdinalIgnoreCase)>-1)
                    {
                        procedure.ExistErrHandler = true;
                    }
                    
                    if (line.IndexOf("END FUNCTION", StringComparison.OrdinalIgnoreCase) > -1 ||
                        line.IndexOf("END SUB", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        procedure.End = j;
                        procedures.Add(procedure);
                        break;
                    }
                }
            }

            return procedures;
        }

        public static List<Procedure> GetProcedureStartEndPosition(List<string> lines, List<Procedure> procedures)
        {
            var multiLine = false;
            for (int i = 0; i < lines.Count; i++)
            {
                var startIndex = i;
                string line = RemoveComment(lines[i]);
                string type;
                string subName = GetFunctionName(line, out type);
                bool need2add = false;
                if (string.IsNullOrEmpty(type) || type == "Enum" || subName.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (line.EndsWith(" _"))
                    multiLine = true;

                var procedure = procedures.FirstOrDefault(p => p.Name.Equals(subName));
                if (procedure == null)
                {
                    procedure = new Procedure { Name = subName, Start = startIndex, Type = type };
                    need2add = true;
                }
                else
                    procedure.Start = startIndex;

                for (int j = i + 1; j < lines.Count; j++)
                {
                    line = RemoveComment(lines[j]);
                    if (multiLine)
                    {
                        if (Regex.IsMatch(line.TrimEnd(), @"\)") && !line.EndsWith(" _"))
                        {
                            procedure.Start = j;
                            multiLine = false;
                        }
                    }

                    if (line.IndexOf("END FUNCTION", StringComparison.OrdinalIgnoreCase) > -1 ||
                        line.IndexOf("END SUB", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        procedure.End = j;
                        if (need2add)
                            procedures.Add(procedure);
                        break;
                    }
                }
            }

            return procedures;
        }

        public static string RemoveComment(string line)
        {
            if (line.IndexOf('\'') > -1)
            {
                return line.Substring(0, line.IndexOf('\''));
            }

            return line;
        }

        public static string GetFunctionName(string line, out string type)
        {
            if (Regex.IsMatch(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase))
            {
                type = "Function";
                return Regex.Match(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase).Groups["func"]
                    .ToString();
            }

            if (Regex.IsMatch(line, @"Enum\s(?<func>\w+)", RegexOptions.IgnoreCase))
            {
                type = "Enum";
                return Regex.Match(line, @"Enum\s(?<func>\w+)", RegexOptions.IgnoreCase).Groups["func"].ToString();
            }

            if (Regex.IsMatch(line, @"Sub\s(?<func>\w+)\(", RegexOptions.IgnoreCase))
            {
                type = "Sub";
                return Regex.Match(line, @"Sub\s(?<func>\w+)\(", RegexOptions.IgnoreCase).Groups["func"].ToString();
            }

            type = "";
            return "";
        }

        public static void MergeBasFile(string file1, string file2, string outputFile)
        {
            if (File.Exists(file1) && File.Exists(file1))
            {
                List<Procedure> procedure1 = GetProcedureOfLine(file1);
                List<Procedure> procedure2 = GetProcedureOfLine(file2);
                List<Procedure> skipLineList = new List<Procedure>();
                foreach (Procedure line in procedure2)
                {
                    if (procedure1.Exists(x => x.Name.Equals(line.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        skipLineList.Add(line);
                    }
                }

                string[] lines1 = File.ReadAllLines(file1);
                List<string> lines2 = ReadBasContent(file2, skipLineList);

                File.WriteAllLines(outputFile, lines1);
                File.AppendAllLines(outputFile, lines2);
            }
        }

        public static List<string> ReadBasContent(string file, List<Procedure> skipLines = null)
        {
            List<string> list = new List<string>();
            List<string> lines = File.ReadAllLines(file).ToList();
            int lastIndex = lines.FindLastIndex(x =>
                x.StartsWith("Attribute VB_Name = ", StringComparison.OrdinalIgnoreCase));
            lines = lines.GetRange(lastIndex + 1, lines.Count - lastIndex - 1);
            for (int i = 0; i < lines.Count; i++)
            {
                if (skipLines != null)
                {
                    if (skipLines.Any(x => x.Start <= i && i <= x.End))
                    {
                        continue;
                    }
                }

                list.Add(lines[i]);
            }

            return list;
        }

        public static void AddComments(string file, List<string> comments)
        {
            if (File.Exists(file))
            {
                List<string> lines = File.ReadAllLines(file).ToList();
                int lastIndex = lines.FindLastIndex(x =>
                    x.StartsWith("Attribute VB_Name = ", StringComparison.OrdinalIgnoreCase));
                if (comments != null)
                {
                    lines.InsertRange(lastIndex + 1, comments);
                }

                File.WriteAllLines(file, lines);
            }
        }
   
       


    }
}