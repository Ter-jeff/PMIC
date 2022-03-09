using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.VBT
{
    public class BasManager
    {
        private List<VbtFunctionBase> _vbtFunctionBases;

        public BasManager(string tempPath)
        {
            _vbtFunctionBases = ReadAllLib(tempPath);
        }

        public List<VbtFunctionBase> ReadAllLib(string dirLib)
        {
            List<VbtFunctionBase> vbtFunctionBaseList = new List<VbtFunctionBase>();
            if (Directory.Exists(dirLib))
            {
                var fileList = Directory.GetFiles(dirLib);
                foreach (var basFile in fileList)
                {
                    var extension = Path.GetExtension(basFile);
                    if (extension != null && extension.ToLower() == ".bas")
                        vbtFunctionBaseList.AddRange(ReadBasFile(basFile));
                }

                foreach (string sub in Directory.GetDirectories(dirLib))
                {
                    var subList = Directory.GetFiles(sub);
                    foreach (var basFile in subList)
                    {
                        var extension = Path.GetExtension(basFile);
                        if (extension != null && extension.ToLower() == ".bas")
                            vbtFunctionBaseList.AddRange(ReadBasFile(basFile));
                    }
                }
            }
            else
            {
                throw new Exception("Read bas lib failed, the bas directory not exist!");
            }
            return vbtFunctionBaseList;
        }

        private List<VbtFunctionBase> ReadBasFile(string fileName)
        {
            List<VbtFunctionBase> vbtFunctionBaseList = new List<VbtFunctionBase>();
            var fInfo = new FileInfo(fileName);
            var sr = new StreamReader(fInfo.FullName);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (!Regex.IsMatch(line, @"^\s*(Public\s)?Function.*\("))
                    continue;
                string functionName = Regex.Match(line, @"(Public\s)?Function\s(?<func>\w+)\(").Groups["func"].ToString();
                string paramStr;
                line = line.TrimEnd('_');
                if (Regex.IsMatch(line, @"\(.*\)"))
                    paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                else
                {
                    paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                    while (((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)")))
                    {
                        line = line.TrimEnd('_');
                        if (!Regex.IsMatch(line, @"\s*\'"))
                            paramStr += line;
                    }
                    if (line != null)
                        paramStr += Regex.Match(line, @"(?<str>.*)\)").Groups["str"].ToString();
                }
                var paramters = GetParamters(paramStr);

                var newVbt = new VbtFunctionBase(functionName);
                newVbt.FileName = fileName;

                for (int a = 0; a < paramters.Count; a++)
                {
                    if (paramters[a].Name.ToLower() == "step_")
                    {
                        paramters.RemoveAt(a);
                        break;
                    }
                }

                newVbt.Parameters = string.Join(",", paramters.Select(x => x.Name));
                newVbt.ParameterDefaults = string.Join(",", paramters.Select(x => x.Default));
                vbtFunctionBaseList.Add(newVbt);
            }
            sr.Close();
            return vbtFunctionBaseList;
        }

        public List<VbtFunctionBase> ReadFileList(List<string> fileList)
        {
            List<VbtFunctionBase> vbtFunctionBaseList = new List<VbtFunctionBase>();
            foreach (var file in fileList)
            {
                if (File.Exists(file))
                {
                    var extension = Path.GetExtension(file);
                    if (extension != null && extension.ToLower() == ".bas")
                        vbtFunctionBaseList.AddRange(ReadBasFile(file));
                }
                else
                {
                    throw new Exception("Read bas lib failed, the bas directory not exist!");
                }
            }
            return vbtFunctionBaseList;
        }
     
        public List<Proc> GetProcOfLine(string fileName)
        {
            List<Proc> procs = new List<Proc>();
            var fInfo = new FileInfo(fileName);
            using (var sr = new StreamReader(fInfo.FullName))
            {
                string line;
                int cnt = -1;
                while ((line = sr.ReadLine()) != null)
                {
                    line = RemoveComment(line);
                    cnt++;
                    if (!(Regex.IsMatch(line, @"Function\s|Sub\s|Enum\s", RegexOptions.IgnoreCase)))
                        continue;

                    if (!(Regex.IsMatch(line, @"Public Function\s|Public Sub\s|Public Enum\s|Private Function\s|Private Sub\s|Private Enum\s", RegexOptions.IgnoreCase)))
                    {
                        // ex: dim FunctionA as string
                        if (!Regex.IsMatch(line.Trim(), @"Function\s|Sub\s|Enum\s", RegexOptions.IgnoreCase))
                            continue;
                    }

                    string type = "";
                    var subName = GetFunctionName(line, ref type);
                    string paramStr;
                    line = line.TrimEnd('_');
                    if (Regex.IsMatch(line, @"\(.*\)"))
                        paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                    else
                    {
                        paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                        while (((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)")))
                        {
                            line = line.TrimEnd('_');
                            cnt++;
                            if (!Regex.IsMatch(line, @"\s*\'"))
                                paramStr += line;
                        }
                        if (line != null)
                        {
                            cnt++;
                            paramStr += Regex.Match(line, @"(?<str>.*)\)").Groups["str"].ToString();
                        }
                    }
                    var paramters = GetParamters(paramStr);

                    if (string.IsNullOrEmpty(subName))
                        continue;
                    int tempCnt = cnt;
                    bool endFlag = false;
                    while (endFlag == false && (line = sr.ReadLine()) != null)
                    {
                        line = RemoveComment(line);
                        cnt++;
                        if (line.ToUpper().Contains("END FUNCTION") && type.Equals("Function", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                        else if (line.ToUpper().Contains("END SUB") && type.Equals("Sub", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                        else if (line.ToUpper().Contains("END ENUM") && type.Equals("Enum", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                    }
                    procs.Add(new Proc { Name = subName, Start = tempCnt, End = cnt, Type = type, Parameters = paramters });
                }
            }
            return procs;
        }

        private List<Parameter> GetParamters(string paramStr)
        {
            if (string.IsNullOrEmpty(paramStr))
                return new List<Parameter>();

            List<Parameter> parameters = new List<Parameter>();
            foreach (var str in paramStr.Split(','))
            {
                Parameter parameter = new Parameter();
                var parameterName = Regex.Match(str, @"(?<param>\w+)\sAs\s", RegexOptions.IgnoreCase).Groups["param"].ToString();
                var paramterType = Regex.Match(paramStr, @"(?<param>\w+)\sAs\s(?<type>[^,]*)", RegexOptions.IgnoreCase).Groups["type"].ToString();
                parameter.Name = parameterName;
                if (paramterType.Contains("="))
                {
                    parameter.Type = paramterType.Replace(" ", "").Split('=')[0].Replace("\"", "");
                    parameter.Default = paramterType.Replace(" ", "").Split('=')[1].Replace("\"", "");
                }
                else
                {
                    parameter.Type = paramterType;
                    parameter.Default = "";
                }
                if (!(parameterName == null || paramterType == null))
                    parameters.Add(parameter);
            }

            return parameters;
        }

        private string RemoveComment(string line)
        {
            if (line.IndexOf('\'') != -1)
                return line.Substring(0, line.IndexOf('\''));
            return line;
        }

        private string GetFunctionName(string line, ref string type)
        {
            if (Regex.IsMatch(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase))
            {
                type = "Function";
                return Regex.Match(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase).Groups["func"].ToString();
            }

            if (Regex.IsMatch(line, @"Sub\s(?<func>\w+)\(", RegexOptions.IgnoreCase))
            {
                type = "Sub";
                return Regex.Match(line, @"Sub\s(?<func>\w+)\(", RegexOptions.IgnoreCase).Groups["func"].ToString();
            }

            if (Regex.IsMatch(line, @"Enum\s(?<func>\w+)", RegexOptions.IgnoreCase))
            {
                type = "Enum";
                return Regex.Match(line, @"Enum\s(?<func>\w+)", RegexOptions.IgnoreCase).Groups["func"].ToString();
            }

            return "";
        }

        public void MergeBasFile(string fileName1, string fileName2, string newFileName)
        {
            //Merge fileName1 intot fileName2
            if (File.Exists(fileName1) && File.Exists(fileName1))
            {
                var proCFile1 = GetProcOfLine(fileName1);
                var proCFile2 = GetProcOfLine(fileName2);
                List<Proc> skipLineList = new List<Proc>();
                foreach (var line in proCFile2)
                {
                    if (proCFile1.Exists(x => x.Name.Equals(line.Name, StringComparison.OrdinalIgnoreCase)))
                        skipLineList.Add(line);
                }
                var file1 = ReadBasContent(fileName1);
                var file2 = ReadBasContent(fileName2, skipLineList);
                using (var sw = new StreamWriter(fileName2))
                {
                    sw.WriteLine("Attribute VB_Name = \"" + Path.GetFileNameWithoutExtension(newFileName) + "\"");
                    foreach (var line in file1)
                        sw.WriteLine(line);
                    foreach (var line in file2)
                        sw.WriteLine(line);
                }
                File.Move(fileName2, newFileName);
            }
        }

        private List<string> ReadBasContent(string fileName, List<Proc> skipLineList = null)
        {
            List<string> list = new List<string>();
            int cnt = 0;
            using (var sw = new StreamReader(fileName))
            {
                string line;
                while ((line = sw.ReadLine()) != null)
                {
                    cnt++;
                    bool isSkip = false;
                    if (skipLineList != null)
                        foreach (var skipLine in skipLineList)
                        {
                            if (cnt >= skipLine.Start && cnt < skipLine.End)
                            {
                                isSkip = true;
                                break;
                            }
                        }

                    if (!isSkip)
                        list.Add(line);
                }
            }
            return list;
        }

        public VbtFunctionBase GetFunctionByName(string functionName)
        {
            var resultVbt = new VbtFunctionBase();
            resultVbt.FunctionName = functionName;

            if (_vbtFunctionBases == null)
                return resultVbt;

            var newVbt = _vbtFunctionBases.Find(a => a.FunctionName.ToLower() == functionName.ToLower());
            if (newVbt == null)
                return resultVbt;

            newVbt.Parameters = newVbt.Parameters;
            newVbt.ParameterDefaults = newVbt.ParameterDefaults;
            newVbt.FileName = newVbt.FileName;
            newVbt.FunctionName = newVbt.FunctionName;
            newVbt.Args = newVbt.Args.ToList();
            return newVbt;
        }
    }

    public class Parameter
    {
        public string Name;
        public string Default;
        public string Type;
    }

    public class Proc
    {
        public string Name;
        public int Start;
        public int End;
        public string Type;
        public List<Parameter> Parameters;
    }
}
