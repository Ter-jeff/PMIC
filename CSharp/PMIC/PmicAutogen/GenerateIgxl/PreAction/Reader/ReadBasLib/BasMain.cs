using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using IgxlData.VBT;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib
{
    public class BasMain
    {
        private readonly List<SrcInfoRow> _srcInfoRows;
        private Dictionary<string, string> _functionNames;

        public BasMain(List<SrcInfoRow> srcInfoRows = null)
        {
            _srcInfoRows = srcInfoRows;
        }

        public void WorkFlow(string outputPath)
        {
            _functionNames = new Dictionary<string, string>();
            var libPath = LocalSpecs.BasLibraryPath;
            if (Directory.Exists(outputPath))
                Directory.Delete(outputPath, true);
            Directory.CreateDirectory(outputPath);

            CopyAllFolder(libPath, outputPath);

            try
            {
                TestProgram.VbtFunctionLib.AddVbtFunctionRange(ReadLocalLib(outputPath));
            }
            catch (Exception e)
            {
                throw new Exception("Read bas lib failed! " + e.Message);
            }
        }

        public List<VbtFunctionBase> ReadLocalLib(string dirLib)
        {
            var vbtFunctionBaseList = new List<VbtFunctionBase>();
            if (Directory.Exists(dirLib))
            {
                var fileList = Directory.GetFiles(dirLib, @"VBT_*");
                foreach (var basFile in fileList)
                {
                    var extension = Path.GetExtension(basFile);
                    if (extension.ToLower() == ".bas")
                        vbtFunctionBaseList.AddRange(ReadBasFile(basFile));
                }

                foreach (var sub in Directory.GetDirectories(dirLib))
                {
                    string[] subList;
                    if (sub.IndexOf("Wireless", StringComparison.OrdinalIgnoreCase) > 0 ||
                        sub.IndexOf("PMIC", StringComparison.OrdinalIgnoreCase) > 0)
                        subList = Directory.GetFiles(sub, @"VBT_*", SearchOption.AllDirectories);
                    else
                        subList = Directory.GetFiles(sub, @"VBT_*");
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

        private void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            if (!Directory.Exists(targetPath))
                Directory.CreateDirectory(targetPath);

            //Now Create all of the directories
            foreach (var path in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
                Directory.CreateDirectory(path.Replace(sourcePath, targetPath));

            //Copy all the files & Replaces any files with the same name
            foreach (var path in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                var dir = path.Replace(sourcePath, targetPath);
                File.Copy(path, dir, true);
            }
        }

        private void CopyAllFolder(string sourcePath, string targetPath)
        {
            if (!Directory.Exists(targetPath))
                Directory.CreateDirectory(targetPath);
            foreach (var file in Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories))
            {
                var extension = Path.GetExtension(file);
                if (extension.ToLower() == ".bas" || extension.ToLower() == ".cls" || extension.ToLower() == ".frm" ||
                    extension.ToLower() == ".frx")
                {
                    File.Copy(file, Path.Combine(targetPath, Path.GetFileName(file)), true);
                    AddComment(Path.Combine(targetPath, Path.GetFileName(file)));
                }
            }
        }

        //private void CopyFolderAndSubFolder(string sourcePath, string targetPath, List<string> excludeFolders = null)
        //{
        //    if (!Directory.Exists(targetPath))
        //        Directory.CreateDirectory(targetPath);
        //    FileCopy(sourcePath, targetPath);
        //    foreach (var sub1 in Directory.GetDirectories(sourcePath))
        //    {
        //        if (excludeFolders != null && excludeFolders.Exists(x => x.Equals(sub1.Split('\\').Last(), StringComparison.OrdinalIgnoreCase)))
        //            continue;
        //        var fileName = Path.GetFileName(sub1.Split('\\').Last());
        //        if (fileName != null)
        //        {
        //            var newFolder1 = Path.Combine(targetPath, fileName);
        //            if (!Directory.Exists(newFolder1))
        //                Directory.CreateDirectory(newFolder1);
        //            FileCopy(sub1, newFolder1);
        //        }
        //    }
        //}

        //private void FileCopy(string sourcePath, string targetPath)
        //{
        //    foreach (var file in Directory.GetFiles(sourcePath))
        //    {
        //        var extension = Path.GetExtension(file);
        //        if (extension.ToLower() == ".bas" || extension.ToLower() == ".cls" || extension.ToLower() == ".frm" || extension.ToLower() == ".frx")
        //        {
        //            File.Copy(file, Path.Combine(targetPath, Path.GetFileName(file)), true);
        //            AddComment(Path.Combine(targetPath, Path.GetFileName(file)));
        //        }
        //    }
        //}

        private List<VbtFunctionBase> ReadBasFile(string fileName)
        {
            var vbtFunctionBaseList = new List<VbtFunctionBase>();
            var fInfo = new FileInfo(fileName);
            var sr = new StreamReader(fInfo.FullName);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (!Regex.IsMatch(line, @"^\s*(Public\s)?Function.*\("))
                    continue;
                var functionName = Regex.Match(line, @"(Public\s)?Function\s(?<func>\w+)\(").Groups["func"].ToString();
                string paramStr;
                line = line.TrimEnd('_');
                if (Regex.IsMatch(line, @"\(.*\)"))
                {
                    paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                }
                else
                {
                    paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                    while ((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)"))
                    {
                        line = line.TrimEnd('_');
                        if (!Regex.IsMatch(line, @"\s*\'"))
                            paramStr += line;
                    }

                    if (line != null)
                        paramStr += Regex.Match(line, @"(?<str>.*)\)").Groups["str"].ToString();
                }

                var parameters = GetParameters(paramStr);

                if (_functionNames != null)
                {
                    if (!_functionNames.ContainsKey(functionName))
                    {
                        _functionNames.Add(functionName, fileName);
                    }
                    else
                    {
                        var outString = "Duplicate VBT function : \"" + functionName + "\" in bas file " + fileName +
                                        "<>" + _functionNames[functionName];
                        ErrorManager.AddError(EnumErrorType.DuplicateVbtModule, EnumErrorLevel.Error,
                            "", 1, outString, fileName, _functionNames[functionName]);
                        Response.Report(outString, EnumMessageLevel.Error, 0);
                    }
                }

                var newVbt = new VbtFunctionBase(functionName);
                newVbt.FileName = fileName;

                for (var a = 0; a < parameters.Count; a++)
                    if (parameters[a].Name.ToLower() == "step_")
                    {
                        parameters.RemoveAt(a);
                        break;
                    }

                newVbt.Parameters = string.Join(",", parameters.Select(x => x.Name));
                newVbt.ParameterDefaults = string.Join(",", parameters.Select(x => x.Default));
                vbtFunctionBaseList.Add(newVbt);
            }

            sr.Close();
            return vbtFunctionBaseList;
        }

        public List<Procedure> GetProcOfLine(string fileName)
        {
            var procedures = new List<Procedure>();
            var fInfo = new FileInfo(fileName);
            using (var sr = new StreamReader(fInfo.FullName))
            {
                string line;
                var cnt = -1;
                while ((line = sr.ReadLine()) != null)
                {
                    line = RemoveComment(line);
                    cnt++;
                    if (!Regex.IsMatch(line, @"Function\s|Sub\s|Enum\s", RegexOptions.IgnoreCase))
                        continue;

                    if (!Regex.IsMatch(line,
                            @"Public Function\s|Public Sub\s|Public Enum\s|Private Function\s|Private Sub\s|Private Enum\s",
                            RegexOptions.IgnoreCase))
                        // ex: dim FunctionA as string
                        if (!Regex.IsMatch(line.Trim(), @"^Function\s|Sub\s|Enum\s", RegexOptions.IgnoreCase))
                            continue;

                    var type = "";
                    var subName = GetFunctionName(line, ref type);
                    string paramStr;
                    line = line.TrimEnd('_');
                    if (Regex.IsMatch(line, @"\(.*\)"))
                    {
                        paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                    }
                    else
                    {
                        paramStr = Regex.Match(line, @"\((?<str>.*)").Groups["str"].ToString();
                        while ((line = sr.ReadLine()) != null && !Regex.IsMatch(line, @".*\)"))
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

                    var parameters = GetParameters(paramStr);

                    if (string.IsNullOrEmpty(subName))
                        continue;
                    var tempCnt = cnt;
                    var endFlag = false;
                    while (endFlag == false && (line = sr.ReadLine()) != null)
                    {
                        line = RemoveComment(line);
                        cnt++;
                        if (line.ToUpper().Contains("END FUNCTION") &&
                            type.Equals("Function", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                        else if (line.ToUpper().Contains("END SUB") &&
                                 type.Equals("Sub", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                        else if (line.ToUpper().Contains("END ENUM") &&
                                 type.Equals("Enum", StringComparison.CurrentCultureIgnoreCase))
                            endFlag = true;
                    }

                    procedures.Add(new Procedure
                    {
                        Name = subName,
                        Start = tempCnt,
                        End = cnt,
                        Type = type,
                        Parameters = parameters
                    });
                }
            }

            return procedures;
        }

        private List<Parameter> GetParameters(string paramStr)
        {
            if (string.IsNullOrEmpty(paramStr))
                return new List<Parameter>();

            var parameters = new List<Parameter>();
            foreach (var str in paramStr.Split(','))
            {
                var parameter = new Parameter();
                var parameterName = Regex.Match(str, @"(?<param>\w+)\sAs\s", RegexOptions.IgnoreCase).Groups["param"]
                    .ToString();
                var parameterType = Regex.Match(paramStr, @"(?<param>\w+)\sAs\s(?<type>[^,]*)", RegexOptions.IgnoreCase)
                    .Groups["type"].ToString();
                parameter.Name = parameterName;
                if (parameterType.Contains("="))
                {
                    parameter.Type = parameterType.Replace(" ", "").Split('=')[0].Replace("\"", "");
                    parameter.Default = parameterType.Replace(" ", "").Split('=')[1].Replace("\"", "");
                }
                else
                {
                    parameter.Type = parameterType;
                    parameter.Default = "";
                }

                if (!(string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(parameterType)))
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

        public string SearchContent(List<string> content, List<string> patterns)
        {
            var hasSearch = true;
            foreach (var perLine in content)
            {
                foreach (var pattern in patterns)
                {
                    if (!perLine.ToLower().Contains(pattern.ToLower()))
                    {
                        hasSearch = false;
                        break;
                    }

                    hasSearch = true;
                }

                if (hasSearch)
                    return perLine;
            }

            return "";
        }

        private string GetFunctionName(string line, ref string type)
        {
            if (Regex.IsMatch(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase))
            {
                type = "Function";
                return Regex.Match(line, @"Function\s(?<func>\w+)\(", RegexOptions.IgnoreCase).Groups["func"]
                    .ToString();
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
            //Merge fileName1 into fileName2
            if (File.Exists(fileName1) && File.Exists(fileName1))
            {
                var proCFile1 = GetProcOfLine(fileName1);
                var proCFile2 = GetProcOfLine(fileName2);
                var skipLineList = new List<Procedure>();
                foreach (var line in proCFile2)
                    if (proCFile1.Exists(x => x.Name.Equals(line.Name, StringComparison.OrdinalIgnoreCase)))
                        skipLineList.Add(line);
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

                if (newFileName != null) File.Move(fileName2, newFileName);
            }
        }

        private List<string> ReadBasContent(string fileName, List<Procedure> skipLineList = null)
        {
            var list = new List<string>();
            var cnt = 0;
            using (var sw = new StreamReader(fileName))
            {
                string line;
                while ((line = sw.ReadLine()) != null)
                {
                    cnt++;
                    var isSkip = false;
                    if (skipLineList != null)
                        foreach (var skipLine in skipLineList)
                            if (cnt >= skipLine.Start && cnt < skipLine.End)
                            {
                                isSkip = true;
                                break;
                            }

                    if (!isSkip)
                        list.Add(line);
                }
            }

            return list;
        }

        public void AddComment(string filePath)
        {
            if (File.Exists(filePath))
            {
                var lines = ReadBasContent(filePath);
                var index = lines.FindLastIndex(x =>
                    x.StartsWith("Attribute VB_Name = \"", StringComparison.OrdinalIgnoreCase));
                var newLines = new List<string>();
                foreach (var line in lines)
                    if (line.StartsWith("'"))
                    {
                        if (line.Contains("MD5=")) continue;
                        if (line.Contains("AutoGen-Version")) continue;
                        if (line.Contains("'Test Plan:")) continue;
                        if (line.Contains("'SCGH:Skip SCGH file")) continue;
                        if (line.Contains("'Pattern List:Skip Pattern List Csv")) continue;
                        if (line.Contains("SettingFolder:")) continue;
                        if (line.Contains("VBT is not using Central")) continue;
                        newLines.Add(line);
                    }
                    else
                    {
                        newLines.Add(line);
                    }

                if (_srcInfoRows != null && index != -1)
                    newLines.InsertRange(index + 1,
                        _srcInfoRows.Select(x => "'" + Combine.CombineString(x.InputFile, x.Comment, " : ")));

                File.WriteAllLines(filePath, newLines);
            }
        }
    }

    public class Parameter
    {
        public string Default;
        public string Name;
        public string Type;
    }

    public class Procedure
    {
        public int End;
        public string Name;
        public List<Parameter> Parameters;
        public int Start;
        public string Type;
    }
}