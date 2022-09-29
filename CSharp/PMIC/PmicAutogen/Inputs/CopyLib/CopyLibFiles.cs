using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutogen.Inputs.CopyLib
{
    public class CopyLibFiles
    {
        public void Work()
        {
            var targetDir = FolderStructure.DirModulesLibTer;
            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            var modulesLibMap = new Dictionary<string, string>();
            modulesLibMap.Add(FolderStructure.StrLibCommon, FolderStructure.DirLibCommon);
            modulesLibMap.Add(FolderStructure.StrLibDc, FolderStructure.DirLibDc);
            modulesLibMap.Add(FolderStructure.StrLibDigital, FolderStructure.DirLibDigital);
            modulesLibMap.Add(FolderStructure.StrLibDsp, FolderStructure.DirLibDsp);
            modulesLibMap.Add(FolderStructure.StrLimitSheet, FolderStructure.DirLibLimitSheet);
            modulesLibMap.Add(FolderStructure.StrLibLimitSheet, FolderStructure.StrLibLimitSheet);
            modulesLibMap.Add(FolderStructure.StrLibRelay, FolderStructure.DirLibRelay);
            modulesLibMap.Add(FolderStructure.StrLibVbtPopgen, FolderStructure.DirLibVbtPopgen);

            var dirInfo = new DirectoryInfo(LocalSpecs.BasLibraryPath);
            var subDirs = dirInfo.GetDirectories().ToList();

            subDirs.ForEach(subDir =>
            {
                if (modulesLibMap.ContainsKey(subDir.Name))
                {
                    var targetFolder = modulesLibMap[subDir.Name];
                    CopyDirectory(subDir.FullName, targetFolder);
                }
            });

            //var dirInfo = new DirectoryInfo(LocalSpecs.BasLibraryPath);
            //var subDirs = dirInfo.GetDirectories().ToList();

            //subDirs.ForEach(subDir =>
            //{
            //    if (FolderStructure.ModulesLibMap.ContainsKey(subDir.Name))
            //    {
            //        string targetFolder = FolderStructure.ModulesLibMap[subDir.Name];
            //        CopyDirectory(subDir.FullName, targetFolder);
            //    }
            //});

            //CreateDefaultEmptyFolders();
            //ClassifyLibFilesByRule();
            FileStructurePostAction();
        }

        private void CopyDirectory(string sourceDir, string targetDir)
        {
            try
            {
                if (!Directory.Exists(targetDir))
                    Directory.CreateDirectory(targetDir);

                var files = Directory.GetFiles(sourceDir);
                foreach (var file in files)
                {
                    var filePath = Path.Combine(targetDir, Path.GetFileName(file));
                    File.Copy(file, filePath, true);
                }

                var subDirectorys = Directory.GetDirectories(sourceDir);
                foreach (var subDirectory in subDirectorys)
                    CopyDirectory(subDirectory, Path.Combine(targetDir, Path.GetFileName(subDirectory)));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //private void CreateDefaultEmptyFolders()
        //{
        //    var otherDir = FolderStructure.DirOtherWaitForClassify;
        //    if (!Directory.Exists(otherDir))
        //        Directory.CreateDirectory(otherDir);

        //    var spikecheckDir = FolderStructure.DirSpikeCheck;
        //    if (!Directory.Exists(spikecheckDir))
        //        Directory.CreateDirectory(spikecheckDir);

        //    var vbtDir = FolderStructure.DirAcore;
        //    if (!Directory.Exists(vbtDir))
        //        Directory.CreateDirectory(vbtDir);

        //    var bstsqDir = FolderStructure.DirBSTSQ;
        //    if (!Directory.Exists(bstsqDir))
        //        Directory.CreateDirectory(bstsqDir);

        //    var buckMutiPhaseDir = FolderStructure.DirBUCKMUTIPHASE;
        //    if (!Directory.Exists(buckMutiPhaseDir))
        //        Directory.CreateDirectory(buckMutiPhaseDir);

        //    var buck1pDir = FolderStructure.DirBUCK1P;
        //    if (!Directory.Exists(buck1pDir))
        //        Directory.CreateDirectory(buck1pDir);

        //    var buck1phDir = FolderStructure.DirBUCK1PH;
        //    if (!Directory.Exists(buck1phDir))
        //        Directory.CreateDirectory(buck1phDir);

        //    var buckswDir = FolderStructure.DirBUCKSW;
        //    if (!Directory.Exists(buckswDir))
        //        Directory.CreateDirectory(buckswDir);

        //    var ldoDir = FolderStructure.DirLdo;
        //    if (!Directory.Exists(ldoDir))
        //        Directory.CreateDirectory(ldoDir);

        //    var libPowerupDir = FolderStructure.DirLibPowerup;
        //    if (!Directory.Exists(libPowerupDir))
        //        Directory.CreateDirectory(libPowerupDir);
        //}


        //private void ClassifyLibFilesByRule()
        //{
        //    var vbtLibs = new DirectoryInfo(FolderStructure.DirModulesLibTer).GetFiles("*", SearchOption.AllDirectories);
        //    foreach (var vbtLib in vbtLibs)
        //    {
        //        if (vbtLib.Extension != ".bas" && vbtLib.Extension != ".cls")
        //            continue;

        //        if (vbtLib.Name.Equals(StrVddLevels, StringComparison.CurrentCultureIgnoreCase))
        //        {
        //            File.Move(vbtLib.FullName, Path.Combine(FolderStructure.DirLibPowerup, vbtLib.Name));
        //            //File.Move(vbtLib.FullName, Path.Combine(FolderStructure.DirCommonSheets, vbtLib.Name));
        //        }
        //    }

        //    var Libs = new DirectoryInfo(FolderStructure.DirOtherWaitForClassify).GetFiles();
        //    foreach (var vbtLib in Libs)
        //    {
        //        if (vbtLib.Extension != ".bas" && vbtLib.Extension != ".cls")
        //            continue;

        //        if (vbtLib.Name.ToUpperInvariant().Contains(OtpPattern) ||
        //            vbtLib.Name.Equals(StrAHBField, StringComparison.CurrentCultureIgnoreCase))
        //        {
        //            File.Move(vbtLib.FullName, Path.Combine(FolderStructure.DirOtp, vbtLib.Name));
        //        }
        //    }
        //}

        private void FileStructurePostAction()
        {
            //var fileNameXls = Path.Combine(FolderStructure.DirOtp, PmicConst.OtpRegisterMap + ".xlsx");
            //var fileNameCsv = Path.Combine(FolderStructure.DirOtp, PmicConst.OtpRegisterMap + ".csv");
            //if (File.Exists(fileNameXls))
            //    File.Delete(fileNameXls);
            //if (File.Exists(fileNameCsv))
            //    File.Delete(fileNameCsv);

            var otherFiles = new DirectoryInfo(FolderStructure.DirOtherWaitForClassify).GetFiles();
            var commonSheetsFiles = new DirectoryInfo(FolderStructure.DirModulesCommonSheets)
                .GetFiles("*", SearchOption.AllDirectories).ToList();
            var blockFiles = new DirectoryInfo(FolderStructure.DirModulesBlock)
                .GetFiles("*", SearchOption.AllDirectories).ToList();
            var libFiles = new DirectoryInfo(FolderStructure.DirModulesLibTer)
                .GetFiles("*", SearchOption.AllDirectories).ToList();

            var moduleFileInfos = new List<FileInfo>();
            moduleFileInfos.AddRange(commonSheetsFiles);
            moduleFileInfos.AddRange(blockFiles);
            moduleFileInfos.AddRange(libFiles);

            foreach (var otherFile in otherFiles)
            {
                var fileName = otherFile.Name;
                if (moduleFileInfos.Find(file =>
                        file.Name.Equals(fileName, StringComparison.CurrentCultureIgnoreCase)) !=
                    null) File.Delete(otherFile.FullName);
            }

            //var scoreFiles = new DirectoryInfo(FolderStructure.DirAcore).GetFiles("*", SearchOption.AllDirectories).ToList();
            //foreach (var scoreFile in scoreFiles)
            //{
            //    var fileName = scoreFile.Name;
            //    if (!fileName.ToUpperInvariant().Contains("ACORE"))
            //    {
            //        File.Move(scoreFile.FullName, Path.Combine(FolderStructure.DirOtherWaitForClassify, fileName));
            //    }
            //}
        }
    }
}