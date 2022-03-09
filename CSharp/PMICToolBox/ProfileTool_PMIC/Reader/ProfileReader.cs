using ProfileTool_PMIC.Output;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProfileTool_PMIC.Reader
{
    public class ProfileReader
    {
        public List<Profile> Read(string profilePath, bool flag, List<string> profileFiles2)
        {
            var profileFiles = Directory.GetFiles(profilePath, "Profile-*.txt", SearchOption.AllDirectories);
            var profileItems = new List<Profile>();
            foreach (var profileFile in profileFiles)
            {
                var profileItem = ReadProfileFile(profileFile, flag, profileFiles2);

                profileItems.Add(profileItem);
            }
            return profileItems;
        }

        public Profile ReadProfileFile(string profile, bool flag, List<string> profileFiles2)
        {
            var profileItem = new Profile();
            var filename = Path.GetFileNameWithoutExtension(profile);
            var arr = filename.Split('-');

            profileItem.FilePath = profile;
            profileItem.Site = int.Parse(Regex.Replace(arr[1], "Site", "", RegexOptions.IgnoreCase));
            profileItem.Pin = arr[2];
            profileItem.SampleRate = double.Parse(arr[3]);
            profileItem.SampleSize = double.Parse(arr[4]);
            if (arr.Count() >= 7)
            {
                profileItem.Item = arr[5];
                profileItem.Date = arr[6];
            }

            var isPower = false;
            var allLines = new List<double>();
            var key = string.Join("-", arr.ToList().GetRange(1, arr.Count() - 2).ToArray());
            if (flag && profileFiles2.Exists(x => x.ToUpper().Contains(key.ToUpper())))
            {
                var profile2 = profileFiles2.Find(x => x.ToUpper().Contains(key.ToUpper()));
                var allLines2 = File.ReadAllLines(profile2).Select(double.Parse).ToList();
                var allLines1 = File.ReadAllLines(profile).Select(double.Parse).ToList();
                if (allLines1.Count() == allLines2.Count())
                {
                    for (var index = 0; index < allLines1.Count; index++)
                        allLines.Add(allLines1[index] * allLines2[index]);
                    isPower = true;
                }
            }
            else
                allLines = File.ReadAllLines(profile).Select(double.Parse).ToList();
            if (allLines.Any())
            {
                profileItem.Value = allLines;
                profileItem.MaxBeforeFilter = profileItem.Value.Max();
                profileItem.MinBeforeFilter = profileItem.Value.Min();
                profileItem.CountBeforeFilter = profileItem.Value.Count;
                profileItem.MaxAfterFilter = profileItem.MaxBeforeFilter;
                profileItem.MaxIndex = profileItem.Value.IndexOf(profileItem.MaxBeforeFilter);
                profileItem.MinAfterFilter = profileItem.MinBeforeFilter;
                profileItem.CountAfterFilter = profileItem.CountBeforeFilter;
                if (isPower)
                profileItem.ChartType = "Power";
                else if (filename.StartsWith("Voltage", StringComparison.CurrentCultureIgnoreCase))
                    profileItem.ChartType = "Voltage(V)";
                else if (filename.StartsWith("Current", StringComparison.CurrentCultureIgnoreCase))
                    profileItem.ChartType = "Current(A)";
            }
            return profileItem;
        }

        public Profile ReadprofileFileWithoutValue(string profileFile)
        {
            var profileItem = new Profile();
            var arr = Path.GetFileNameWithoutExtension(profileFile).Split('-');
            profileItem.FilePath = profileFile;
            profileItem.Site = int.Parse(Regex.Replace(arr[1], "Site", "", RegexOptions.IgnoreCase));
            profileItem.Pin = arr[2];
            profileItem.SampleRate = double.Parse(arr[3]);
            profileItem.SampleSize = double.Parse(arr[4]);
            if (arr.Count() >= 7)
            {
                profileItem.Item = arr[5];
                profileItem.Date = arr[6];
            }
            return profileItem;
        }

        public bool Filter(Profile profileFile, double pulseWidth, int stdevSpec)
        {
            const int countRange = 50;
            if (profileFile.Value.Count == 0)
                return false;
            var max = profileFile.Value.Max();

            var index = profileFile.Value.IndexOf(max);
            var start = index - countRange < 0 ? 0 : index - countRange;
            var count = start + 2 * countRange + 1 >= profileFile.Value.Count ? profileFile.Value.Count - start : 2 * countRange + 1;
            var getRange = profileFile.Value.GetRange(start, count);
            var standardDeviation = getRange.GetStandardDeviation();
            var average = getRange.Average();
            var filteRange = (int)(profileFile.SampleRate * pulseWidth);
            var startFilter = index - filteRange < 0 ? 0 : index - filteRange;
            var endFilter = index + filteRange >= profileFile.Value.Count ? 0 : index + filteRange;
            for (var i = endFilter; i > startFilter; i--)
            {
                if (profileFile.Value[i] > average + stdevSpec * standardDeviation)
                {
                    profileFile.Value.RemoveRange(i, 1);
                    profileFile.MaxAfterFilter = profileFile.Value.Max();
                    profileFile.MinAfterFilter = profileFile.Value.Min();
                    profileFile.CountAfterFilter = profileFile.Value.Count;
                    return true;
                }
            }
            return false;
        }

        public void ExportProfileFiles(List<Profile> profileFiles, string profilePath, string tempPath)
        {
            foreach (var profileFile in profileFiles)
            {
                var file = profileFile.FilePath.Replace(profilePath, tempPath);
                var path = Path.GetDirectoryName(file);
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                using (var sw = new StreamWriter(file, false))
                {
                    foreach (var value in profileFile.Value)
                        sw.WriteLine("{0:e15}", value);
                }
            }
        }

    }
}