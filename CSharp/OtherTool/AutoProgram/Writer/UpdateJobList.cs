using IgxlData.IgxlSheets;
using System.Linq;

namespace AutoProgram.Writer
{
    public class UpdateJobList
    {
        public JobListSheet Work(JobListSheet jobListSheet, string patSetsAllCz,
            string patSetCz, string instanceSheet, string charSheet)
        {
            foreach (var row in jobListSheet.Rows)
            {
                var testInstances = row.TestInstance.Split(',').ToList();
                testInstances.Add(instanceSheet);
                row.TestInstance = string.Join(",", testInstances.Where(x => !string.IsNullOrEmpty(x)));
                var patternSets = row.PatternSets.Split(',').ToList();
                patternSets.Add(patSetsAllCz);
                patternSets.Add(patSetCz);
                row.PatternSets = string.Join(",", patternSets.Where(x => !string.IsNullOrEmpty(x)));
                var characterizations = row.Characterization.Split(',').ToList();
                characterizations.Add(charSheet);
                row.Characterization = string.Join(",", characterizations.Where(x => !string.IsNullOrEmpty(x)));
            }

            return jobListSheet;
        }
    }
}