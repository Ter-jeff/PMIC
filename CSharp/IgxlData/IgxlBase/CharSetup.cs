using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace IgxlData.IgxlBase
{
    public class CharSetupConst
    {
        public const string TestMethodReBurst = "Reburst";
        public const string TestMethodRetest = "Retest";
        public const string TestMethodReBurstSerial = "Reburst Serial";
        public const string TestMethodRunFunction = "Run Function";
        public const string TestMethodRunPattern = "Run Pattern";
        public static readonly Dictionary<string, string> TestMethod = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "Reburst", "Reburst" }, { "Retest", "Retest" }, { "ReburstSerial", "Reburst Serial" }, { "RunFunction", "Run Function" }, { "RunPattern", "Run Pattern" } };
    }

    [Serializable]
    public class CharSetup:IgxlItem
    {
        #region Property
        public string SetupName { set; get; }
        public string TestMethod { set; get; }
        public List<CharStep> CharSteps { set; get; }
        #endregion

        #region Constructor
        public CharSetup()
        {
            SetupName = "";
            TestMethod = "";
            CharSteps = new List<CharStep>();
        }

        public void AddStep(CharStep step)
        {
            CharSteps.Add(step);
        }
        #endregion

        public List<CharStep> DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, CharSteps);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as List<CharStep>;
            }
        }
    }
}