using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{SetupName}")]
    [Serializable]
    public class CharSetup : IgxlRow
    {
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

        public string SetupName { set; get; }
        public string TestMethod { set; get; }
        public List<CharStep> CharSteps { set; get; }

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
    }
}