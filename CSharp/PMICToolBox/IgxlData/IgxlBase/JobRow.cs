namespace IgxlData.IgxlBase
{
    public class JobRow : IgxlRow
    {
        public string LineNum { get; set; }
        public string JobName { get; set; }
        public string PinMap { get; set; }
        public string TestInstances { get; set; }
        public string FlowTable { get; set; }
        public string AcSpecs { get; set; }
        public string DcSpecs { get; set; }
        public string PatternSets { get; set; }
        public string PatternGroups { get; set; }
        public string BinTable { get; set; }
        public string Characterization { get; set; }
        public string TestProcedures { get; set; }
        public string MixedSignalTiming { get; set; }
        public string WaveDefinitions { get; set; }
        public string Psets { get; set; }
        public string Signals { get; set; }
        public string PortMap { get; set; }
        public string FractionalBus { get; set; }
        public string ConcurrentSequence { get; set; }
        public string SpikeCheckConfig { get; set; }
        public string Comment { get; set; }

        public JobRow()
        {
            LineNum = "";
            JobName = "";
            PinMap = "";
            TestInstances = "";
            FlowTable = "";
            AcSpecs = "";
            DcSpecs = "";
            PatternSets = "";
            PatternGroups = "";
            BinTable = "";
            Characterization = "";
            TestProcedures = "";
            MixedSignalTiming = "";
            WaveDefinitions = "";
            Psets = "";
            Signals = "";
            PortMap = "";
            FractionalBus = "";
            ConcurrentSequence = "";
            Comment = "";
        }
    }
}
