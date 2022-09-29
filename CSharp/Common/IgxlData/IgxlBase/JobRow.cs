namespace IgxlData.IgxlBase
{
    public class JobRow : IgxlRow
    {
        #region Property

        public string LineNum { get; set; }
        public string JobName { get; set; }
        public string PinMap { get; set; }
        public string TestInstance { get; set; }
        public string FlowTable { get; set; }
        public string AcSpecs { get; set; }
        public string DcSpecs { get; set; }
        public string PatternSets { get; set; }
        public string PatternGroups { get; set; }
        public string BinTable { get; set; }
        public string Characterization { get; set; }
        public string TestProcedures { get; set; }
        public string MixedSignalTiming { get; set; }
        public string WaveDefinition { get; set; }
        public string PSets { get; set; }
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
            TestInstance = "";
            FlowTable = "";
            AcSpecs = "";
            DcSpecs = "";
            PatternSets = "";
            PatternGroups = "";
            BinTable = "";
            Characterization = "";
            TestProcedures = "";
            MixedSignalTiming = "";
            WaveDefinition = "";
            PSets = "";
            Signals = "";
            PortMap = "";
            FractionalBus = "";
            ConcurrentSequence = "";
            SpikeCheckConfig = "";
            Comment = "";
        }

        #endregion
    }
}