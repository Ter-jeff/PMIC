namespace FWFrame.nWireDefinition.InputModel
{
    public class Port
    {
        public string Name { get; set; }
        public string Group { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }

        public Port()
        {
            Name = string.Empty;
            Group = string.Empty;
            Type = string.Empty;
            Description = string.Empty;
        }
    }
}
