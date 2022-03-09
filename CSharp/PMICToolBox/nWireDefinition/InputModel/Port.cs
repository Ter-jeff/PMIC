namespace nWireDefinition.InputModel
{
    public class Port
    {
        public string Name { get; set; }
        public string Group { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }

        public Port()
        {
            this.Name = string.Empty;
            this.Group = string.Empty;
            this.Type = string.Empty;
            this.Description = string.Empty;
        }
    }
}
