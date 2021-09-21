namespace APITests.Passport.Json.Configuration.Model
{
    public class ListOptions
    {
        public string[] DisplayRegions { get; set; }
        public bool HasSecondaryActions { get; set; }
        public RegionToFieldMappings RegionToFieldMappings { get; set; }
    }
}
