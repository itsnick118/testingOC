namespace APITests.Passport.Json.Configuration.Model
{
    public class Application
    {
        public ModuleDefinition GlobalDocuments { get; set; }
        public ModuleDefinition Matter { get; set; }
        public string Name { get; set; }
        public ModuleDefinition Spend { get; set; }
        public UserDefinedFields UserDefinedFields { get; set; }
    }
}
