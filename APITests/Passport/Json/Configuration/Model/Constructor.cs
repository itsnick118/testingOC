namespace APITests.Passport.Json.Configuration.Model
{
    public class Constructor
    {
        // [{""name"":""expire"",""options"":{""expireTime"":15000}},{""name"":""favorite""},{""name"":""update""}]

        public string Name { get; set; }
        public ConstructorOptions Options { get; set; }
    }
}
