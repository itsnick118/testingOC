namespace APITests.Passport.Json.Configuration.Model
{
    public class CalculatedColumn
    {
        public ServiceMethod Function { get; set; }
        public string[] InputFields { get; set; }
        public dynamic Name { get; set; }
        public ServiceMethod Value { get; set; }
    }
}
